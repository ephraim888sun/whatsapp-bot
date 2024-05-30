require('dotenv').config()

const express = require('express')
const bodyParser = require('body-parser')
const { BotFrameworkAdapter, MemoryStorage, UserState } = require('botbuilder')
const { DialogSet, WaterfallDialog, TextPrompt } = require('botbuilder-dialogs')
const twilio = require('twilio')
const { Client } = require('@microsoft/microsoft-graph-client')
require('isomorphic-fetch')

const app = express()
app.use(bodyParser.json())

const twilioClient = twilio(
  process.env.TWILIO_ACCOUNT_SID,
  process.env.TWILIO_AUTH_TOKEN
)

// Bot Framework Adapter
const adapter = new BotFrameworkAdapter({
  appId: process.env.MICROSOFT_APP_ID,
  appPassword: process.env.MICROSOFT_APP_PASSWORD,
})

// Memory Storage
const memoryStorage = new MemoryStorage()
const userState = new UserState(memoryStorage)
const dialogState = userState.createProperty('dialogState')
const dialogs = new DialogSet(dialogState)

// Add dialogs
dialogs.add(new TextPrompt('textPrompt'))
dialogs.add(
  new WaterfallDialog('mainDialog', [
    async (step) => {
      return await step.prompt('textPrompt', 'Please provide your name:')
    },
    async (step) => {
      step.values.name = step.result
      return await step.prompt(
        'textPrompt',
        'Please provide the date and time for the appointment:'
      )
    },
    async (step) => {
      step.values.datetime = step.result
      await createOutlookEvent(step.values.name, step.values.datetime)
      await step.context.sendActivity('Your appointment has been scheduled.')
      return await step.endDialog()
    },
  ])
)

async function createOutlookEvent(name, datetime) {
  const client = Client.init({
    authProvider: (done) => {
      done(null, process.env.MS_GRAPH_ACCESS_TOKEN)
    },
  })

  const event = {
    subject: `Appointment with ${name}`,
    start: {
      dateTime: datetime,
      timeZone: 'Your/TimeZone',
    },
    end: {
      dateTime: datetime, // Add duration as needed
      timeZone: 'Your/TimeZone',
    },
    attendees: [
      {
        emailAddress: {
          address: 'doctor@example.com',
          name: 'Dr. Smith',
        },
        type: 'required',
      },
    ],
  }

  await client.api('/me/events').post(event)
}

app.post('/api/messages', (req, res) => {
  adapter.processActivity(req, res, async (context) => {
    const dc = await dialogs.createContext(context)
    if (context.activity.type === 'message') {
      await dc.continueDialog()
      if (!context.responded) {
        await dc.beginDialog('mainDialog')
      }
    }
  })
})

app.post('/twilio', (req, res) => {
  const message = req.body.Body
  const activity = {
    type: 'message',
    from: { id: req.body.From },
    recipient: { id: req.body.To },
    text: message,
  }

  adapter.processActivity(activity, null, async (context) => {
    const dc = await dialogs.createContext(context)
    await dc.continueDialog()
    if (!context.responded) {
      await dc.beginDialog('mainDialog')
    }
  })

  const twiml = new twilio.twiml.MessagingResponse()
  twiml.message('Processing your request.')
  res.writeHead(200, { 'Content-Type': 'text/xml' })
  res.end(twiml.toString())
})

app.listen(3000, () => {
  console.log('Server is running on port 3000')
})


