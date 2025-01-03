// index.js is used to setup and configure your bot

// Import required packages
const restify = require("restify");

// Import required bot services.
// See https://aka.ms/bot-services to learn more about the different parts of a bot.
const {
  BotFrameworkAdapter,
} = require("botbuilder");
const { SearchApp } = require("./searchApp");
const config = require("./config");


// Create adapter.
// See https://aka.ms/about-bot-adapter to learn more about adapters.

const adapter = new BotFrameworkAdapter({
  appId: config.botId,
  appPassword: config.botPassword
});


adapter.onTurnError = async (context, error) => {
  // This check writes out errors to console log .vs. app insights.
  // NOTE: In production environment, you should consider logging this to Azure
  //       application insights. See https://aka.ms/bottelemetry for telemetry
  //       configuration instructions.
  console.error(`\n [onTurnError] unhandled error: ${error}`);

  // Send a message to the user
  await context.sendActivity(`The bot encountered an unhandled error:\n ${error.message}`);
  await context.sendActivity("To continue to run this bot, please fix the bot source code.");
};

// Create the bot that will handle incoming messages.
const searchApp = new SearchApp();

// Create HTTP server.
const server = restify.createServer();
server.use(restify.plugins.bodyParser());

// Use the assigned port in Azure, or a default for local testing.
const port = process.env.PORT || 3978;
console.log("PORT:", port);
server.listen(port, function () {
  console.log(`\nBot started, ${server.name} listening on port ${port}`);
});

// Listen for incoming requests.
server.post("/api/messages", async (req, res) => {
  await adapter.process(req, res, async (context) => {
    console.log("Context parameter:", context._activity.value.parameters);
    
    await searchApp.run(context);
  });
});

// server.post("/api/messages", async (req, res) => {
//   try {
//     await adapter.process(req, res, async (context) => {
//       console.log("Request received:", req.body);
//       await searchApp.run(context);
//     });
//   } catch (error) {
//     console.error("Error in POST /api/messages:", error);
//     res.send(500);
//   }
// });

// Gracefully shutdown HTTP server
["exit", "uncaughtException", "SIGINT", "SIGTERM", "SIGUSR1", "SIGUSR2"].forEach((event) => {
  process.on(event, () => {
    server.close();
  });
});