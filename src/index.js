const path = require('path');
const restify = require('restify');
const { ConfigurationServiceClientCredentialFactory } = require('botbuilder');
const { TeamsAdapter } = require('@microsoft/teams-ai');
const { app } = require('./searchApp');
const config = require("./config");

const adapter = new TeamsAdapter(
    {},
    new ConfigurationServiceClientCredentialFactory({
        MicrosoftAppId: config.botId,
        MicrosoftAppPassword: config.botPassword,
        MicrosoftAppType: 'MultiTenant'
    })
);

adapter.onTurnError = async (context, error) => {
    console.error(`\n [onTurnError] unhandled error: ${error.toString()}`);
    await context.sendActivity('The bot encountered an error or bug.');
    await context.sendActivity('To continue to run this bot, please fix the bot source code.');
};


const server = restify.createServer();
server.use(restify.plugins.bodyParser());

server.listen(process.env.PORT || 3978, () => {
  console.log(`\nBot Started, ${server.name} listening to ${server.url}`);
});


server.post('/api/messages', async (req, res) => {
    await adapter.process(req, res, async (context) => {
        await app.run(context);
    });
});


server.get(
    '/auth-:name(start|end).html',
    restify.plugins.serveStatic({
        directory: path.join(__dirname, 'public')
    })
);
