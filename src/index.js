const path = require('path');
const restify = require('restify');
const { ConfigurationServiceClientCredentialFactory } = require('botbuilder');
const { TeamsAdapter } = require('@microsoft/teams-ai');
const { app } = require('./searchApp');
const config = require("./config");

// 建立 Adapter
const adapter = new TeamsAdapter(
    {},
    new ConfigurationServiceClientCredentialFactory({
        MicrosoftAppId: config.botId,
        MicrosoftAppPassword: config.botPassword,
        MicrosoftAppType: 'MultiTenant'
    })
);

// 錯誤處理
adapter.onTurnError = async (context, error) => {
    console.error(`\n [onTurnError] unhandled error: ${error.toString()}`);
    await context.sendActivity('The bot encountered an error or bug.');
    await context.sendActivity('To continue to run this bot, please fix the bot source code.');
};

// 建立 HTTP 伺服器
const server = restify.createServer();
server.use(restify.plugins.bodyParser());

server.listen(process.env.PORT || 3978, () => {
  console.log(`\nBot Started, ${server.name} listening to ${server.url}`);
});

// 處理訊息請求
server.post('/api/messages', async (req, res) => {
    await adapter.process(req, res, async (context) => {
        await app.run(context);
    });
});

// 提供靜態 HTML 頁面（用於 OAuth）
server.get(
    '/auth-:name(start|end).html',
    restify.plugins.serveStatic({
        directory: path.join(__dirname, 'public')
    })
);
