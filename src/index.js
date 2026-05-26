const path = require('path');
const dotenv = require('dotenv');
const restify = require('restify');
const {
    CloudAdapter,
    ConfigurationBotFrameworkAuthentication
} = require('botbuilder');

// Load env if present
const ENV_FILE = path.join(__dirname, '..', '.env');
dotenv.config({ path: ENV_FILE });

const { EchoBot } = require('./bot');

// Create HTTP server
const server = restify.createServer();
server.use(restify.plugins.bodyParser());

const port = process.env.port || process.env.PORT || 3978;
server.listen(port, () => {});

// Bot Framework credentials
const botFrameworkAuthConfig = {
    MicrosoftAppId: process.env.MicrosoftAppId,
    MicrosoftAppPassword: process.env.MicrosoftAppPassword,
    MicrosoftAppTenantId: process.env.MicrosoftAppTenantId,
    MicrosoftAppType: process.env.MicrosoftAppType
};

const botFrameworkAuthentication = new ConfigurationBotFrameworkAuthentication(botFrameworkAuthConfig);
const adapter = new CloudAdapter(botFrameworkAuthentication);

// Catch-all for errors.
const onTurnErrorHandler = async (context, error) => {
    await context.sendActivity('The bot encountered an error or bug.');
};
adapter.onTurnError = onTurnErrorHandler;

// Create bot
const bot = new EchoBot();

// Messages endpoint
server.post('/api/messages', async (req, res) => {
    const incomingAuth = req.headers.authorization;
    await adapter.process(req, res, async (context) => {
        if (incomingAuth) {
            context.turnState.set('msAuthHeader', incomingAuth);
        }
        await bot.run(context);
    });
});

// Streaming (optional)
server.on('upgrade', async (req, socket, head) => {
    const streamingAdapter = new CloudAdapter(botFrameworkAuthentication);
    streamingAdapter.onTurnError = onTurnErrorHandler;
    await streamingAdapter.process(req, socket, head, (context) => bot.run(context));
});
