const { CardFactory } = require('botbuilder');

function buildMenuCard() {
    const card = {
        type: 'AdaptiveCard',
        $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
        version: '1.4',
        body: [
            { type: 'TextBlock', text: 'JD Process', weight: 'Bolder', size: 'Large', wrap: true },
            { type: 'TextBlock', text: 'Please select an option to proceed:', wrap: true, size: 'Small', spacing: 'Small' }
        ],
        actions: [
            { type: 'Action.Submit', title: '📄 JD Creation', data: { action: 'jd_start', flowSource: 'creation' } },
            { type: 'Action.Submit', title: '📋 JD Edit', data: { action: 'fetch_indent', flowSource: 'fetch' } }
        ]
    };
    return CardFactory.adaptiveCard(card);
}

module.exports = { buildMenuCard };
