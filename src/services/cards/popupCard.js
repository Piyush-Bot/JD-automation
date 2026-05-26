const { CardFactory } = require('botbuilder');

function buildPopupCard(options = {}) {
    const title = options.title || 'Confirmation';
    const bodyTextLines = Array.isArray(options.bodyTextLines) ? options.bodyTextLines : ['Please confirm your action.'];
    const details = options.details || null; // { headerIntro, department: { name, status }, role: { name, status } }
    const footerHint = options.footerHint || null; // optional subtle footer hint
    const okData = options.okData || { action: 'popup_ok' };
    const cancelData = options.cancelData || { action: 'popup_cancel' };
    const proceedTitle = options.proceedTitle || 'Proceed';
    const cancelTitle = options.cancelTitle || 'Cancel';

    const statusToText = (status) => {
        switch (status) {
            case 'exact': return 'Found in GPM_DB';
            case 'closest': return 'Closest found in GPM_DB';
            case 'not_found': return 'Not found in GPM_DB';
            default: return '';
        }
    };
    const statusToColor = (status) => {
        switch (status) {
            case 'exact': return 'good';
            case 'closest': return 'warning';
            case 'not_found': return 'attention';
            default: return 'default';
        }
    };

    const baseHeader = { type: 'TextBlock', text: title, weight: 'Bolder', size: 'Medium', wrap: true };

    let bodyItems = [];
    if (details && (details.department || details.role)) {
        const headerIntro = details.headerIntro || '';
        const dept = details.department || {};
        const role = details.role || {};
        bodyItems = [
            baseHeader,
            ...(headerIntro ? [{ type: 'TextBlock', text: headerIntro, wrap: true, isSubtle: true, spacing: 'Small' }] : []),
            {
                type: 'ColumnSet',
                spacing: 'Medium',
                columns: [
                    { type: 'Column', width: 'auto', items: [{ type: 'TextBlock', text: 'Department:', weight: 'Bolder', wrap: true }] },
                    {
                        type: 'Column',
                        width: 'stretch',
                        items: [{
                            type: 'RichTextBlock',
                            wrap: true,
                            inlines: [
                                { type: 'TextRun', text: dept.name || '—' },
                                ...(dept.status ? [
                                    { type: 'TextRun', text: ' — ', isSubtle: true },
                                    { type: 'TextRun', text: statusToText(dept.status), color: statusToColor(dept.status) }
                                ] : [])
                            ]
                        }]
                    }
                ]
            },
            {
                type: 'ColumnSet',
                spacing: 'Medium',
                columns: [
                    { type: 'Column', width: 'auto', items: [{ type: 'TextBlock', text: 'Role:', weight: 'Bolder', wrap: true }] },
                    {
                        type: 'Column',
                        width: 'stretch',
                        items: [{
                            type: 'RichTextBlock',
                            wrap: true,
                            inlines: [
                                { type: 'TextRun', text: role.name || '—' },
                                ...(role.status ? [
                                    { type: 'TextRun', text: ' — ', isSubtle: true },
                                    { type: 'TextRun', text: statusToText(role.status), color: statusToColor(role.status) }
                                ] : [])
                            ]
                        }]
                    }
                ]
            },
            { type: 'Container', separator: true, spacing: 'Medium', items: footerHint ? [{ type: 'TextBlock', text: footerHint, isSubtle: true, wrap: true }] : [] }
        ];
    } else {
        bodyItems = [
            baseHeader,
            ...bodyTextLines.map((line) => ({ type: 'TextBlock', text: String(line), wrap: true, spacing: 'Medium' }))
        ];
    }

    const card = {
        type: 'AdaptiveCard',
        $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
        version: '1.4',
        body: [
            { type: 'Container', items: bodyItems },
            {
                type: 'ActionSet',
                spacing: 'Medium',
                actions: [
                    { type: 'Action.Submit', title: proceedTitle, style: 'positive', data: { ...okData, action: 'popup_ok' } },
                    { type: 'Action.Submit', title: cancelTitle, data: { ...cancelData, action: 'popup_cancel' } }
                ]
            }
        ]
    };
    return CardFactory.adaptiveCard(card);
}

function buildTaskModuleResponse(options = {}) {
    const title = options.title || 'Confirmation';
    return {
        task: {
            type: 'continue',
            value: {
                title,
                width: 500,
                card: buildPopupCard(options)
            }
        }
    };
}

module.exports = { buildPopupCard, buildTaskModuleResponse };
