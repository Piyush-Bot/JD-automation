const { CardFactory } = require('botbuilder');

function formatSectionTitle(key) {
    return key
        .replace(/_/g, ' ')
        .replace(/\b\w/g, (c) => c.toUpperCase());
}

function buildTableSection(title, records) {
    if (!records || records.length === 0) return [];
    const headers = Object.keys(records[0]);
    const items = [];

    items.push({
        type: 'TextBlock',
        text: title,
        weight: 'Bolder',
        size: 'Medium',
        spacing: 'Medium',
        separator: true,
        wrap: true
    });

    items.push({
        type: 'ColumnSet',
        style: 'emphasis',
        columns: headers.map((h) => ({
            type: 'Column',
            width: 'stretch',
            items: [{
                type: 'TextBlock',
                text: h.replace(/_/g, ' '),
                weight: 'Bolder',
                wrap: true,
                size: 'Small'
            }]
        }))
    });

    records.forEach((record) => {
        items.push({
            type: 'ColumnSet',
            columns: headers.map((h) => ({
                type: 'Column',
                width: 'stretch',
                items: [{
                    type: 'TextBlock',
                    text: (record[h] !== undefined && record[h] !== null) ? String(record[h]) : '—',
                    wrap: true,
                    size: 'Small'
                }]
            }))
        });
    });

    return items;
}

function buildStringSection(title, value) {
    return [
        {
            type: 'TextBlock',
            text: title,
            weight: 'Bolder',
            size: 'Medium',
            spacing: 'Medium',
            separator: true,
            wrap: true
        },
        {
            type: 'Container',
            style: 'emphasis',
            spacing: 'Small',
            items: [{ type: 'TextBlock', text: value, wrap: true, size: 'Small' }]
        }
    ];
}

function buildJdResultCard(output, title = '✅ JD Created Successfully', editCtx = {}, acceptCtx = {}, { editEnabled = true, acceptEnabled = true, acceptAction = 'jd_accept' } = {}) {
    const body = [
        {
            type: 'ColumnSet',
            columns: [
                {
                    type: 'Column', width: 'stretch',
                    items: [{ type: 'TextBlock', text: title, weight: 'Bolder', size: 'Large', color: 'Good', wrap: true }]
                },
                {
                    type: 'Column', width: 'auto',
                    items: [{ type: 'ActionSet', actions: [{ type: 'Action.Submit', title: '✕', data: { action: 'card_close' } }] }]
                }
            ]
        }
    ];

    for (const [key, value] of Object.entries(output || {})) {
        const title = formatSectionTitle(key);
        if (Array.isArray(value) && value.length > 0 && typeof value[0] === 'object') {
            body.push(...buildTableSection(title, value));
        } else if (typeof value === 'string' && value) {
            body.push(...buildStringSection(title, value));
        }
    }

    const card = {
        type: 'AdaptiveCard',
        $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
        version: '1.4',
        msteams: { width: 'Full' },
        body,
        actions: [
            { type: 'Action.Submit', title: '✏️ Edit', data: { action: 'jd_edit', role: editCtx.role || '', department: editCtx.department || '', originator: editCtx.originator || '', reviewer: editCtx.reviewer || '', approver: editCtx.approver || '', rawOutput: editCtx.rawOutput ? JSON.stringify(editCtx.rawOutput) : '', jdId: editCtx.jdId || '', flowSource: editCtx.flowSource || '' }, style: 'positive', isEnabled: editEnabled },
            { type: 'Action.Submit', title: '✅ Accept', data: { action: acceptAction, role: acceptCtx.role || '', department: acceptCtx.department || '', originator: acceptCtx.originator || '', reviewer: acceptCtx.reviewer || '', approver: acceptCtx.approver || '', output: acceptCtx.output ? JSON.stringify(acceptCtx.output) : '', jdId: acceptCtx.jdId || '', flowSource: acceptCtx.flowSource || '' }, style: 'positive', isEnabled: acceptEnabled }
        ]
    };

    return CardFactory.adaptiveCard(card);
}

module.exports = { buildJdResultCard };
