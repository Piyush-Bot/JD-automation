const { CardFactory } = require('botbuilder');

function buildJdEditCard(departments, roles) {
    const deptChoices = (departments || []).map((d) => ({ title: d.name, value: String(d.id) }));
    const roleChoices = (roles || []).map((r) => ({ title: r.name, value: String(r.id) }));

    const card = {
        type: 'AdaptiveCard', $schema: 'http://adaptivecards.io/schemas/adaptive-card.json', version: '1.4', msteams: { width: 'Full' },
        body: [
            {
                type: 'ColumnSet',
                columns: [
                    { type: 'Column', width: 'stretch', items: [{ type: 'TextBlock', text: 'Edit JD', weight: 'Bolder', size: 'Large', wrap: true }] },
                    { type: 'Column', width: 'auto', items: [{ type: 'ActionSet', actions: [{ type: 'Action.Submit', title: '✕', data: { action: 'card_close' } }] }] }
                ]
            },
            { type: 'TextBlock', text: 'Choose Department and Role, then submit to fetch filtered JD records.', wrap: true, size: 'Small', spacing: 'Small' },
            { type: 'Input.Text', id: 'flowSource', value: 'fetch', isVisible: false },
            {
                type: 'ColumnSet', spacing: 'Medium',
                columns: [
                    { type: 'Column', width: 'stretch', items: [
                        { type: 'TextBlock', text: 'Department:', wrap: true, size: 'Small', weight: 'Bolder' },
                        { type: 'Input.ChoiceSet', id: 'departmentId', style: 'compact', placeholder: '-- Choose Department --', choices: deptChoices }
                    ]},
                    { type: 'Column', width: 'stretch', items: [
                        { type: 'TextBlock', text: 'Role:', wrap: true, size: 'Small', weight: 'Bolder' },
                        { type: 'Input.ChoiceSet', id: 'roleId', style: 'compact', placeholder: '-- Choose Role --', choices: roleChoices }
                    ]}
                ]
            }
        ],
        actions: [{ type: 'Action.Submit', title: '✅ Submit', data: { action: 'fetch_indent_submit' } }]
    };

    return CardFactory.adaptiveCard(card);
}

module.exports = { buildJdEditCard };
