const { CardFactory } = require('botbuilder');

function buildJdCreatFormCard(departments, roles, members) {
    const deptChoices = (departments || []).map((d) => ({ title: d.name, value: String(d.id) }));
    const roleChoices = (roles || []).map((r) => ({ title: r.name, value: String(r.id) }));
    const memberChoices = (members || []).map((m) => ({ title: m.name, value: String(m.id) }));

    const card = {
        type: 'AdaptiveCard', $schema: 'http://adaptivecards.io/schemas/adaptive-card.json', version: '1.4', msteams: { width: 'Full' },
        body: [
            {
                type: 'ColumnSet',
                columns: [
                    { type: 'Column', width: 'stretch', items: [{ type: 'TextBlock', text: 'JD Creation', weight: 'Bolder', size: 'Large', wrap: true }] },
                    { type: 'Column', width: 'auto', items: [{ type: 'ActionSet', actions: [{ type: 'Action.Submit', title: '✕', data: { action: 'card_close' } }] }] }
                ]
            },
            { type: 'TextBlock', text: 'Select Department, Role and the workflow participants, then submit.', wrap: true, size: 'Small', spacing: 'Small' },
            {
                type: 'ColumnSet', spacing: 'Medium',
                columns: [
                    { type: 'Column', width: 'stretch', items: [
                        { type: 'TextBlock', text: 'Department:', wrap: true, size: 'Small', weight: 'Bolder' },
                        { type: 'Input.ChoiceSet', id: 'deptId', style: 'compact', placeholder: '-- Choose Department --', choices: deptChoices }
                    ]},
                    { type: 'Column', width: 'stretch', items: [
                        { type: 'TextBlock', text: 'Role:', wrap: true, size: 'Small', weight: 'Bolder' },
                        { type: 'Input.ChoiceSet', id: 'roleId', style: 'compact', placeholder: '-- Choose Role --', choices: roleChoices }
                    ]}
                ]
            },
            {
                type: 'ColumnSet', spacing: 'Medium',
                columns: [
                    { type: 'Column', width: 'stretch', items: [
                        { type: 'TextBlock', text: 'Originator:', wrap: true, size: 'Small', weight: 'Bolder' },
                        { type: 'Input.ChoiceSet', id: 'originatorId', style: 'compact', placeholder: '-- Choose Originator --', choices: memberChoices }
                    ]},
                    { type: 'Column', width: 'stretch', items: [
                        { type: 'TextBlock', text: 'Reviewer:', wrap: true, size: 'Small', weight: 'Bolder' },
                        { type: 'Input.ChoiceSet', id: 'reviewerId', style: 'compact', placeholder: '-- Choose Reviewer --', choices: memberChoices }
                    ]}
                ]
            },
            { type: 'TextBlock', text: 'Approver:', wrap: true, size: 'Small', weight: 'Bolder', spacing: 'Medium' },
            { type: 'Input.ChoiceSet', id: 'approverId', style: 'compact', placeholder: '-- Choose Approver --', choices: memberChoices }
        ],
        actions: [{ type: 'Action.Submit', title: '✅ Submit', data: { action: 'jd_form_submit' } }]
    };

    return CardFactory.adaptiveCard(card);
}

module.exports = { buildJdCreatFormCard };
