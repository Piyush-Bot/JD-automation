const { CardFactory } = require('botbuilder');

function buildEditFormCard(ctx = {}) {
    const card = {
        type: 'AdaptiveCard',
        $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
        version: '1.4',
        msteams: { width: 'Full' },
        body: [
            {
                type: 'ColumnSet',
                columns: [
                    {
                        type: 'Column',
                        width: 'stretch',
                        items: [{ type: 'TextBlock', text: 'Edit JD', weight: 'Bolder', size: 'Medium', wrap: true }]
                    },
                    {
                        type: 'Column',
                        width: 'auto',
                        items: [{
                            type: 'ActionSet',
                            actions: [{ type: 'Action.Submit', title: '✕', data: { action: 'card_close' } }]
                        }]
                    }
                ]
            },
            { type: 'TextBlock', text: 'Edit Type:', wrap: true, size: 'Small', weight: 'Bolder', spacing: 'Medium' },
            {
                type: 'Input.ChoiceSet',
                id: 'editTypes',
                isMultiSelect: true,
                placeholder: '-- Select edit type(s) --',
                choices: [
                    { title: 'Skills', value: 'skills' },
                    { title: 'Experience', value: 'experience' },
                    { title: 'Education', value: 'education' },
                    { title: 'Comp Range', value: 'comp_range' },
                    { title: 'Role & Responsibility', value: 'roles_and_respo' }
                ]
            },
            { type: 'TextBlock', text: 'Description:', wrap: true, size: 'Small', weight: 'Bolder', spacing: 'Medium' },
            { type: 'Input.Text', id: 'description', placeholder: 'Enter description...', isMultiline: true },
            { type: 'Input.Text', id: 'role', value: ctx.role || '', isVisible: false },
            { type: 'Input.Text', id: 'department', value: ctx.department || '', isVisible: false },
            { type: 'Input.Text', id: 'rawOutput', value: ctx.rawOutput || '', isVisible: false },
            { type: 'Input.Text', id: 'jdId', value: ctx.jdId || '', isVisible: false }
        ],
        actions: [
            { type: 'Action.Submit', title: '✅ Submit', data: { action: 'edit_form_submit' } }
        ]
    };
    return CardFactory.adaptiveCard(card);
}

module.exports = { buildEditFormCard };
