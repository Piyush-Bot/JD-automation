const { CardFactory } = require('botbuilder');

function buildMenuCard() {
    const card = {
        type: 'AdaptiveCard',
        $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
        version: '1.4',
        body: [
            { type: 'TextBlock', text: 'JD Process', weight: 'Bolder', size: 'Large', wrap: true },
            { type: 'TextBlock', text: 'Please select an option to proceed:', wrap: true, spacing: 'Small' }
        ],
        actions: [
            { type: 'Action.Submit', title: '📄 JD Creation', data: { action: 'jd_start' } },
            { type: 'Action.Submit', title: '📋 Fetch JD', data: { action: 'fetch_indent' } }
        ]
    };
    return CardFactory.adaptiveCard(card);
}

function buildDeptCard(departments) {
    const choices = departments.map((d) => ({ title: d.name, value: String(d.id) }));
    const card = {
        type: 'AdaptiveCard', $schema: 'http://adaptivecards.io/schemas/adaptive-card.json', version: '1.4',
        body: [
            { type: 'TextBlock', text: 'JD Creation — Step 1', weight: 'Bolder', size: 'Medium', wrap: true },
            { type: 'TextBlock', text: 'Select a Department:', wrap: true, spacing: 'Small' },
            { type: 'Input.ChoiceSet', id: 'deptId', style: 'compact', isRequired: true, errorMessage: 'Please select a department.', placeholder: '-- Choose Department --', choices }
        ],
        actions: [{ type: 'Action.Submit', title: 'Select Department →', data: { action: 'dept_selected' } }]
    };
    return CardFactory.adaptiveCard(card);
}

function buildRoleCard(deptId, deptName, roles) {
    const choices = roles.map((r) => ({ title: r.name, value: String(r.id) }));
    const card = {
        type: 'AdaptiveCard', $schema: 'http://adaptivecards.io/schemas/adaptive-card.json', version: '1.4',
        body: [
            { type: 'TextBlock', text: 'JD Creation — Step 2', weight: 'Bolder', size: 'Medium', wrap: true },
            { type: 'FactSet', facts: [{ title: 'Department:', value: deptName }], spacing: 'Small' },
            { type: 'TextBlock', text: 'Select a Role:', wrap: true, spacing: 'Medium' },
            { type: 'Input.ChoiceSet', id: 'roleId', style: 'compact', isRequired: true, errorMessage: 'Please select a role.', placeholder: '-- Choose Role --', choices },
            { type: 'Input.Text', id: 'deptId', value: String(deptId), isVisible: false },
            { type: 'Input.Text', id: 'deptName', value: String(deptName), isVisible: false }
        ],
        actions: [{ type: 'Action.Submit', title: '✅ Submit', data: { action: 'jd_submit' } }]
    };
    return CardFactory.adaptiveCard(card);
}

function buildJdCreationFormCard(departments, roles, members) {
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
            { type: 'TextBlock', text: 'Select Department, Role and the workflow participants, then submit.', wrap: true, spacing: 'Small' },
            { type: 'TextBlock', text: 'Department:', wrap: true, spacing: 'Medium' },
            { type: 'Input.ChoiceSet', id: 'deptId', style: 'compact', placeholder: '-- Choose Department --', choices: deptChoices },
            { type: 'TextBlock', text: 'Role:', wrap: true, spacing: 'Medium' },
            { type: 'Input.ChoiceSet', id: 'roleId', style: 'compact', placeholder: '-- Choose Role --', choices: roleChoices },
            { type: 'TextBlock', text: 'Originator:', wrap: true, spacing: 'Medium' },
            { type: 'Input.ChoiceSet', id: 'originatorId', style: 'compact', placeholder: '-- Choose Originator --', choices: memberChoices },
            { type: 'TextBlock', text: 'Reviewer:', wrap: true, spacing: 'Medium' },
            { type: 'Input.ChoiceSet', id: 'reviewerId', style: 'compact', placeholder: '-- Choose Reviewer --', choices: memberChoices },
            { type: 'TextBlock', text: 'Approver:', wrap: true, spacing: 'Medium' },
            { type: 'Input.ChoiceSet', id: 'approverId', style: 'compact', placeholder: '-- Choose Approver --', choices: memberChoices }
        ],
        actions: [{ type: 'Action.Submit', title: '✅ Submit', data: { action: 'jd_form_submit' } }]
    };

    return CardFactory.adaptiveCard(card);
}

function buildJdFormConfirmCard(deptName, roleName, originatorName, reviewerName, approverName) {
    const card = {
        type: 'AdaptiveCard', $schema: 'http://adaptivecards.io/schemas/adaptive-card.json', version: '1.4', msteams: { width: 'Full' },
        body: [
            { type: 'TextBlock', text: '✅ JD Creation Submitted', weight: 'Bolder', size: 'Medium', color: 'Good', wrap: true },
            { type: 'TextBlock', text: 'Your selections:', wrap: true, spacing: 'Small' },
            {
                type: 'FactSet',
                facts: [
                    { title: 'Department:', value: deptName || '—' },
                    { title: 'Role:', value: roleName || '—' },
                    { title: 'Originator:', value: originatorName || '—' },
                    { title: 'Reviewer:', value: reviewerName || '—' },
                    { title: 'Approver:', value: approverName || '—' }
                ],
                spacing: 'Small'
            }
        ]
    };
    return CardFactory.adaptiveCard(card);
}

function buildConfirmCard(deptName, roleName) {
    const card = {
        type: 'AdaptiveCard',
        $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
        version: '1.4',
        body: [
            {
                type: 'TextBlock',
                text: '✅ JD Request Submitted',
                weight: 'Bolder',
                size: 'Medium',
                color: 'Good',
                wrap: true
            },
            {
                type: 'TextBlock',
                text: 'Your selections have been recorded:',
                wrap: true,
                spacing: 'Small'
            },
            {
                type: 'FactSet',
                facts: [
                    { title: 'Department:', value: deptName },
                    { title: 'Role:', value: roleName }
                ],
                spacing: 'Small'
            },
            {
                type: 'TextBlock',
                text: 'The JD creation request has been submitted successfully.',
                wrap: true,
                spacing: 'Medium',
                isSubtle: true
            }
        ]
    };
    return CardFactory.adaptiveCard(card);
}

const EDIT_TYPE_LABELS = {
    edit_all: 'Edit All',
    edit_skills: 'Edit Skills',
    edit_education: 'Edit Education',
    edit_comp: 'Edit Comp',
    edit_experience: 'Edit Experience'
};

function buildEditFormCard(editType) {
    const title = EDIT_TYPE_LABELS[editType] || 'Edit';
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
                        items: [{ type: 'TextBlock', text: title, weight: 'Bolder', size: 'Medium', wrap: true }]
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
            { type: 'TextBlock', text: 'Label:', wrap: true, spacing: 'Medium' },
            { type: 'Input.Text', id: 'label', placeholder: 'Enter label...' },
            { type: 'TextBlock', text: 'Description:', wrap: true, spacing: 'Medium' },
            { type: 'Input.Text', id: 'description', placeholder: 'Enter description...', isMultiline: true },
            { type: 'Input.Text', id: 'editType', value: editType, isVisible: false }
        ],
        actions: [
            { type: 'Action.Submit', title: '✅ Save', data: { action: 'edit_form_submit' } }
        ]
    };
    return CardFactory.adaptiveCard(card);
}

function buildFetchIndentFilterCard(departments, roles) {
    const deptChoices = (departments || []).map((d) => ({ title: d.name, value: String(d.id) }));
    const roleChoices = (roles || []).map((r) => ({ title: r.name, value: String(r.id) }));

    const card = {
        type: 'AdaptiveCard', $schema: 'http://adaptivecards.io/schemas/adaptive-card.json', version: '1.4', msteams: { width: 'Full' },
        body: [
            {
                type: 'ColumnSet',
                columns: [
                    { type: 'Column', width: 'stretch', items: [{ type: 'TextBlock', text: 'Fetch JD', weight: 'Bolder', size: 'Large', wrap: true }] },
                    { type: 'Column', width: 'auto', items: [{ type: 'ActionSet', actions: [{ type: 'Action.Submit', title: '✕', data: { action: 'card_close' } }] }] }
                ]
            },
            { type: 'TextBlock', text: 'Choose Department and Role, then submit to fetch filtered JD records.', wrap: true, spacing: 'Small' },
            { type: 'TextBlock', text: 'Department:', wrap: true, spacing: 'Medium' },
            { type: 'Input.ChoiceSet', id: 'departmentId', style: 'compact', placeholder: '-- Choose Department --', choices: deptChoices },
            { type: 'TextBlock', text: 'Role:', wrap: true, spacing: 'Medium' },
            { type: 'Input.ChoiceSet', id: 'roleId', style: 'compact', placeholder: '-- Choose Role --', choices: roleChoices }
        ],
        actions: [{ type: 'Action.Submit', title: '✅ Submit', data: { action: 'fetch_indent_submit' } }]
    };

    return CardFactory.adaptiveCard(card);
}

function formatCell(value) {
    if (value === null || value === undefined) return '—';
    if (value instanceof Date) return value.toISOString();
    const s = String(value).trim();
    return s || '—';
}

function rowField(row, key) {
    if (!row || typeof row !== 'object') return undefined;
    if (Object.prototype.hasOwnProperty.call(row, key)) return row[key];
    const lower = key.toLowerCase();
    if (Object.prototype.hasOwnProperty.call(row, lower)) return row[lower];
    return undefined;
}

const INDENT_CARD_MAX_ROWS = 12;

function buildIndentJobMasterCard(rows) {
    const body = [
        {
            type: 'ColumnSet',
            columns: [
                { type: 'Column', width: 'stretch', items: [{ type: 'TextBlock', text: 'JD — Job master', weight: 'Bolder', size: 'Large', wrap: true }] },
                { type: 'Column', width: 'auto', items: [{ type: 'ActionSet', actions: [{ type: 'Action.Submit', title: '✕', data: { action: 'card_close' } }] }] }
            ]
        },
        { type: 'TextBlock', text: rows.length === 0 ? 'No records found.' : `${rows.length} record(s) loaded from database.`, wrap: true, spacing: 'Small', isSubtle: true }
    ];

    const display = rows.slice(0, INDENT_CARD_MAX_ROWS);
    if (rows.length > INDENT_CARD_MAX_ROWS) {
        body.push({ type: 'TextBlock', text: `Showing first ${INDENT_CARD_MAX_ROWS} of ${rows.length} (see INDENT_JOB_MASTER_LIMIT in env).`, wrap: true, color: 'Warning', spacing: 'Small' });
    }

    const fields = [
        ['ID', 'id'],
        ['Department', 'Department'],
        ['Role', 'Role'],
        ['Name', 'name'],
        ['Description', 'description'],
        ['Roles & responsibilities', 'roles_and_respo'],
        ['Originator', 'Originator'],
        ['Reviewer', 'Reviewer'],
        ['Approver', 'Approver'],
        ['Status', 'status'],
        ['Active', 'Active'],
        ['Created by', 'created_by'],
        ['Modified by', 'modified_by'],
        ['Created', 'created'],
        ['Modified', 'modified']
    ];

    display.forEach((row, index) => {
        body.push({ type: 'TextBlock', text: `Record ${index + 1}`, weight: 'Bolder', separator: true, spacing: 'Medium', wrap: true });
        for (const [label, key] of fields) {
            body.push({ type: 'TextBlock', text: `**${label}:** ${formatCell(rowField(row, key))}`, wrap: true, spacing: 'None' });
        }
    });

    const card = {
        type: 'AdaptiveCard', $schema: 'http://adaptivecards.io/schemas/adaptive-card.json', version: '1.4', msteams: { width: 'Full' },
        body,
        actions: [
            { type: 'Action.Submit', title: 'Edit All', data: { action: 'edit_all' } },
            { type: 'Action.Submit', title: 'Edit Skills', data: { action: 'edit_skills' } },
            { type: 'Action.Submit', title: 'Edit Education', data: { action: 'edit_education' } },
            { type: 'Action.Submit', title: 'Edit Comp', data: { action: 'edit_comp' } },
            { type: 'Action.Submit', title: 'Edit Experience', data: { action: 'edit_experience' } }
        ]
    };
    return CardFactory.adaptiveCard(card);
}

module.exports = {
    buildMenuCard,
    buildDeptCard,
    buildRoleCard,
    buildConfirmCard,
    buildJdCreationFormCard,
    buildJdFormConfirmCard,
    buildFetchIndentFilterCard,
    buildIndentJobMasterCard,
    buildEditFormCard
};
