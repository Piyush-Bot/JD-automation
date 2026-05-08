const { ActivityHandler, MessageFactory } = require('botbuilder');
const {
    buildMenuCard,
    buildDeptCard,
    buildRoleCard,
    buildConfirmCard,
    buildJdCreationFormCard,
    buildJdFormConfirmCard,
    buildFetchIndentFilterCard,
    buildIndentJobMasterCard
} = require('./services/jdCardService');

const {
    checkMenuEligibility,
    getDepartments,
    getRolesByDepartment,
    getJobMasterForIndent,
    getJobMasterForIndentByFilters,
    getCollabMembers,
    createJD
} = require('./services/apiService');

async function handleCardAction(context) {
    const value = context.activity.value || {};
    const action = value.action;

    switch (action) {
    case 'jd_start': {
        let departments;
        let roles;
        let members;
        try {
            departments = await getDepartments();
            roles = await getRolesByDepartment();
            members = await getCollabMembers();
        } catch (err) {
            console.error('[bot] Failed to load JD form dropdown data from DB:', err.message || err);
            await context.sendActivity(
                MessageFactory.text('Sorry, could not load JD form data. Please check the database connection and tables.')
            );
            return;
        }

        if (!departments || !roles || !members || departments.length === 0 || roles.length === 0 || members.length === 0) {
            await context.sendActivity(
                MessageFactory.text('JD form could not be loaded because dropdown data is empty. Please check your HR modules and pcollab_members data.')
            );
            return;
        }

        await context.sendActivity({ attachments: [buildJdCreationFormCard(departments, roles, members)] });
        break;
    }

    case 'dept_selected': {
        const deptId = value.deptId;
        if (!deptId) {
            await context.sendActivity(MessageFactory.text('Please select a department before continuing.'));
            return;
        }

        // Resolve department name by fetching all depts and finding the selected one
        let departments;
        try {
            departments = await getDepartments();
        } catch (err) {
            console.error('[bot] Failed to load departments from DB:', err.message || err);
            await context.sendActivity(MessageFactory.text('Sorry, could not load departments. Please check the database connection.'));
            return;
        }
        const selectedDept = departments.find((d) => String(d.id) === String(deptId));
        const deptName = selectedDept ? selectedDept.name : deptId;

        let roles;
        try {
            roles = await getRolesByDepartment(deptId);
        } catch (err) {
            console.error('[bot] Failed to load roles from DB:', err.message || err);
            await context.sendActivity(MessageFactory.text('Sorry, could not load roles. Please check the database connection.'));
            return;
        }
        if (!roles || roles.length === 0) {
            await context.sendActivity(MessageFactory.text(`No roles found for the department "${deptName}".`));
            return;
        }
        await context.sendActivity({ attachments: [buildRoleCard(deptId, deptName, roles)] });
        break;
    }

    case 'jd_submit': {
        const deptName = value.deptName || value.deptId || 'Unknown';
        const roleId = value.roleId;

        if (!roleId) {
            await context.sendActivity(MessageFactory.text('Please select a role before submitting.'));
            return;
        }

        // Resolve role name
        let roles;
        try {
            roles = await getRolesByDepartment(value.deptId);
        } catch (err) {
            console.error('[bot] Failed to load roles from DB:', err.message || err);
        }
        const selectedRole = roles && roles.find((r) => String(r.id) === String(roleId));
        const roleName = selectedRole ? selectedRole.name : roleId;

        await context.sendActivity({ attachments: [buildConfirmCard(deptName, roleName)] });
        break;
    }

    case 'indent_start': {
        let rows;
        try {
            rows = await getJobMasterForIndent();
        } catch (err) {
            console.error('[bot] Failed to load job master for indent:', err.message || err);
            await context.sendActivity(MessageFactory.text('Sorry, could not load job master data. Please check the database connection and table names.'));
            return;
        }
        await context.sendActivity({ attachments: [buildIndentJobMasterCard(rows)] });
        break;
    }

    // Backward/compat alias for the menu button title "Fetch Indent Data"
    case 'fetch_indent': {
        let departments;
        let roles;
        try {
            departments = await getDepartments();
            roles = await getRolesByDepartment();
        } catch (err) {
            console.error('[bot] Failed to load fetch-jd dropdown data from DB:', err.message || err);
            await context.sendActivity(
                MessageFactory.text('Sorry, could not load department/role data for fetch-jd. Please check the database connection and tables.')
            );
            return;
        }

        if (!departments || !roles || departments.length === 0 || roles.length === 0) {
            await context.sendActivity(
                MessageFactory.text('Fetch Indent Data form could not be loaded because department/role data is empty.')
            );
            return;
        }

        await context.sendActivity({ attachments: [buildFetchIndentFilterCard(departments, roles)] });
        break;
    }

    // Dummy edit buttons from the "Fetch JD" results card
    case 'edit_all':
        await context.sendActivity(MessageFactory.text('Dummy: Edit All (not implemented yet).'));
        break;
    case 'edit_skills':
        await context.sendActivity(MessageFactory.text('Dummy: Edit Skills (not implemented yet).'));
        break;
    case 'edit_education':
        await context.sendActivity(MessageFactory.text('Dummy: Edit Education (not implemented yet).'));
        break;
    case 'edit_comp':
        await context.sendActivity(MessageFactory.text('Dummy: Edit Comp (not implemented yet).'));
        break;
    case 'edit_experience':
        await context.sendActivity(MessageFactory.text('Dummy: Edit Experience (not implemented yet).'));
        break;

    case 'fetch_indent_submit': {
        const departmentId = value.departmentId;
        const roleId = value.roleId;

        if (!departmentId || !roleId) {
            await context.sendActivity(MessageFactory.text('Please select Department and Role before submitting.'));
            return;
        }

        let rows;
        try {
            rows = await getJobMasterForIndentByFilters(roleId, departmentId);
        } catch (err) {
            console.error('[bot] Failed to load filtered indent data:', err.message || err);
            await context.sendActivity(
                MessageFactory.text('Sorry, could not fetch filtered indent data. Please check the database connection and table names.')
            );
            return;
        }

        await context.sendActivity({ attachments: [buildIndentJobMasterCard(rows)] });
        break;
    }

    case 'jd_form_submit': {
        const deptId = value.deptId;
        const roleId = value.roleId;
        const originatorId = value.originatorId;
        const reviewerId = value.reviewerId;
        const approverId = value.approverId;

        if (!deptId || !roleId || !originatorId || !reviewerId || !approverId) {
            await context.sendActivity(MessageFactory.text('Please select all dropdown values before submitting.'));
            return;
        }

        // Create JD via API service
        let createRes;
        try {
            createRes = await createJD({ deptId, roleId, originatorId, reviewerId, approverId });
        } catch (err) {
            console.error('[bot] JD create failed:', err.message || err);
            await context.sendActivity(MessageFactory.text('Sorry, could not create JD. Please try again.'));
            return;
        }
        if (!createRes || createRes.ok !== true) {
            await context.sendActivity(MessageFactory.text((createRes && createRes.error) ? `Failed to create JD: ${createRes.error}` : 'Sorry, could not create JD.'));
            return;
        }

        let departments;
        let roles;
        let members;
        try {
            departments = await getDepartments();
            roles = await getRolesByDepartment();
            members = await getCollabMembers();
        } catch (err) {
            console.error('[bot] Failed to resolve JD selections from DB:', err.message || err);
            await context.sendActivity(MessageFactory.text('Sorry, could not resolve your selections. Please try again.'));
            return;
        }

        const dept = departments.find((d) => String(d.id) === String(deptId));
        const role = roles.find((r) => String(r.id) === String(roleId));
        const originator = members.find((m) => String(m.id) === String(originatorId));
        const reviewer = members.find((m) => String(m.id) === String(reviewerId));
        const approver = members.find((m) => String(m.id) === String(approverId));

        await context.sendActivity({
            attachments: [
                buildJdFormConfirmCard(
                    dept ? dept.name : deptId,
                    role ? role.name : roleId,
                    originator ? originator.name : originatorId,
                    reviewer ? reviewer.name : reviewerId,
                    approver ? approver.name : approverId
                )
            ]
        });
        break;
    }

    default:
        await context.sendActivity(MessageFactory.text('Unknown action. Please try again.'));
    }
}

class EchoBot extends ActivityHandler {
    constructor() {
        super();

        this.onMessage(async (context, next) => {
            // Adaptive Card Action.Submit
            if (context.activity.value) {
                await handleCardAction(context);
                await next();
                return;
            }

            const userText = (context.activity.text || '').trim();

            if (/start jd process/i.test(userText)) {
                let allowedResp;
                try {
                    allowedResp = await checkMenuEligibility({
                        userId: context && context.activity && context.activity.from && context.activity.from.id,
                        conversationId: context && context.activity && context.activity.conversation && context.activity.conversation.id,
                        channelId: context && context.activity && context.activity.channelId,
                        tenantId: context && context.activity && context.activity.channelData && context.activity.channelData.tenant && context.activity.channelData.tenant.id
                    });
                } catch (err) {
                    console.error('[bot] Eligibility check failed:', err.message || err);
                    await context.sendActivity(MessageFactory.text('Sorry, could not verify eligibility to start. Please try again later.'));
                    await next();
                    return;
                }

                if (!allowedResp || allowedResp.allowed !== true) {
                    await context.sendActivity(MessageFactory.text((allowedResp && allowedResp.reason) || 'You are not allowed to start the JD process at this time.'));
                    await next();
                    return;
                }

                await context.sendActivity({ attachments: [buildMenuCard()] });
                await next();
                return;
            }

            await context.sendActivity(MessageFactory.text('Type "start jd process" to begin.'));
            await next();
        });

        this.onMembersAdded(async (context, next) => {
            const membersAdded = context.activity.membersAdded || [];
            const welcomeText = 'Hello! Type "start jd process" to start the JD flow.';
            for (const m of membersAdded) {
                if (m.id !== context.activity.recipient.id) {
                    await context.sendActivity(MessageFactory.text(welcomeText, welcomeText));
                }
            }
            await next();
        });
    }
}

module.exports = { EchoBot };
