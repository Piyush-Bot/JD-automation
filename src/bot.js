const { TeamsInfo, ActivityHandler, MessageFactory } = require('botbuilder');
const {
    buildMenuCard,
    buildJdCreatFormCard,
    buildJdEditCard,
    buildEditFormCard,
    buildJdResultCard
} = require('./services/cards/cardIndex');

const {
    checkMenuEligibility,
    loginForDataApi,
    getDepartments,
    getRolesByDepartment,
    getCollabMembers,
    getJdByRoleAndDept,
    triggerJdWorkflow,
    saveGeneratedJd,
    createJD
} = require('./services/apiService');

async function closeSourceCard(context) {
    const targetId = context.activity.replyToId || context.activity.id;
    if (!targetId) return;
    try {
        await context.deleteActivity(targetId);
    } catch (_) {
        try {
            await context.updateActivity({
                id: targetId,
                type: 'message',
                conversation: context.activity.conversation,
                attachments: [],
                text: '✅ Submitted.'
            });
        } catch (_ignore) {}
    }
}

async function handleCardAction(context) {
    const value = context.activity.value || {};
    const action = value.action;

    // For now, use the fallback email configured in .env for data API calls,
    // but prefer the Teams user's email when available.
    let userEmail = process.env.API_LOGIN_EMAIL_FALLBACK || null;
    try {
        const member = await TeamsInfo.getMember(context, context.activity.from && context.activity.from.id);
        const mEmail = (member && (member.email || member.userPrincipalName)) || null;
        if (mEmail) userEmail = mEmail;
    } catch (_e) {}

    switch (action) {
    case 'jd_start': {
        let departments;
        let roles;
        let members;
        try {
            // Ensure login first so downstream calls reuse the token
            await loginForDataApi(userEmail);
            departments = await getDepartments(userEmail);
            roles = await getRolesByDepartment(undefined, userEmail);
            members = await getCollabMembers(userEmail);
        } catch (err) {
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

        await context.sendActivity({ attachments: [buildJdCreatFormCard(departments, roles, members)] });
        break;
    }

    case 'fetch_indent': {
        let departments;
        let roles;
        try {
            await loginForDataApi(userEmail);
            departments = await getDepartments(userEmail);
            roles = await getRolesByDepartment(undefined, userEmail);
        } catch (err) {
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

        await context.sendActivity({ attachments: [buildJdEditCard(departments, roles)] });
        break;
    }

    case 'jd_edit': {
        const ctx = {
            role: value.role || '',
            department: value.department || '',
            rawOutput: value.rawOutput || ''
        };
        await context.sendActivity({ attachments: [buildEditFormCard(ctx)] });
        break;
    }

    case 'card_close': {
        const targetId = context.activity.replyToId || context.activity.id;
        if (!targetId) {
            return;
        }
        try {
            // Delete the message that contains the card
            await context.deleteActivity(targetId);
        } catch (err) {
            // Fallback: replace with a tiny stub so the large card disappears
            try {
                await context.updateActivity({
                    id: targetId,
                    type: 'message',
                    conversation: context.activity.conversation,
                    attachments: [],
                    text: 'Card closed.'
                });
            } catch (_ignore) {
                // Silent no-op if update also fails
            }
        }
        return;
    }

    case 'edit_form_submit': {
        const editTypes = value.editTypes || '';
        const description = value.description || '';
        if (!editTypes || !description) {
            await context.sendActivity(MessageFactory.text('Please select at least one Edit Type and enter a Description before submitting.'));
            return;
        }
        const typeLabels = editTypes.split(',').map((t) => t.trim()).filter(Boolean).join(', ');
        const prompt = `Edit Type: ${typeLabels}, Description: ${description}`;

        let rawOutput = null;
        try { rawOutput = value.rawOutput ? JSON.parse(value.rawOutput) : null; } catch (_) { rawOutput = null; }

        const editPayload = {
            role: value.role || '',
            department: value.department || '',
            location: process.env.DEFAULT_LOCATION || 'Bangalore',
            prompt,
            output: rawOutput
        };
        const loadingMsg = await context.sendActivity(MessageFactory.text('⏳ Updating JD...'));
        let editRes;
        try {
            editRes = await triggerJdWorkflow(editPayload, userEmail);
        } catch (err) {
            try { await context.deleteActivity(loadingMsg.id); } catch (_) {}
            await context.sendActivity(MessageFactory.text('Sorry, could not update JD. Please try again.'));
            return;
        }

        if (!editRes || editRes.ok !== true) {
            try { await context.deleteActivity(loadingMsg.id); } catch (_) {}
            await context.sendActivity(MessageFactory.text((editRes && editRes.error) ? `Failed to update JD: ${editRes.error}` : 'Sorry, could not update JD.'));
            return;
        }

        const triggerRawOutput = editRes.raw && editRes.raw.workflow_response
            ? [{ output: editRes.raw.workflow_response.output }]
            : null;
        const triggerEditCtx = {
            role: value.role || '',
            department: value.department || '',
            rawOutput: triggerRawOutput
        };
        await context.sendActivity({ attachments: [buildJdResultCard(editRes.output, '✅ JD Updated Successfully', triggerEditCtx, {})] });
        await closeSourceCard(context);
        try { await context.deleteActivity(loadingMsg.id); } catch (_) {}
        break;
    }

    case 'fetch_indent_submit': {
        const departmentId = value.departmentId;
        const roleId = value.roleId;

        const missing = [
            !departmentId && 'Department',
            !roleId && 'Role'
        ].filter(Boolean);
        if (missing.length > 0) {
            await context.sendActivity(MessageFactory.text(`Please select the following before submitting: **${missing.join(', ')}**.`));
            return;
        }

        const loadingMsg = await context.sendActivity(MessageFactory.text('⏳ Fetching JD...'));

        let fetchRes;
        try {
            fetchRes = await getJdByRoleAndDept(roleId, departmentId, userEmail);
        } catch (err) {
            try { await context.deleteActivity(loadingMsg.id); } catch (_) {}
            await context.sendActivity(MessageFactory.text('Sorry, could not fetch JD. Please try again.'));
            return;
        }

        if (!fetchRes || fetchRes.ok !== true) {
            try { await context.deleteActivity(loadingMsg.id); } catch (_) {}
            await context.sendActivity(MessageFactory.text((fetchRes && fetchRes.error) ? `Failed to fetch JD: ${fetchRes.error}` : 'Sorry, could not fetch JD.'));
            return;
        }

        const record = fetchRes.raw && fetchRes.raw.records && fetchRes.raw.records[0];
        const fetchEditCtx = {
            role: (record && record.role) || roleId,
            department: (record && record.department) || departmentId,
            rawOutput: record && record.output
        };

        await context.sendActivity({ attachments: [buildJdResultCard(fetchRes.output, '✅ JD Fetched Successfully', fetchEditCtx, {}, { editEnabled: true, acceptEnabled: false })] });
        await closeSourceCard(context);
        try { await context.deleteActivity(loadingMsg.id); } catch (_) {}
        break;
    }

    case 'jd_accept': {
        let acceptOutput = null;
        try { acceptOutput = value.output ? JSON.parse(value.output) : null; } catch (_) {}
        const acceptPayload = {
            role: value.role || '',
            department: value.department || '',
            originator: value.originator || '',
            reviewer: value.reviewer || '',
            approver: value.approver || '',
            output: acceptOutput
        };
        console.log('[bot] jd_accept payload', JSON.stringify(acceptPayload, null, 2));

        const loadingMsg = await context.sendActivity(MessageFactory.text('⏳ Saving JD...'));
        let saveRes;
        try {
            saveRes = await saveGeneratedJd(acceptPayload, userEmail);
        } catch (err) {
            try { await context.deleteActivity(loadingMsg.id); } catch (_) {}
            await context.sendActivity(MessageFactory.text('Sorry, could not save JD. Please try again.'));
            return;
        }
        if (!saveRes || saveRes.ok !== true) {
            try { await context.deleteActivity(loadingMsg.id); } catch (_) {}
            await context.sendActivity(MessageFactory.text((saveRes && saveRes.error) ? `Failed to save JD: ${saveRes.error}` : 'Sorry, could not save JD.'));
            return;
        }
        await context.sendActivity(MessageFactory.text('✅ JD saved successfully.'));
        try { await context.deleteActivity(loadingMsg.id); } catch (_) {}
        break;
    }

    case 'jd_form_submit': {
        const deptId = value.deptId;
        const roleId = value.roleId;
        const originatorId = value.originatorId;
        const reviewerId = value.reviewerId;
        const approverId = value.approverId;

        const missing = [
            !deptId && 'Department',
            !roleId && 'Role',
            !originatorId && 'Originator',
            !reviewerId && 'Reviewer',
            !approverId && 'Approver'
        ].filter(Boolean);
        if (missing.length > 0) {
            await context.sendActivity(MessageFactory.text(`Please select the following before submitting: **${missing.join(', ')}**.`));
            return;
        }

        // Respond immediately to prevent Teams 5s timeout error banner
        const loadingMsg = await context.sendActivity(MessageFactory.text('⏳ Creating JD...'));

        // Resolve names from IDs first — API payload requires names
        let departments, roles, members;
        try {
            departments = await getDepartments(userEmail);
            roles = await getRolesByDepartment(undefined, userEmail);
            members = await getCollabMembers(userEmail);
        } catch (err) {
            try { await context.deleteActivity(loadingMsg.id); } catch (_) {}
            await context.sendActivity(MessageFactory.text('Sorry, could not resolve your selections. Please try again.'));
            return;
        }

        const dept = departments.find((d) => String(d.id) === String(deptId));
        const role = roles.find((r) => String(r.id) === String(roleId));
        const originator = members.find((m) => String(m.id) === String(originatorId));
        const reviewer = members.find((m) => String(m.id) === String(reviewerId));
        const approver = members.find((m) => String(m.id) === String(approverId));

        // Build name-based payload for the API
        const jdPayload = {
            department: dept ? dept.name : deptId,
            role: role ? role.name : roleId,
            originator: originator ? originator.name : originatorId,
            reviewer: reviewer ? reviewer.name : reviewerId,
            approver: approver ? approver.name : approverId
        };

        // Create JD via API service
        let createRes;
        try {
            createRes = await createJD(jdPayload, userEmail);
        } catch (err) {
            try { await context.deleteActivity(loadingMsg.id); } catch (_) {}
            await context.sendActivity(MessageFactory.text('Sorry, could not create JD. Please try again.'));
            return;
        }

        if (!createRes || createRes.ok !== true) {
            try { await context.deleteActivity(loadingMsg.id); } catch (_) {}
            await context.sendActivity(MessageFactory.text((createRes && createRes.error) ? `Failed to create JD: ${createRes.error}` : 'Sorry, could not create JD.'));
            return;
        }

        if (createRes.output && Object.keys(createRes.output).length > 0) {
            const createEditCtx = {
                role: jdPayload.role,
                department: jdPayload.department,
                rawOutput: createRes.raw && createRes.raw.workflow_result
            };
            const createAcceptCtx = {
                role: jdPayload.role,
                department: jdPayload.department,
                originator: jdPayload.originator,
                reviewer: jdPayload.reviewer,
                approver: jdPayload.approver,
                output: createRes.output
            };
            await context.sendActivity({ attachments: [buildJdResultCard(createRes.output, '✅ JD Created Successfully', createEditCtx, createAcceptCtx)] });
        } else {
            await context.sendActivity(MessageFactory.text('✅ JD Creation request submitted successfully.'));
        }
        await closeSourceCard(context);
        try { await context.deleteActivity(loadingMsg.id); } catch (_) {}
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
            const msAuthHeader = context.turnState.get('msAuthHeader');

            let member = null;
            try {
                member = await TeamsInfo.getMember(context, context.activity.from && context.activity.from.id);
            } catch (_e) {}

            const userEmail = (member && (member.email || member.userPrincipalName)) || process.env.API_LOGIN_EMAIL_FALLBACK || null;
            const aadObjectId = (member && member.aadObjectId) || null;
            const displayName = (member && member.name) || null;

            let eligibility;
            try {
                eligibility = await checkMenuEligibility({
                    userId: context.activity.from && context.activity.from.id,
                    aadObjectId,
                    email: userEmail,
                    displayName,
                    conversationId: context.activity.conversation && context.activity.conversation.id,
                    channelId: context.activity.channelId,
                    tenantId: context.activity.channelData && context.activity.channelData.tenant && context.activity.channelData.tenant.id,
                    serviceUrl: context.activity.serviceUrl,
                    text: userText,
                    msAuthHeader
                });
            } catch (err) {
                await context.sendActivity(MessageFactory.text('Sorry, could not verify eligibility. Please try again later.'));
                await next();
                return;
            }

            if (!eligibility || eligibility.allowed !== true) {
                await context.sendActivity(MessageFactory.text((eligibility && eligibility.reason) || 'You are not allowed to start the JD process at this time.'));
                await next();
                return;
            }

            const intent = eligibility.intent;

            // UNKNOWN_INTENT (or no intent): show JD Process menu card
            if (!intent || intent === 'UNKNOWN_INTENT' || intent === 'UNKOWN_INTENT') {
                await context.sendActivity({ attachments: [buildMenuCard()] });
                await next();
                return;
            }

            // JD_CREATE: skip menu — login then open JD Creation form directly
            if (intent === 'JD_CREATE') {
                let departments, roles, members;
                try {
                    await loginForDataApi(userEmail);
                    departments = await getDepartments(userEmail);
                    roles = await getRolesByDepartment(undefined, userEmail);
                    members = await getCollabMembers(userEmail);
                } catch (err) {
                    await context.sendActivity(MessageFactory.text('Sorry, could not load JD form data. Please try again.'));
                    await next();
                    return;
                }

                if (!departments.length || !roles.length || !members.length) {
                    await context.sendActivity(MessageFactory.text('JD form could not be loaded because dropdown data is empty.'));
                    await next();
                    return;
                }

                await context.sendActivity({ attachments: [buildJdCreatFormCard(departments, roles, members)] });
                await next();
                return;
            }

            // JD_FETCH: skip menu — login then open Fetch JD filter form directly
            if (intent === 'JD_FETCH') {
                let departments, roles;
                try {
                    await loginForDataApi(userEmail);
                    departments = await getDepartments(userEmail);
                    roles = await getRolesByDepartment(undefined, userEmail);
                } catch (err) {
                    await context.sendActivity(MessageFactory.text('Sorry, could not load Fetch JD form data. Please try again.'));
                    await next();
                    return;
                }

                if (!departments.length || !roles.length) {
                    await context.sendActivity(MessageFactory.text('Fetch JD form could not be loaded because dropdown data is empty.'));
                    await next();
                    return;
                }

                await context.sendActivity({ attachments: [buildJdEditCard(departments, roles)] });
                await next();
                return;
            }

            // Unrecognised intent
            await context.sendActivity(MessageFactory.text((eligibility.reason) || 'Unrecognised intent from eligibility service.'));
            await next();
        });

        this.onMembersAdded(async (context, next) => {
            const membersAdded = context.activity.membersAdded || [];
            const welcomeText = 'Hello! Send any message to begin the JD flow.';
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
