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
    saveUpdatedJd,
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
        const startFlowSource = (value && value.flowSource) || 'creation';
        let departments;
        let roles;
        let members;
        try {
            // Ensure login first so downstream calls reuse the token
            await loginForDataApi(userEmail, startFlowSource);
            departments = await getDepartments(userEmail, startFlowSource);
            roles = await getRolesByDepartment(undefined, userEmail, startFlowSource);
            members = await getCollabMembers(userEmail, startFlowSource);
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
        const fetchFlowSource = (value && value.flowSource) || 'fetch';
        let departments;
        let roles;
        try {
            await loginForDataApi(userEmail, fetchFlowSource);
            departments = await getDepartments(userEmail, fetchFlowSource);
            roles = await getRolesByDepartment(undefined, userEmail, fetchFlowSource);
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
            originator: value.originator || '',
            reviewer: value.reviewer || '',
            approver: value.approver || '',
            rawOutput: value.rawOutput || '',
            jdId: value.jdId || '',
            flowSource: value.flowSource || ''
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
        let outputList;
        if (rawOutput && Array.isArray(rawOutput.response)) {
            outputList = rawOutput.response;
        } else if (Array.isArray(rawOutput)) {
            outputList = rawOutput.map((x) => (x && x.output !== undefined ? x.output : x));
        } else if (rawOutput && rawOutput.output !== undefined) {
            outputList = Array.isArray(rawOutput.output) ? rawOutput.output : [rawOutput.output];
        } else if (rawOutput != null) {
            outputList = [rawOutput];
        } else {
            outputList = [];
        }
<<<<<<< Updated upstream
        console.log('[edit_form_submit] normalized output isArray:', Array.isArray(outputList), 'length:', Array.isArray(outputList) ? outputList.length : 0);
=======
        // console.log('[edit_form_submit] normalized output isArray:', Array.isArray(outputList), 'length:', Array.isArray(outputList) ? outputList.length : 0);
>>>>>>> Stashed changes

        const normalizedForApi = outputList.map((item) => (
            item && typeof item === 'object' && Object.prototype.hasOwnProperty.call(item, 'output')
                ? item
                : { output: item }
        ));
<<<<<<< Updated upstream
        console.log('[edit_form_submit] normalizedForApi first item has output:', !!(normalizedForApi[0] && normalizedForApi[0].output));
=======
        // console.log('[edit_form_submit] normalizedForApi first item has output:', !!(normalizedForApi[0] && normalizedForApi[0].output));
>>>>>>> Stashed changes

        const editPayload = {
            role: value.role || '',
            department: value.department || '',
            location: process.env.DEFAULT_LOCATION || 'Bangalore',
            prompt,
            output: normalizedForApi
        };
        const loadingMsg = await context.sendActivity(MessageFactory.text('⏳ Updating JD...'));
<<<<<<< Updated upstream
        let editRes;
        try {
            const flowSource = value.flowSource || (value.jdId ? 'fetch' : 'creation');
            editRes = await triggerJdWorkflow(editPayload, userEmail, flowSource);
            console.log('[edit_form_submit] editRes:', JSON.stringify(editRes, null, 2));
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
        const flowSource = value.flowSource || (value.jdId ? 'fetch' : 'creation');
        const triggerEditCtx = {
            role: value.role || '',
            department: value.department || '',
            originator: value.originator || '',
            reviewer: value.reviewer || '',
            approver: value.approver || '',
            rawOutput: triggerRawOutput,
            jdId: value.jdId || '',
            flowSource
=======
        const reference = TurnContext.getConversationReference(context.activity);
        const adapter = context.adapter;
        const appId = process.env.MicrosoftAppId || null;
        const captured = { value: { ...value }, userEmail, editPayload, loadingId: loadingMsg.id };
        const runProactive = async (proactiveContext) => {
            try {
                const flowSource = captured.value.flowSource || (captured.value.jdId ? 'fetch' : 'creation');
                const editRes = await triggerJdWorkflow(captured.editPayload, captured.userEmail, flowSource);
                // console.log('[edit_form_submit] editRes:', JSON.stringify(editRes, null, 2));
                if (!editRes || editRes.ok !== true) {
                    try { await proactiveContext.deleteActivity(captured.loadingId); } catch (_) {}
                    await proactiveContext.sendActivity(MessageFactory.text((editRes && editRes.error) ? `Failed to update JD: ${editRes.error}` : 'Sorry, could not update JD.'));
                    return;
                }
                const triggerRawOutput = editRes.raw && editRes.raw.workflow_response ? [{ output: editRes.raw.workflow_response.output }] : null;
                const triggerEditCtx = {
                    role: captured.value.role || '',
                    department: captured.value.department || '',
                    originator: captured.value.originator || '',
                    reviewer: captured.value.reviewer || '',
                    approver: captured.value.approver || '',
                    rawOutput: triggerRawOutput,
                    jdId: captured.value.jdId || '',
                    flowSource
                };
                if (flowSource === 'fetch') {
                    const updateAcceptCtx = {
                        jdId: captured.value.jdId || '',
                        output: editRes.output,
                        flowSource
                    };
                    await proactiveContext.sendActivity({ attachments: [buildJdResultCard(editRes.output, '✅ JD Updated Successfully', triggerEditCtx, updateAcceptCtx, { acceptAction: 'jd_update_accept' })] });
                } else {
                    const createAcceptCtx = {
                        role: captured.value.role || '',
                        department: captured.value.department || '',
                        originator: captured.value.originator || '',
                        reviewer: captured.value.reviewer || '',
                        approver: captured.value.approver || '',
                        output: editRes.output,
                        flowSource
                    };
                    await proactiveContext.sendActivity({ attachments: [buildJdResultCard(editRes.output, '✅ JD Updated Successfully', triggerEditCtx, createAcceptCtx)] });
                }
            } catch (err) {
                try { await proactiveContext.deleteActivity(captured.loadingId); } catch (_) {}
                await proactiveContext.sendActivity(MessageFactory.text('Sorry, could not update JD. Please try again.'));
                return;
            } finally {
                try { await proactiveContext.deleteActivity(captured.loadingId); } catch (_) {}
            }
>>>>>>> Stashed changes
        };

        if (flowSource === 'fetch') {
            const updateAcceptCtx = {
                jdId: value.jdId || '',
                output: editRes.output,
                flowSource
            };
            await context.sendActivity({ attachments: [buildJdResultCard(editRes.output, '✅ JD Updated Successfully', triggerEditCtx, updateAcceptCtx, { acceptAction: 'jd_update_accept' })] });
        } else {
            const createAcceptCtx = {
                role: value.role || '',
                department: value.department || '',
                originator: value.originator || '',
                reviewer: value.reviewer || '',
                approver: value.approver || '',
                output: editRes.output,
                flowSource
            };
            await context.sendActivity({ attachments: [buildJdResultCard(editRes.output, '✅ JD Updated Successfully', triggerEditCtx, createAcceptCtx)] });
        }
        try { await context.deleteActivity(loadingMsg.id); } catch (_) {}
        break;
    }

    case 'fetch_indent_submit': {
        const departmentId = value.departmentId;
        const roleId = value.roleId;
        const flowSource = value.flowSource || 'fetch';

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
            fetchRes = await getJdByRoleAndDept(roleId, departmentId, userEmail, flowSource);
        } catch (err) {
            await context.updateActivity({
                id: loadingMsg.id,
                type: 'message',
                conversation: context.activity.conversation,
                attachments: [],
                text: 'Sorry, could not fetch JD. Please try again.'
            });
            return;
        }

        if (!fetchRes || fetchRes.ok !== true) {
            const errText = (fetchRes && fetchRes.error) ? `Failed to fetch JD: ${fetchRes.error}` : 'Sorry, could not fetch JD.';
            await context.updateActivity({
                id: loadingMsg.id,
                type: 'message',
                conversation: context.activity.conversation,
                attachments: [],
                text: errText
            });
            return;
        }

        const data = fetchRes && fetchRes.raw;
        const noRecords = !data || data.found === false || !Array.isArray(data.records) || data.records.length === 0;
        if (noRecords) {
            const notFoundText = (data && data.message)
                ? data.message
                : `No JD found for role_id=${roleId}, department_id=${departmentId}`;
            await context.updateActivity({
                id: loadingMsg.id,
                type: 'message',
                conversation: context.activity.conversation,
                attachments: [],
                text: notFoundText
            });
            return;
        }

        const record = fetchRes.raw && fetchRes.raw.records && fetchRes.raw.records[0];
        const fetchEditCtx = {
            role: (record && record.role) || roleId,
            department: (record && record.department) || departmentId,
            rawOutput: record && record.output,
            jdId: (record && record.jd_id) || '',
            flowSource: 'fetch'
        };

        await context.sendActivity({ attachments: [buildJdResultCard(fetchRes.output, '✅ JD Fetched Successfully', fetchEditCtx, {}, { editEnabled: true, acceptEnabled: false })] });
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
        const flowSourceAccept = value.flowSource || 'creation';
        console.log('[bot] jd_accept payload', { flowSource: flowSourceAccept, payload: acceptPayload });

        const loadingMsg = await context.sendActivity(MessageFactory.text('⏳ Saving JD...'));
        let saveRes;
        try {
            saveRes = await saveGeneratedJd(acceptPayload, userEmail, flowSourceAccept);
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
        const creationFlow = value.flowSource || 'creation';

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
            departments = await getDepartments(userEmail, creationFlow);
            roles = await getRolesByDepartment(undefined, userEmail, creationFlow);
            members = await getCollabMembers(userEmail, creationFlow);
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
            createRes = await createJD(jdPayload, userEmail, creationFlow);
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
                originator: jdPayload.originator,
                reviewer: jdPayload.reviewer,
                approver: jdPayload.approver,
                rawOutput: createRes.raw && createRes.raw.workflow_result,
                flowSource: 'creation'
            };
            const createAcceptCtx = {
                role: jdPayload.role,
                department: jdPayload.department,
                originator: jdPayload.originator,
                reviewer: jdPayload.reviewer,
                approver: jdPayload.approver,
                output: createRes.output,
                flowSource: 'creation'
            };
            await context.sendActivity({ attachments: [buildJdResultCard(createRes.output, '✅ JD Created Successfully', createEditCtx, createAcceptCtx)] });
        } else {
            await context.sendActivity(MessageFactory.text('✅ JD Creation request submitted successfully.'));
        }
        try { await context.deleteActivity(loadingMsg.id); } catch (_) {}
        break;
    }

    case 'jd_update_accept': {
        let acceptOutput = null;
        try { acceptOutput = value.output ? JSON.parse(value.output) : null; } catch (_) {}
        const updatePayload = {
            jd_id: value.jdId || '',
            output: acceptOutput
        };
        const flowSourceUpdate = value.flowSource || 'fetch';
        console.log('[bot] jd_update_accept payload:', { flowSource: flowSourceUpdate, payload: updatePayload });
        const loadingMsg = await context.sendActivity(MessageFactory.text('⏳ Saving updated JD...'));
        let saveRes;
        try {
            saveRes = await saveUpdatedJd(updatePayload, userEmail, flowSourceUpdate);
        } catch (err) {
            try { await context.deleteActivity(loadingMsg.id); } catch (_) {}
            await context.sendActivity(MessageFactory.text('Sorry, could not save updated JD. Please try again.'));
            return;
        }
        if (!saveRes || saveRes.ok !== true) {
            try { await context.deleteActivity(loadingMsg.id); } catch (_) {}
            await context.sendActivity(MessageFactory.text((saveRes && saveRes.error) ? `Failed to save updated JD: ${saveRes.error}` : 'Sorry, could not save updated JD.'));
            return;
        }
        await context.sendActivity(MessageFactory.text('✅ JD updated and saved successfully.'));
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
                try {
                    await handleCardAction(context);
                } catch (err) {
                    console.error('[onMessage] handleCardAction error:', err);
                }
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
                const flowFromIntent = 'creation';
                let departments, roles, members;
                try {
                    await loginForDataApi(userEmail, flowFromIntent);
                    departments = await getDepartments(userEmail, flowFromIntent);
                    roles = await getRolesByDepartment(undefined, userEmail, flowFromIntent);
                    members = await getCollabMembers(userEmail, flowFromIntent);
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
                const flowFromIntent = 'fetch';
                let departments, roles;
                try {
                    await loginForDataApi(userEmail, flowFromIntent);
                    departments = await getDepartments(userEmail, flowFromIntent);
                    roles = await getRolesByDepartment(undefined, userEmail, flowFromIntent);
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

    async onInvokeActivity(context) {
        return { status: 200, body: { statusCode: 200 } };
    }
}

module.exports = { EchoBot };
