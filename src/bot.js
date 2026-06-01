const { TeamsInfo, ActivityHandler, MessageFactory, TurnContext } = require('botbuilder');
const {
    buildMenuCard,
    buildJdCreatFormCard,
    buildJdEditCard,
    buildEditFormCard,
    buildJdResultCard,
    buildPopupCard,
    buildTaskModuleResponse
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

const {
    parseScheduleCommand,
    parseResponseQuery,
    scheduleInterview,
    getInterviewResponses,
    getCandidateList,
    findCandidateByName,
    createDateAtTime
} = require('./services/interviewScheduler');

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
    case 'popup_ok':
    case 'popup_cancel': {
        const flow = (value && value.flow) || 'creation';
        const prefill = !!(value && value.prefill);
        const deptId = value && value.deptId != null ? String(value.deptId) : null;
        const roleId = value && value.roleId != null ? String(value.roleId) : null;

        await closeSourceCard(context);

        try {
            await loginForDataApi(userEmail, flow);
            const departments = await getDepartments(userEmail, flow);
            const roles = await getRolesByDepartment(undefined, userEmail, flow);
            if (flow === 'creation') {
                const members = await getCollabMembers(userEmail, flow);
                if (!departments.length || !roles.length || !members.length) {
                    await context.sendActivity(MessageFactory.text('JD form could not be loaded because dropdown data is empty.'));
                    return;
                }
                let defaults;
                if (prefill) {
                    defaults = {};
                    if (deptId) {
                        const hasDept = departments.some((d) => String(d.id) === String(deptId));
                        if (hasDept) defaults.deptId = deptId;
                    }
                    if (roleId) {
                        const hasRole = roles.some((r) => String(r.id) === String(roleId));
                        if (hasRole) defaults.roleId = roleId;
                    }
                    if (!defaults.deptId && !defaults.roleId) defaults = undefined;
                }
                await context.sendActivity({ attachments: [buildJdCreatFormCard(departments, roles, members, defaults)] });
            } else {
                if (!departments.length || !roles.length) {
                    await context.sendActivity(MessageFactory.text('Fetch JD form could not be loaded because dropdown data is empty.'));
                    return;
                }
                let defaults;
                if (prefill) {
                    defaults = {};
                    if (deptId) {
                        const hasDept = departments.some((d) => String(d.id) === String(deptId));
                        if (hasDept) defaults.departmentId = deptId;
                    }
                    if (roleId) {
                        const hasRole = roles.some((r) => String(r.id) === String(roleId));
                        if (hasRole) defaults.roleId = roleId;
                    }
                    if (!defaults.departmentId && !defaults.roleId) defaults = undefined;
                }
                await context.sendActivity({ attachments: [buildJdEditCard(departments, roles, defaults)] });
            }
        } catch (err) {
            await context.sendActivity(MessageFactory.text('Sorry, could not load data. Please try again.'));
        }
        return;
    }
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
        const normalizedForApi = outputList.map((item) => (
            item && typeof item === 'object' && Object.prototype.hasOwnProperty.call(item, 'output')
                ? item
                : { output: item }
        ));
        const editPayload = {
            role: value.role || '',
            department: value.department || '',
            location: process.env.DEFAULT_LOCATION || 'Bangalore',
            prompt,
            output: normalizedForApi
        };
        const loadingMsg = await context.sendActivity(MessageFactory.text('⏳ Updating JD...'));
        const reference = TurnContext.getConversationReference(context.activity);
        const adapter = context.adapter;
        const appId = process.env.MicrosoftAppId || null;
        const captured = { value: { ...value }, userEmail, editPayload, loadingId: loadingMsg.id };
        const runProactive = async (proactiveContext) => {
            try {
                const flowSource = captured.value.flowSource || (captured.value.jdId ? 'fetch' : 'creation');
                const editRes = await triggerJdWorkflow(captured.editPayload, captured.userEmail, flowSource);
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
        };
        setImmediate(() => {
            if (typeof adapter.continueConversationAsync === 'function') {
                adapter.continueConversationAsync(appId, reference, async (proactiveContext) => {
                    await runProactive(proactiveContext);
                });
            } else if (typeof adapter.continueConversation === 'function') {
                adapter.continueConversation(reference, async (proactiveContext) => {
                    await runProactive(proactiveContext);
                });
            }
        });
        return;
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

        const loadingMsg = await context.sendActivity(MessageFactory.text('⏳ Creating JD...'));
        const reference = TurnContext.getConversationReference(context.activity);
        const adapter = context.adapter;
        const appId = process.env.MicrosoftAppId || null;
        const captured = { userEmail, creationFlow, deptId, roleId, originatorId, reviewerId, approverId, loadingId: loadingMsg.id };
        const runProactiveCreate = async (proactiveContext) => {
            let departments, roles, members;
            try {
                departments = await getDepartments(captured.userEmail, captured.creationFlow);
                roles = await getRolesByDepartment(undefined, captured.userEmail, captured.creationFlow);
                members = await getCollabMembers(captured.userEmail, captured.creationFlow);
            } catch (err) {
                try { await proactiveContext.deleteActivity(captured.loadingId); } catch (_) {}
                await proactiveContext.sendActivity(MessageFactory.text('Sorry, could not resolve your selections. Please try again.'));
                return;
            }
            const dept = departments.find((d) => String(d.id) === String(captured.deptId));
            const role = roles.find((r) => String(r.id) === String(captured.roleId));
            const originator = members.find((m) => String(m.id) === String(captured.originatorId));
            const reviewer = members.find((m) => String(m.id) === String(captured.reviewerId));
            const approver = members.find((m) => String(m.id) === String(captured.approverId));
            const jdPayload = {
                department: dept ? dept.name : captured.deptId,
                role: role ? role.name : captured.roleId,
                originator: originator ? originator.name : captured.originatorId,
                reviewer: reviewer ? reviewer.name : captured.reviewerId,
                approver: approver ? approver.name : captured.approverId
            };
            let createRes;
            try {
                createRes = await createJD(jdPayload, captured.userEmail, captured.creationFlow);
            } catch (err) {
                try { await proactiveContext.deleteActivity(captured.loadingId); } catch (_) {}
                await proactiveContext.sendActivity(MessageFactory.text('Sorry, could not create JD. Please try again.'));
                return;
            }
            if (!createRes || createRes.ok !== true) {
                try { await proactiveContext.deleteActivity(captured.loadingId); } catch (_) {}
                await proactiveContext.sendActivity(MessageFactory.text((createRes && createRes.error) ? `Failed to create JD: ${createRes.error}` : 'Sorry, could not create JD.'));
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
                await proactiveContext.sendActivity({ attachments: [buildJdResultCard(createRes.output, '✅ JD Created Successfully', createEditCtx, createAcceptCtx)] });
            } else {
                await proactiveContext.sendActivity(MessageFactory.text('✅ JD Creation request submitted successfully.'));
            }
            try { await proactiveContext.deleteActivity(captured.loadingId); } catch (_) {}
        };
        setImmediate(() => {
            if (typeof adapter.continueConversationAsync === 'function') {
                adapter.continueConversationAsync(appId, reference, async (proactiveContext) => {
                    await runProactiveCreate(proactiveContext);
                });
            } else if (typeof adapter.continueConversation === 'function') {
                adapter.continueConversation(reference, async (proactiveContext) => {
                    await runProactiveCreate(proactiveContext);
                });
            }
        });
        return;
    }

    case 'jd_update_accept': {
        let acceptOutput = null;
        try { acceptOutput = value.output ? JSON.parse(value.output) : null; } catch (_) {}
        const updatePayload = {
            jd_id: value.jdId || '',
            output: acceptOutput
        };
        const flowSourceUpdate = value.flowSource || 'fetch';
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
                } catch (err) {}
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

            // Check for interview scheduling commands first (before intent check)
            const scheduleDetails = parseScheduleCommand(userText);

            if (scheduleDetails) {
                if (scheduleDetails.error) {
                    await context.sendActivity(MessageFactory.text(scheduleDetails.error));
                    await next();
                    return;
                }

                const durationMinutes = scheduleDetails.durationMinutes;
                const timeText = scheduleDetails.timeText;
                const candidateName = scheduleDetails.candidateName;
                const interviewerEmail = scheduleDetails.interviewerEmail;
                const startTime = createDateAtTime(timeText, scheduleDetails.dayOffset);

                if (!startTime) {
                    await context.sendActivity(MessageFactory.text('Invalid time format. Example: 5:00 PM'));
                    await next();
                    return;
                }

                const candidate = findCandidateByName(candidateName);
                if (!candidate) {
                    await context.sendActivity(MessageFactory.text(`Candidate '${candidateName}' not found. Ask 'candidate list' to see available names.`));
                    await next();
                    return;
                }

                const result = await scheduleInterview(candidate, startTime, {
                    durationMinutes,
                    timezone: scheduleDetails.timezone,
                    interviewerEmail
                });

                await context.sendActivity(MessageFactory.text(result));
                await next();
                return;
            }

            const responseQuery = parseResponseQuery(userText);
            if (responseQuery) {
                if (responseQuery.error) {
                    await context.sendActivity(MessageFactory.text(responseQuery.error));
                    await next();
                    return;
                }

                const result = await getInterviewResponses(responseQuery.candidateName);
                await context.sendActivity(MessageFactory.text(result));
                await next();
                return;
            }

            if (userText.toLowerCase().includes('candidate list')) {
                const result = await getCandidateList();
                await context.sendActivity(MessageFactory.text(result));
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

            // JD_CREATE: apply semantic prefill logic
            if (intent === 'JD_CREATE') {
                const flowFromIntent = 'creation';
                const sp = eligibility && eligibility.semantic_prefill;

                // Case 1: No match
                if (sp && sp.status === false && !sp.role && !sp.department) {
                    const lines = [
                        "Couldn't find the requested Department and Role in GPM_DB.",
                        'You can continue and select them manually.'
                    ];
                    await context.sendActivity({ attachments: [buildPopupCard({
                        title: 'No match found',
                        bodyTextLines: lines,
                        okData: { flow: flowFromIntent, prefill: false },
                        cancelData: { flow: flowFromIntent, prefill: false }
                    })] });
                    await next();
                    return;
                }

                // Case 2: Partial match
                if (sp && sp.status === false && (sp.role || sp.department)) {
                    const deptName = sp.department ? ((sp.department.record && (sp.department.record.department || sp.department.record.name)) || sp.department.value || 'Department') : 'Not entered';
                    const roleName = sp.role ? ((sp.role.record && (sp.role.record.role || sp.role.record.name)) || sp.role.value || 'Role') : 'Not entered';
                    const deptExact = !!(sp.department && sp.department.exact);
                    const roleExact = !!(sp.role && sp.role.exact);
                    const deptId = sp && sp.department && sp.department.record && sp.department.record.id;
                    const roleId = sp && sp.role && sp.role.record && sp.role.record.id;
                    await context.sendActivity({ attachments: [buildPopupCard({
                        title: 'Confirm details',
                        details: {
                            headerIntro: 'Creating JD for:',
                            department: { name: deptName, status: sp.department ? (deptExact ? 'exact' : 'closest') : null },
                            role: { name: roleName, status: sp.role ? (roleExact ? 'exact' : 'closest') : null }
                        },
                        footerHint: 'Note: Proceed will use the suggested Department and Role. \nCancel lets you choose them manually.',
                        okData: { flow: flowFromIntent, prefill: true, deptId, roleId },
                        cancelData: { flow: flowFromIntent, prefill: false }
                    })] });
                    await next();
                    return;
                }

                // Case 3: Exact match
                if (sp && sp.status === true && (sp.role || sp.department)) {
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
                    const deptId = sp && sp.department && sp.department.record && sp.department.record.id;
                    const roleId = sp && sp.role && sp.role.record && sp.role.record.id;
                    const defaults = {};
                    if (deptId) defaults.deptId = deptId;
                    if (roleId) defaults.roleId = roleId;
                    await context.sendActivity({ attachments: [buildJdCreatFormCard(departments, roles, members, Object.keys(defaults).length ? defaults : undefined)] });
                    await next();
                    return;
                }

                // Fallback: open form without prefill
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

            // JD_FETCH: apply semantic prefill logic
            if (intent === 'JD_FETCH') {
                const flowFromIntent = 'fetch';
                const sp = eligibility && eligibility.semantic_prefill;

                // Case 1: No match
                if (sp && sp.status === false && !sp.role && !sp.department) {
                    const lines = [
                        "Couldn't find the requested Department and Role in GPM_DB.",
                        'You can continue and select them manually.'
                    ];
                    await context.sendActivity({ attachments: [buildPopupCard({
                        title: 'No match found',
                        bodyTextLines: lines,
                        okData: { flow: flowFromIntent, prefill: false },
                        cancelData: { flow: flowFromIntent, prefill: false }
                    })] });
                    await next();
                    return;
                }

                // Case 2: Partial match
                if (sp && sp.status === false && (sp.role || sp.department)) {
                    const deptName = sp.department ? ((sp.department.record && (sp.department.record.department || sp.department.record.name)) || sp.department.value || 'Department') : 'Not entered';
                    const roleName = sp.role ? ((sp.role.record && (sp.role.record.role || sp.role.record.name)) || sp.role.value || 'Role') : 'Not entered';
                    const deptExact = !!(sp.department && sp.department.exact);
                    const roleExact = !!(sp.role && sp.role.exact);
                    const deptId = sp && sp.department && sp.department.record && sp.department.record.id;
                    const roleId = sp && sp.role && sp.role.record && sp.role.record.id;
                    await context.sendActivity({ attachments: [buildPopupCard({
                        title: 'Confirm details',
                        details: {
                            headerIntro: 'Fetching JD for:',
                            department: { name: deptName, status: sp.department ? (deptExact ? 'exact' : 'closest') : null },
                            role: { name: roleName, status: sp.role ? (roleExact ? 'exact' : 'closest') : null }
                        },
                        footerHint: 'Note: Proceed will use the suggested Department and Role. \nCancel lets you choose them manually.',
                        okData: { flow: flowFromIntent, prefill: true, deptId, roleId },
                        cancelData: { flow: flowFromIntent, prefill: false }
                    })] });
                    await next();
                    return;
                }

                // Case 3: Exact match
                if (sp && sp.status === true && (sp.role || sp.department)) {
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
                    const deptId = sp && sp.department && sp.department.record && sp.department.record.id;
                    const roleId = sp && sp.role && sp.role.record && sp.role.record.id;
                    const defaults = {};
                    if (deptId) defaults.departmentId = deptId;
                    if (roleId) defaults.roleId = roleId;
                    await context.sendActivity({ attachments: [buildJdEditCard(departments, roles, Object.keys(defaults).length ? defaults : undefined)] });
                    await next();
                    return;
                }

                // Fallback: open form without prefill
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
        const invokeType = context.activity && context.activity.name;

        if (invokeType === 'task/fetch') {
            const invokeData = (context.activity.value && context.activity.value.data) || context.activity.value || {};
            const action = invokeData.action;
            if (action === 'open_popup' || !action) {
                return { status: 200, body: buildTaskModuleResponse() };
            }
            return { status: 200, body: { task: { type: 'message', value: 'Unknown popup action.' } } };
        }

        if (invokeType === 'task/submit') {
            const submitData = (context.activity.value && context.activity.value.data) || {};
            const action = submitData.action;
            if (action === 'popup_ok') {
                // TODO: define OK action
                await context.sendActivity({ type: 'invokeResponse', value: { status: 200, body: { task: { type: 'message', value: '✅ Confirmed.' } } } });
            } else if (action === 'popup_cancel') {
                // TODO: define Cancel action
                await context.sendActivity({ type: 'invokeResponse', value: { status: 200, body: { task: { type: 'message', value: '❌ Cancelled.' } } } });
            } else {
                await context.sendActivity({ type: 'invokeResponse', value: { status: 200, body: {} } });
            }
            return null;
        }

        await context.sendActivity({ type: 'invokeResponse', value: { status: 200, body: { statusCode: 200 } } });
        return null;
    }
}

module.exports = { EchoBot };
