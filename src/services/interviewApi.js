const { Client } = require('@microsoft/microsoft-graph-client');
const { ClientSecretCredential } = require('@azure/identity');
require('isomorphic-fetch');

function getGraphApplicationClient() {
    const tenantId = process.env.MicrosoftAppTenantId;
    const clientId = process.env.MicrosoftAppId;
    const clientSecret = process.env.MicrosoftAppPassword;
    const schedulerUserId = process.env.GRAPH_SCHEDULER_USER_ID || process.env.GRAPH_SENDER_EMAIL;

    if (!tenantId || !clientId || !clientSecret) {
        return {
            client: null,
            userId: null,
            error: 'Missing Microsoft app credentials in .env (MicrosoftAppTenantId, MicrosoftAppId, MicrosoftAppPassword).'
        };
    }

    if (!schedulerUserId || !String(schedulerUserId).trim()) {
        return {
            client: null,
            userId: null,
            error: 'Missing GRAPH_SCHEDULER_USER_ID (or GRAPH_SENDER_EMAIL) for app-only scheduling mailbox.'
        };
    }

    const credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
    return {
        client: Client.init({
            authProvider: async (done) => {
                try {
                    const token = await credential.getToken('https://graph.microsoft.com/.default');
                    done(null, token && token.token ? token.token : null);
                } catch (err) {
                    done(err, null);
                }
            }
        }),
        userId: String(schedulerUserId).trim(),
        error: null
    };
}

function buildGraphAttendees(candidateEmail, interviewerEmail) {
    const attendees = [];

    if (candidateEmail) {
        attendees.push({
            emailAddress: { address: candidateEmail },
            type: 'required'
        });
    }

    if (interviewerEmail) {
        attendees.push({
            emailAddress: { address: interviewerEmail },
            type: 'required'
        });
    }

    return attendees;
}

async function scheduleInterviewInTeams(name, time, options = {}) {
    const { client, userId, error } = getGraphApplicationClient();
    if (error) {
        return { ok: false, message: error };
    }

    const durationMinutes = options.durationMinutes || 60;
    const timezone = options.timezone || 'Asia/Kolkata';
    const interviewerEmail = options.interviewerEmail || null;
    const endTime = new Date(time.getTime() + durationMinutes * 60 * 1000);

    const candidateName = typeof name === 'object' ? name.name : name;
    const candidateEmail = typeof name === 'object' ? name.email : null;
    const attendees = buildGraphAttendees(candidateEmail, interviewerEmail);

    const eventPayload = {
        subject: `Interview with ${candidateName}`,
        body: {
            contentType: 'text',
            content: 'AI scheduled interview (Teams)'
        },
        start: {
            dateTime: time.toISOString(),
            timeZone: timezone
        },
        end: {
            dateTime: endTime.toISOString(),
            timeZone: timezone
        },
        attendees,
        isOnlineMeeting: true,
        onlineMeetingProvider: 'teamsForBusiness'
    };

    try {
        const response = await client.api(`/users/${encodeURIComponent(userId)}/events`).post(eventPayload);
        const teamsJoinUrl = response?.onlineMeeting?.joinUrl || response?.webLink || null;
        return {
            ok: true,
            message: `Interview scheduled in Microsoft Teams successfully. Meeting link: ${teamsJoinUrl || 'created (join URL not returned)'}`,
            provider: 'teams',
            raw: response
        };
    } catch (error) {
        const msg = error?.message || 'Unknown Teams scheduling error.';
        return {
            ok: false,
            message: `Teams scheduling failed: ${msg}`
        };
    }
}

async function getInterviewResponsesFromTeams(candidateName) {
    const { client, userId, error } = getGraphApplicationClient();
    if (error) {
        return { ok: false, message: error };
    }

    try {
        const escaped = candidateName.replace(/'/g, "''");
        const response = await client
            .api(`/users/${encodeURIComponent(userId)}/events`)
            .filter(`startswith(subject,'Interview with ${escaped}')`)
            .orderby('start/dateTime desc')
            .top(10)
            .select('id,subject,start,attendees,onlineMeeting,webLink,responseStatus')
            .get();

        const events = response.value || [];
        if (!events.length) {
            return { ok: false, message: `No Teams interview event found for ${candidateName}.` };
        }

        const event = events[events.length - 1];
        const attendees = event.attendees || [];
        if (!attendees.length) {
            return {
                ok: true,
                message: `Teams interview found for ${candidateName}, but no attendees are present on this event.`
            };
        }

        const attendeeLines = attendees.map((attendee) => {
            const status = attendee?.status?.response || attendee?.responseStatus?.response || 'none';
            const email = attendee?.emailAddress?.address || 'unknown';
            return `- ${email}: ${status}`;
        });

        const teamsJoinUrl = event?.onlineMeeting?.joinUrl || event?.webLink || 'N/A';
        return {
            ok: true,
            message: `Teams interview responses for ${candidateName}:\n${attendeeLines.join('\n')}\nMeeting link: ${teamsJoinUrl}`
        };
    } catch (error) {
        return { ok: false, message: `Teams response lookup failed: ${error?.message || 'Unknown error'}` };
    }
}

module.exports = {
    getGraphApplicationClient,
    buildGraphAttendees,
    scheduleInterviewInTeams,
    getInterviewResponsesFromTeams
};
