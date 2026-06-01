const crypto = require('crypto');
const meetingStore = require('./meetingStore');
const {
    scheduleInterviewInTeams,
    getInterviewResponsesFromTeams
} = require('./interviewApi');

const candidates = [
    {
        id: 1,
        name: 'Pranay',
        email: 'pranay@bellwether.org.in',
        role: 'backend Developer',
        experience: '5 years'
    },
    {
        id: 2,
        name: 'Rajat',
        email: 'rajat@bellwether.org.in',
        role: 'Testing Developer',
        experience: '5 years'
    }
];

function findCandidateByName(name) {
    if (!name) {
        return null;
    }

    const normalized = name.trim().toLowerCase();
    return candidates.find((candidate) => candidate.name.toLowerCase() === normalized) || null;
}

async function scheduleInterview(name, time, options = {}) {
    const result = await scheduleInterviewWithMetadata(name, time, options);
    return result.message;
}

async function scheduleInterviewWithMetadata(name, time, options = {}) {
    try {
        const teamsResult = await scheduleInterviewInTeams(name, time, options);
        if (teamsResult.ok) {
            const durationMinutes = options.durationMinutes || 60;
            const startTime = new Date(time);
            const endTime = new Date(startTime.getTime() + durationMinutes * 60 * 1000);
            const candidateName = typeof name === 'object' ? name.name : name;
            const candidateEmail = typeof name === 'object' ? name.email : null;
            const meetingId = teamsResult?.raw?.id || `local-${crypto.randomUUID()}`;
            const joinUrl = teamsResult?.raw?.onlineMeeting?.joinUrl || teamsResult?.raw?.webLink || null;

            meetingStore.saveMeeting({
                id: meetingId,
                provider: 'teams',
                subject: `Interview with ${candidateName}`,
                candidateName,
                candidateEmail,
                interviewerEmail: options.interviewerEmail || null,
                startTime: startTime.toISOString(),
                endTime: endTime.toISOString(),
                timezone: options.timezone || 'Asia/Kolkata',
                joinUrl,
                autoJoinBot: true,
                organizerId: process.env.GRAPH_SCHEDULER_USER_ID || process.env.GRAPH_SENDER_EMAIL || null,
                tenantId: process.env.MicrosoftAppTenantId || null,
                status: 'scheduled'
            });

            return {
                ok: true,
                provider: 'teams',
                meetingId,
                joinUrl,
                message: `${teamsResult.message} Auto-join bot is enabled for this meeting. Meeting ID: ${meetingId}`
            };
        }

        return {
            ok: false,
            provider: null,
            meetingId: null,
            joinUrl: null,
            message: `Failed to schedule interview. ${teamsResult.message}`
        };
    } catch (error) {
        console.error('Calendar error:', error.message);
        return {
            ok: false,
            provider: null,
            meetingId: null,
            joinUrl: null,
            message: 'Failed to schedule interview.'
        };
    }
}

async function getInterviewResponses(candidateName) {
    try {
        const graphResult = await getInterviewResponsesFromTeams(candidateName);
        if (graphResult.ok) {
            return graphResult.message;
        }

        return `Failed to fetch interview responses. ${graphResult.message}`;
    } catch (error) {
        console.error('Calendar responses error:', error.message);
        return 'Failed to fetch interview responses.';
    }
}

async function getCandidateList() {
    const lines = candidates.map((candidate) =>
        `${candidate.id}. ${candidate.name} | ${candidate.role} | ${candidate.experience} | Email: ${candidate.email}`
    );

    return `Candidate list (${candidates.length}):\n${lines.join('\n')}`;
}

function parseTimeTo24Hour(timeText) {
    const match = timeText.trim().match(/^(\d{1,2})(?::(\d{2}))?\s*(am|pm)$/i);
    if (!match) {
        return null;
    }

    let hour = parseInt(match[1], 10);
    const minute = parseInt(match[2] || '0', 10);
    const meridiem = match[3].toLowerCase();

    if (hour < 1 || hour > 12 || minute < 0 || minute > 59) {
        return null;
    }

    if (meridiem === 'pm' && hour !== 12) {
        hour += 12;
    }
    if (meridiem === 'am' && hour === 12) {
        hour = 0;
    }

    return { hour, minute };
}

function createDateAtTime(timeText, dayOffset = 1) {
    const parsed = parseTimeTo24Hour(timeText);
    if (!parsed) {
        return null;
    }

    const targetDate = new Date();
    targetDate.setDate(targetDate.getDate() + dayOffset);
    targetDate.setHours(parsed.hour, parsed.minute, 0, 0);
    return targetDate;
}

function extractEmail(text) {
    const match = text.match(/[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}/);
    return match ? match[0] : null;
}

function parseScheduleCommand(inputText) {
    const lower = inputText.toLowerCase();
    if (!lower.includes('schedule') || !lower.includes('interview')) {
        return null;
    }

    const durationMatch = lower.match(/(\d+)\s*minute/);
    const durationMinutes = durationMatch ? parseInt(durationMatch[1], 10) : 60;

    const timeMatch = inputText.match(/at\s+(\d{1,2}(?::\d{2})?\s*(?:AM|PM|am|pm))/);
    const timeText = timeMatch ? timeMatch[1] : null;
    if (!timeText) {
        return { error: 'Missing time. Example: at 5:00 PM' };
    }

    const timezoneMatch = inputText.match(/\b([A-Za-z_]+\/[A-Za-z_]+)\b/);
    const timezone = timezoneMatch ? timezoneMatch[1] : 'Asia/Kolkata';
    const dayOffset = /\btoday\b/i.test(inputText) ? 0 : 1;

    const candidateMatch = inputText.match(
        /\b(?:with\s+(?:candidate\s+)?|candidate\s+)([A-Za-z]+(?:\s+[A-Za-z]+){0,2})(?=\s+(?:today|tomorrow|at|on|for|and|interviewer)\b|$)/i
    );
    let candidateName = candidateMatch ? candidateMatch[1].trim() : null;
    if (candidateName) {
        candidateName = candidateName.replace(/\b(today|tomorrow)\b/gi, '').trim();
    }
    if (!candidateName) {
        return { error: 'Missing candidate name. Example: schedule interview with Pranay today at 2 PM' };
    }

    const interviewerPartMatch = inputText.match(/interviewer\s+(.+)$/i);
    const interviewerEmail = interviewerPartMatch ? extractEmail(interviewerPartMatch[1]) : null;

    return {
        durationMinutes,
        timeText,
        timezone,
        dayOffset,
        candidateName,
        interviewerEmail
    };
}

function parseResponseQuery(inputText) {
    const cleanedInput = inputText
        .replace(/<[^>]*>/g, ' ')
        .replace(/\s+/g, ' ')
        .trim();
    const lower = cleanedInput.toLowerCase();

    if (!lower.includes('response') && !lower.includes('accepted') && !lower.includes('rejected')) {
        return null;
    }
    if (!lower.includes('interview') && !lower.includes('meeting')) {
        return null;
    }

    const candidateMatch = cleanedInput.match(/candidate\s+([A-Za-z]+(?:\s+[A-Za-z]+){0,2})/i) ||
        cleanedInput.match(/for\s+([A-Za-z]+(?:\s+[A-Za-z]+){0,2})/i);
    if (!candidateMatch) {
        return { error: 'Please provide candidate name. Example: interview response for candidate Aarav Sharma' };
    }

    return { candidateName: candidateMatch[1].trim() };
}

module.exports = {
    scheduleInterview,
    scheduleInterviewWithMetadata,
    getInterviewResponses,
    getCandidateList,
    findCandidateByName,
    parseScheduleCommand,
    parseResponseQuery,
    createDateAtTime
};
