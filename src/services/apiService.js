const {
  getDepartments: dbGetDepartments,
  getRolesByDepartment: dbGetRolesByDepartment,
  getJobMasterForIndent: dbGetJobMasterForIndent,
  getJobMasterForIndentByFilters: dbGetJobMasterForIndentByFilters,
  getCollabMembers: dbGetCollabMembers
} = require('./dbService');

// NOTE: This service is the single place to integrate your deployed APIs.
// Keep return shapes the same as current DB functions so the bot code stays unchanged.
// When ready, replace the fallbacks below with HTTP calls to your deployed server.
// You can use process.env.API_BASE_URL and process.env.API_TOKEN if needed.

const { getGatewayAccessToken } = require('./tokenClient');

async function checkMenuEligibility(ctx) {
  if (process.env.API_BASE_URL) {
    try {
      const aadToken = await getGatewayAccessToken().catch(() => null);
      const url = `${process.env.API_BASE_URL}/bot/jd/eligibility`;
      const headers = {
        'Content-Type': 'application/json',
        ...(ctx && ctx.msAuthHeader ? { 'X-Forwarded-Authorization': ctx.msAuthHeader } : {}),
        'X-Microsoft-AppId': process.env.MicrosoftAppId || '',
        ...(aadToken ? { Authorization: `Bearer ${aadToken}` } : {}),
        ...(!aadToken && process.env.API_TOKEN ? { Authorization: `Bearer ${process.env.API_TOKEN}` } : {})
      };
      const body = {
        userId: ctx && ctx.userId,
        aadObjectId: ctx && ctx.aadObjectId,
        email: ctx && ctx.email,
        displayName: ctx && ctx.displayName,
        tenantId: ctx && ctx.tenantId,
        channelId: ctx && ctx.channelId,
        conversationId: ctx && ctx.conversationId,
        serviceUrl: ctx && ctx.serviceUrl,
        text: ctx && ctx.text
      };
      // console.log('[apiService] checkMenuEligibility body', body);
      const res = await fetch(url, {
        method: 'POST',
        headers,
        body: JSON.stringify(body)
      });
      let data = null;
      try {
        data = await res.json();
      } catch (_) {
        data = null;
      }
      return { allowed: !!(data && data.allowed), intent: data && data.intent, reason: data && data.message };
    } catch (e) {
      console.error('[apiService] checkMenuEligibility error', e && e.message ? e.message : e);
      return { allowed: false, reason: 'eligibility check failed' };
    }
  }
  return { allowed: true };
}

async function getDepartments() {
  // ADD YOUR DEPLOYED ENDPOINT HERE to fetch departments
  // Must return: Array<{ id: string, name: string }>
  return dbGetDepartments();
}

async function getRolesByDepartment(departmentId) {
  // ADD YOUR DEPLOYED ENDPOINT HERE to fetch roles (optionally filtered by departmentId)
  // Must return: Array<{ id: string, name: string }>
  return dbGetRolesByDepartment(departmentId);
}

async function getCollabMembers() {
  // ADD YOUR DEPLOYED ENDPOINT HERE to fetch collaboration members
  // Must return: Array<{ id: string, name: string }>
  return dbGetCollabMembers();
}

async function getJobMasterForIndent(limit) {
  // ADD YOUR DEPLOYED ENDPOINT HERE to fetch JD list (unfiltered)
  // Should return rows with fields expected by cards (Department, Role, name, description, roles_and_respo, Originator, Reviewer, Approver, status, Active, created_by, modified_by, created, modified)
  void limit; // limit can be passed to your API if supported
  return dbGetJobMasterForIndent();
}

async function getJobMasterForIndentByFilters(roleId, departmentId, limit) {
  // ADD YOUR DEPLOYED ENDPOINT HERE to fetch JD list filtered by role/department
  // Should return rows with fields expected by cards
  void limit; // limit can be passed to your API if supported
  return dbGetJobMasterForIndentByFilters(roleId, departmentId);
}

async function createJD(payload) {
  // ADD YOUR DEPLOYED ENDPOINT HERE to create a JD
  // Example expected payload: { deptId, roleId, originatorId, reviewerId, approverId }
  // Return { ok: true, id?: string } on success; { ok: false, error?: string } on failure
  // Fallback: pretend success without persisting
  void payload;
  return { ok: true };
}

module.exports = {
  checkMenuEligibility,
  getDepartments,
  getRolesByDepartment,
  getCollabMembers,
  getJobMasterForIndent,
  getJobMasterForIndentByFilters,
  createJD
};
