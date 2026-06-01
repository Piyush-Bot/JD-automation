
const { getGatewayAccessToken } = require('./tokenClient');

// External API bases and simple in-memory token cache per email
const DB_API = process.env.DB_API_BASE_URL;
const Agent_API = process.env.API_BASE_URL;
const tokenCacheByEmail = new Map();

function formatAuthHeader(token) {
  const scheme = (process.env.API_AUTH_SCHEME || 'raw').toLowerCase();
  return scheme === 'bearer' ? `Bearer ${token}` : token;
}

async function loginAndGetToken(email) {
  const url = `${Agent_API}/login`;
  const loginBody = { user_email: email };
  let res = await fetch(url, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Accept': '*/*',
      'User-Agent': 'PostmanRuntime/7.36.3',
      'Accept-Encoding': 'gzip, deflate, br',
      'Connection': 'keep-alive'
    },
    body: JSON.stringify(loginBody)
  });
  if (!res.ok) {
    // Retry using application/x-www-form-urlencoded to mimic alternate clients
    const form = new URLSearchParams({ user_email: email });
    res = await fetch(url, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
        'Accept': '*/*',
        'User-Agent': 'PostmanRuntime/7.36.3',
        'Accept-Encoding': 'gzip, deflate, br',
        'Connection': 'keep-alive'
      },
      body: form
    });

    if (!res.ok) {
      throw new Error(`login failed: ${res.status}`);
    }
  }
  let data = null;
  try {
    data = await res.json();
  } catch (_) {
    data = null;
  }
  if (typeof data === 'string' && data) return data;
  if (data && (data.token || data.jwt || data.access_token)) {
    return data.token || data.jwt || data.access_token;
  }
  throw new Error('login: token not found in response');
}

async function getApiHeaders(email) {
  if (!Agent_API) return { 'Content-Type': 'application/json' };
  const loginEmail = (email || process.env.API_LOGIN_EMAIL_FALLBACK || '').trim();
  if (!loginEmail) throw new Error('No email available for API login. Set API_LOGIN_EMAIL_FALLBACK in .env');
  let token = tokenCacheByEmail.get(loginEmail);
  if (!token) {
    token = await loginAndGetToken(loginEmail);
    tokenCacheByEmail.set(loginEmail, token);
  }
  const headers = {
    'Content-Type': 'application/json',
    'Accept': '*/*',
    'User-Agent': 'PostmanRuntime/7.36.3',
    'Accept-Encoding': 'gzip, deflate, br',
    'Connection': 'keep-alive',
    Authorization: formatAuthHeader(token)
  };
  return headers;
}

async function apiGet(path, email) {
  const headers = await getApiHeaders(email);
  const url = `${Agent_API}${path}`;
  const res = await fetch(url, { method: 'GET', headers });
  const resText = await res.text().catch(() => null);
  if (!res.ok) {
    throw new Error(`GET ${path} failed: ${res.status}`);
  }
  let data = null;
  try { data = resText ? JSON.parse(resText) : null; } catch (_) { data = null; }
  return data;
}

async function loginForDataApi(email, flowSource) {
  if (!Agent_API) throw new Error('API_BASE_URL not configured');
  await getApiHeaders(email);
}

// function normalizeList(data, idKeys = ['id'], nameKeys = ['name']) {
//   const pick = (obj, keys) => keys.find((k) => obj && Object.prototype.hasOwnProperty.call(obj, k));
//   const arr = Array.isArray(data) ? data : (data && Array.isArray(data.records) ? data.records : []);
//   if (!Array.isArray(arr)) return [];
//   return arr
//     .map((x) => {
//       const idKey = pick(x, idKeys);
//       const nameKey = pick(x, nameKeys);
//       const id = idKey ? x[idKey] : undefined;
//       const name = nameKey ? x[nameKey] : undefined;
//       return { id: id !== undefined && id !== null ? String(id) : '', name: name !== undefined && name !== null ? String(name) : '' };
//     })
//     .filter((i) => i.id && i.name);
// }

function normalizeList(data, idKeys = ['id'], nameKeys = ['name']) {
  const pick = (obj, keys) => keys.find((k) => obj && Object.prototype.hasOwnProperty.call(obj, k));
 
  const unwrapped = data && Array.isArray(data.data) ? data.data : data;
  const arr = Array.isArray(unwrapped) ? unwrapped : (unwrapped && Array.isArray(unwrapped.records) ? unwrapped.records : []);
  
  if (!Array.isArray(arr)) return [];
  return arr
    .map((x) => {
      const idKey = pick(x, idKeys);
      const nameKey = pick(x, nameKeys);
      const id = idKey ? x[idKey] : undefined;
      const name = nameKey ? x[nameKey] : undefined;
      return { id: id !== undefined && id !== null ? String(id) : '', name: name !== undefined && name !== null ? String(name) : '' };
    })
    .filter((i) => i.id && i.name);
}

async function checkMenuEligibility(ctx) {
  if (Agent_API) {
    try {
      const aadToken = await getGatewayAccessToken().catch(() => null);
      const url = `${Agent_API}/bot/jd/eligibility`;
      const headers = {
        'Content-Type': 'application/json',
        ...(ctx && ctx.msAuthHeader ? { 'X-Forwarded-Authorization': ctx.msAuthHeader } : {}),
        'X-Microsoft-AppId': process.env.MicrosoftAppId || '',
        ...(aadToken ? { Authorization: `Bearer ${aadToken}` } : {})
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
      if (res.ok && data && data.allowed === true) {
        const loginEmail = ((ctx && ctx.email) || process.env.API_LOGIN_EMAIL_FALLBACK || '').trim();
        if (loginEmail) {
          const refreshedToken = await loginAndGetToken(loginEmail);
          tokenCacheByEmail.set(loginEmail, refreshedToken);
        }
      }
      return {
        allowed: !!(data && data.allowed),
        intent: data && data.intent,
        reason: data && data.message,
        form: data && data.form,
        semantic_prefill: data && data.semantic_prefill
      };
    } catch (e) {
      return { allowed: false, reason: 'eligibility check failed' };
    }
  }
  return { allowed: true };
}

// async function getDepartments(email, flowSource) {
//   if (!Agent_API) throw new Error('API_BASE_URL not configured');
//   const data = await apiGet('/departments', email);
//   const list = normalizeList(data, ['id', 'department_id', 'value', 'Id'], ['department', 'name', 'department_name', 'title', 'Name']);
//   return list;
// }
async function getDepartments(email, flowSource) {
  if (!Agent_API) throw new Error('API_BASE_URL not configured');
  const data = await apiGet('/departments', email);
  const list = normalizeList(data, ['id', 'department_id', 'value', 'Id'], ['department', 'name', 'department_name', 'title', 'Name']);
  return list;
}

// async function getRolesByDepartment(departmentId, email, flowSource) {
//   if (!Agent_API) throw new Error('API_BASE_URL not configured');
//   const raw = await apiGet('/roles', email);
//   let arr = Array.isArray(raw) ? raw : (raw && Array.isArray(raw.records) ? raw.records : []);
//   if (departmentId && Array.isArray(arr)) {
//     const depKey = (r) => r.departmentId || r.department_id || r.deptId || r.dept_id;
//     const filtered = arr.filter((r) => String(depKey(r)) === String(departmentId));
//     if (filtered.length) arr = filtered;
//   }
//   const list = normalizeList(arr, ['id', 'role_id', 'value', 'Id'], ['role', 'role_name', 'name', 'title', 'Name']);
//   return list;
// }
async function getRolesByDepartment(departmentId, email, flowSource) {
  if (!Agent_API) throw new Error('API_BASE_URL not configured');
  const raw = await apiGet('/roles', email);
  let arr = Array.isArray(raw?.data) ? raw.data : (Array.isArray(raw) ? raw : (raw && Array.isArray(raw.records) ? raw.records : []));
  if (departmentId && Array.isArray(arr)) {
    const depKey = (r) => r.departmentId || r.department_id || r.deptId || r.dept_id;
    const filtered = arr.filter((r) => String(depKey(r)) === String(departmentId));
    if (filtered.length) arr = filtered;
  }
  const list = normalizeList(arr, ['id', 'role_id', 'value', 'Id'], ['role', 'role_name', 'name', 'title', 'Name']);
  return list;
}

// async function getCollabMembers(email, flowSource) {
//   if (!Agent_API) throw new Error('API_BASE_URL not configured');
//   const raw = await apiGet('/originators', email);
//   const arr = Array.isArray(raw) ? raw : (raw && Array.isArray(raw.records) ? raw.records : []);
//   const out = arr.map((m) => {
//     const full = [m.firstname, m.lastname].filter(Boolean).join(' ').trim();
//     const name = full || m.full_name || m.name || '';
//     const id = m && (m.id ?? m.user_id ?? m.Id);
//     return { id: id != null ? String(id) : '', name };
//   }).filter((i) => i.id && i.name);
//   return out;
// }

async function getCollabMembers(email, flowSource) {
  if (!Agent_API) throw new Error('API_BASE_URL not configured');
  const raw = await apiGet('/originators', email);
  const arr = Array.isArray(raw?.data) ? raw.data : (Array.isArray(raw) ? raw : (raw && Array.isArray(raw.records) ? raw.records : []));
  const out = arr.map((m) => {
    const full = [m.firstname, m.lastname].filter(Boolean).join(' ').trim();
    const name = full || m.full_name || m.name || '';
    const id = m && (m.id ?? m.user_id ?? m.Id);
    return { id: id != null ? String(id) : '', name };
  }).filter((i) => i.id && i.name);
  return out;
}

async function getJdByRoleAndDept(roleId, departmentId, email, flowSource) {
  if (!Agent_API) throw new Error('API_BASE_URL not configured');
  const url = `${Agent_API}/job-description?role_id=${encodeURIComponent(roleId)}&department_id=${encodeURIComponent(departmentId)}`;
  const headers = await getApiHeaders(email);

  const res = await fetch(url, { method: 'GET', headers });
  const resText = await res.text().catch(() => null);
  if (!res.ok) {
    return { ok: false, error: `API error: ${res.status}` };
  }
  let data = null;
  try { data = JSON.parse(resText); } catch (_) { data = null; }
  const output = data && data.records && data.records[0] && data.records[0].output && data.records[0].output[0] && data.records[0].output[0].output;
  return { ok: true, output: output || {}, raw: data };
}

async function triggerJdWorkflow(payload, email, flowSource) {
  if (!Agent_API) throw new Error('API_BASE_URL not configured');
  const url = `${Agent_API}/trigger-jd-workflow`;
  const headers = await getApiHeaders(email);
  const res = await fetch(url, { method: 'POST', headers, body: JSON.stringify(payload) });
  const resText = await res.text().catch(() => null);
  if (!res.ok) {
    return { ok: false, error: `API error: ${res.status}` };
  }
  let data = null;
  try { data = JSON.parse(resText); } catch (_) { data = null; }
  const output = data && data.workflow_response && data.workflow_response.output;
  return { ok: true, output: output || {}, raw: data };
}

async function saveGeneratedJd(payload, email, flowSource) {
  if (!Agent_API) throw new Error('API_BASE_URL not configured');
  const url = `${Agent_API}/save-generated-jd`;
  const headers = await getApiHeaders(email);
  const res = await fetch(url, { method: 'POST', headers, body: JSON.stringify(payload) });
  const resText = await res.text().catch(() => null);
  if (!res.ok) {
    return { ok: false, error: `API error: ${res.status}` };
  }
  let data = null;
  try { data = JSON.parse(resText); } catch (_) { data = null; }
  return { ok: true, data };
}

async function saveUpdatedJd(payload, email, flowSource) {
  if (!Agent_API) throw new Error('API_BASE_URL not configured');
  const jdId = payload && (payload.jd_id ?? payload.jdId);
  if (!jdId) throw new Error('saveUpdatedJd: missing jd_id');
  const url = `${Agent_API}/update-generated-jd?jd_id=${encodeURIComponent(jdId)}`;
  const headers = await getApiHeaders(email);
  const bodyObj = (payload && payload.output) ? payload.output : payload;
  const res = await fetch(url, { method: 'POST', headers, body: JSON.stringify(bodyObj) });
  const resText = await res.text().catch(() => null);
  if (!res.ok) {
    return { ok: false, error: `API error: ${res.status}` };
  }
  let data = null;
  try { data = JSON.parse(resText); } catch (_) { data = null; }
  return { ok: true, data };
}

async function createJD(payload, email, flowSource) {
  const body = { email, ...payload };
  const url = `${Agent_API}/workflow-payload`;
  const headers = await getApiHeaders(email);
  const res = await fetch(url, { method: 'POST', headers, body: JSON.stringify(body) });
  const resText = await res.text().catch(() => null);
  if (!res.ok) {
    return { ok: false, error: `API error: ${res.status}` };
  }
  let data = null;
  try { data = JSON.parse(resText); } catch (_) { data = null; }
  const output = (data && data.workflow_result && data.workflow_result.response && data.workflow_result.response[0])
    || (Array.isArray(data) ? (data[0] && data[0].output) : (data && data.output));
  return { ok: true, output: output || {}, raw: data };
}

module.exports = {
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
};
