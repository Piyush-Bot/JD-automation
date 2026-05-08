const mysql = require('mysql2/promise');

let pool = null;

function getPool() {
    if (!pool) {
        pool = mysql.createPool({
            host: process.env.DB_HOST || 'localhost',
            port: parseInt(process.env.DB_PORT || '3306', 10),
            user: process.env.DB_USER || 'root',
            password: process.env.DB_PASSWORD || '',
            database: process.env.DB_NAME || '',
            waitForConnections: true,
            connectionLimit: 10,
            queueLimit: 0
        });
    }
    return pool;
}

function normalizeOptionRow(row, fallbackIndex) {
    const id = row.id ?? row.module_id ?? row.dept_id ?? row.department_id ?? fallbackIndex;
    const name = row.description ?? row.name ?? row.module_name ?? row.dept_name ?? row.department_name ?? String(id);
    return { ...row, id: String(id), name: String(name) };
}

async function getDepartments() {
    const [rows] = await getPool().query('SELECT * FROM gpm.hrm_modules where module_id = 2');
    return rows.map((row, index) => normalizeOptionRow(row, index + 1));
}

async function getRolesByDepartment(departmentId) {
    void departmentId;
    const [rows] = await getPool().query('SELECT * FROM gpm.hrm_modules where module_id = 8');
    return rows.map((row, index) => normalizeOptionRow(row, index + 1));
}

const JOB_MASTER_INDENT_SQL = `
SELECT
    a.id,
    g.description AS Department,
    h.description AS Role,
    a.name,
    a.description,
    a.roles_and_respo,
    b.name AS Originator,
    c.name AS Reviewer,
    d.name AS Approver,
    a.status,
    j.description AS Active,
    e.name AS created_by,
    f.name AS modified_by,
    a.created,
    a.modified
FROM hrm_job_master a
LEFT OUTER JOIN pcollab_members b ON b.id = a.originator
LEFT OUTER JOIN pcollab_members c ON c.id = a.reviewer
LEFT OUTER JOIN pcollab_members d ON d.id = a.approver
LEFT OUTER JOIN pcollab_members e ON e.id = a.created_by
LEFT OUTER JOIN pcollab_members f ON f.id = a.modified_by
LEFT OUTER JOIN hrm_modules g ON g.id = a.department AND g.module_id = 2
LEFT OUTER JOIN hrm_modules h ON h.id = a.role AND h.module_id = 8
LEFT OUTER JOIN hrm_modules j ON j.id = a.active AND j.module_id = 60
ORDER BY a.id DESC
`;

async function getJobMasterForIndent() {
    const limit = Math.min(Math.max(parseInt(process.env.INDENT_JOB_MASTER_LIMIT || '50', 10) || 50, 1), 200);
    const sql = `${JOB_MASTER_INDENT_SQL.trim()}\nLIMIT ?`;
    const [rows] = await getPool().query(sql, [limit]);
    return rows;
}

const JOB_MASTER_FILTERED_SQL = `
SELECT
    a.id,
    g.description AS Department,
    h.description AS Role,
    a.name,
    a.description,
    a.roles_and_respo,
    b.name AS Originator,
    c.name AS Reviewer,
    d.name AS Approver,
    a.status,
    j.description AS Active,
    e.name AS created_by,
    f.name AS modified_by,
    a.created,
    a.modified
FROM hrm_job_master a
LEFT OUTER JOIN pcollab_members b ON b.id = a.originator
LEFT OUTER JOIN pcollab_members c ON c.id = a.reviewer
LEFT OUTER JOIN pcollab_members d ON d.id = a.approver
LEFT OUTER JOIN pcollab_members e ON e.id = a.created_by
LEFT OUTER JOIN pcollab_members f ON f.id = a.modified_by
LEFT OUTER JOIN hrm_modules g ON g.id = a.department AND g.module_id = 2
LEFT OUTER JOIN hrm_modules h ON h.id = a.role AND h.module_id = 8
LEFT OUTER JOIN hrm_modules j ON j.id = a.active AND j.module_id = 60
WHERE a.role = ? AND a.department = ?
ORDER BY a.id DESC
`;

async function getJobMasterForIndentByFilters(roleId, departmentId) {
    const limit = Math.min(Math.max(parseInt(process.env.INDENT_JOB_MASTER_LIMIT || '50', 10) || 50, 1), 200);
    const sql = `${JOB_MASTER_FILTERED_SQL.trim()}\nLIMIT ?`;
    const [rows] = await getPool().query(sql, [roleId, departmentId, limit]);
    return rows;
}

async function getCollabMembers() {
    const [rows] = await getPool().query('SELECT id, firstname as first_name, lastname as last_name FROM gpm.pcollab_members');
    return rows.map((r) => {
        const id = r.id ?? r.ID;
        const first = r.first_name ?? r.firstName ?? '';
        const last = r.last_name ?? r.lastName ?? '';
        const fullName = `${first} ${last}`.trim();
        return { id: String(id), name: fullName || String(id) };
    });
}

module.exports = {
    getDepartments,
    getRolesByDepartment,
    getJobMasterForIndent,
    getJobMasterForIndentByFilters,
    getCollabMembers
};
