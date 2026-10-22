'use strict';

const fs = require('fs');

const config = JSON.parse(fs.readFileSync('config.json', 'utf8'));

function formatDate(date = new Date()) {
    const day = String(date.getDate()).padStart(2, '0');
    const monthNames = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    const month = monthNames[date.getMonth()];
    const year = String(date.getFullYear()).slice(-2);
    return `${day}-${month}-${year}`;
}

function getBaseFields(dateStr) {
    return {
        field_1: config.user.name,
        field_2: config.user.email,
        field_3: dateStr,
        field_4: '',
        field_6: '',
        field_7: '',
        T1Comments: '',
        T1_x002d_ProjectSearch: '',
        T2Category: '',
        T2ProjectName: '',
        T2Hours: '',
        T2Comments: '',
        T3ProjectName: '',
        T3hours: '00:00',
        T3Comments: '',
        T4ProjectName: '',
        T4Hours: '00:00',
        T4Comments: '',
        T5ProjectName: '',
        T5Hours: '00:00',
        T5Comments: '',
        T6ProjectName: '',
        T6Hours: '00:00',
        T6Comments: '',
        T7ProjectName: '',
        T7Hours: '00:00',
        T7Comments: '',
        T8ProjectName: '',
        T8Hours: '00:00',
        T8Comments: '',
        field_8: formatDate(),
        WorkStatus: '',
        TotalEfforts: '',
    };
}

function buildPayload(dateStr, type) {
    const template = config.payloadTemplates[type];
    if (!template) throw new Error(`Unknown payload type: ${type}`);
    return { ...getBaseFields(dateStr), ...template };
}

function getAuthHeader() {
    return config.token.startsWith('Bearer ') ? config.token : `Bearer ${config.token}`;
}

// Source data contains non-breaking spaces (\u00A0) and inconsistent spacing.
function normalizeName(name) {
    return String(name).replace(/\s+/g, ' ').trim();
}

async function fetchAvailableProjects() {
    if (!config.projectsEndpoint) throw new Error('Missing "projectsEndpoint" in config.json');

    const headers = {
        'accept': 'application/json',
        'accept-language': 'pl-PL',
        'authorization': getAuthHeader(),
        'cache-control': 'no-cache, no-store',
        'x-ms-client-app-id': '/providers/Microsoft.PowerApps/apps/53efe29c-10a7-4ae2-b5bf-a353b2c9c49e',
        'x-ms-client-environment-id': '/providers/Microsoft.PowerApps/environments/default-30459df5-1e53-4d8b-a162-0ad2348546f1',
        'x-ms-client-object-id': '3be33023-6970-4a54-bd80-eb2b59fe211c',
        'x-ms-client-request-id': `${Date.now()}-3b961e96d4d6`,
        'x-ms-client-session-id': 'f7505377-e472-4c91-b7da-b8a2027032b7',
        'x-ms-client-tenant-id': '30459df5-1e53-4d8b-a162-0ad2348546f1',
        'x-ms-protocol-semantics': 'cdp',
        'x-ms-user-agent': 'PowerApps/3.26094.5 (Web Player; AppName=53efe29c-10a7-4ae2-b5bf-a353b2c9c49e)',
        'Referer': 'https://apps.powerapps.com/',
    };

    const projects = new Set();
    let url = config.projectsEndpoint;
    while (url) {
        const response = await fetch(url, { method: 'GET', headers });
        if (!response.ok) {
            throw new Error(`Failed to fetch available projects (status ${response.status}): ${await response.text()}`);
        }
        const body = await response.json();
        const items = Array.isArray(body) ? body : body.value || [];
        for (const item of items) {
            if (item.ProjectName) projects.add(normalizeName(item.ProjectName));
        }
        url = body['@odata.nextLink'] || null;
    }
    return projects;
}

async function validateTemplates() {
    const usedTypes = [...new Set(config.days.map((day) => day.payload))];
    const toCheck = usedTypes
        .map((type) => ({ type, project: (config.payloadTemplates[type] || {}).T1_x002d_ProjectSearch }))
        .filter(({ project }) => project);

    if (toCheck.length === 0) return;

    const available = await fetchAvailableProjects();
    const invalid = toCheck.filter(({ project }) => !available.has(normalizeName(project)));

    if (invalid.length > 0) {
        console.error('Available projects:');
        [...available].sort().forEach((name) => console.error(`  - ${name}`));
        const details = invalid.map(({ type, project }) => `"${type}" -> "${project}"`).join(', ');
        throw new Error(`Template project not available: ${details}. No requests were sent.`);
    }

    console.log('Template projects validated:', toCheck.map(({ project }) => project).join(', '));
}

async function sendRequest(data) {
    const token = getAuthHeader();

    const options = {
        method: 'POST',
        headers: {
            'accept': '*/*',
            'accept-language': 'pl-PL',
            'authorization': token,
            'cache-control': 'no-cache, no-store',
            'content-type': 'application/json',
            'priority': 'u=1, i',
            'sec-ch-ua': '"Not(A:Brand";v="8", "Chromium";v="144", "Google Chrome";v="144"',
            'sec-ch-ua-mobile': '?0',
            'sec-ch-ua-platform': '"macOS"',
            'sec-fetch-dest': 'empty',
            'sec-fetch-mode': 'cors',
            'sec-fetch-site': 'cross-site',
            'x-ms-client-app-id': '/providers/Microsoft.PowerApps/apps/53efe29c-10a7-4ae2-b5bf-a353b2c9c49e',
            'x-ms-client-environment-id': '/providers/Microsoft.PowerApps/environments/default-30459df5-1e53-4d8b-a162-0ad2348546f1',
            'x-ms-client-object-id': '3be33023-6970-4a54-bd80-eb2b59fe211c',
            'x-ms-client-request-id': `${Date.now()}-3b961e96d4d6`,
            'x-ms-client-session-id': 'f7505377-e472-4c91-b7da-b8a2027032b7',
            'x-ms-client-tenant-id': '30459df5-1e53-4d8b-a162-0ad2348546f1',
            'x-ms-protocol-semantics': 'cdp',
            'x-ms-user-agent': 'PowerApps/3.26074.10 (Web Player; AppName=53efe29c-10a7-4ae2-b5bf-a353b2c9c49e)',
            'Referer': 'https://apps.powerapps.com/',
        },
        body: JSON.stringify(data),
    };

    try {
        const response = await fetch(config.endpoint, options);
        console.log(`Status: ${response.status}`);
        const body = await response.text();
        console.log('Response:', body);
    } catch (err) {
        console.error('Request error:', err);
    }
}

async function run() {
    await validateTemplates();

    for (const day of config.days) {
        const { date, payload } = day;
        const builtPayload = buildPayload(date, payload);
        console.log(`Sending ${payload} for date: ${date}...`);
        // await sendRequest(builtPayload);
    }
}

run().catch((err) => {
    console.error(err.message);
    process.exit(1);
});

