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

// Base repetitive fields
function getBaseFields(dateStr) {
    return {
        field_1: '',
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
        T3Hours: '00:00',
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

// Build payload for a single day
function buildPayload(dateStr, type) {
    const template = config.payloadTemplates[type];
    if (!template) throw new Error(`Unknown payload type: ${type}`);
    return { ...getBaseFields(dateStr), ...template };
}

// Send request using fetch API
async function sendRequest(data) {
    const options = {
        method: 'POST',
        headers: {
            'accept': '*/*',
            'accept-language': 'pl-PL',
            'authorization': config.token,
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
            'x-ms-client-request-id': 'beeda2fa-1191-4cf6-9503-3b961e96d4d6',
            'x-ms-client-session-id': 'f7505377-e472-4c91-b7da-b8a2027032b7',
            'x-ms-client-tenant-id': '30459df5-1e53-4d8b-a162-0ad2348546f1',
            'x-ms-protocol-semantics': 'cdp',
            'x-ms-request-method': 'POST',
            'x-ms-request-url': '/apim/sharepointonline/82103605c16148c1a2e7f83b9e8bc687/datasets/https%253A%252F%252Fdazngroup.sharepoint.com%252Fsites%252FTimesheetsTracker/tables/fe668198-afc4-4269-9a28-ae1d2ba32593/items',
            'x-ms-user-agent': 'PowerApps/3.26021.10 (Web Player; AppName=53efe29c-10a7-4ae2-b5bf-a353b2c9c49e)',
            'Referer': 'https://apps.powerapps.com/',
        },
        body: data,
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

// Iterate through config.days
async function run() {
    for (const day of config.days) {
        const { date, payload } = day;
        const builtPayload = buildPayload(date, payload);
        const data = JSON.stringify(builtPayload);
        console.log(`Sending ${payload} for date: ${date}.`);

        await sendRequest(data);
    }
}

run();


