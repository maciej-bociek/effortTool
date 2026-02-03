'use strict';

const fs = require('fs');
const https = require('https');
const { URL } = require('url');

// Links:

// https://app.powerbi.com/groups/me/apps/f9a0dc10-4c97-457b-84a2-df7a5bea679b/reports/71623059-8e47-4746-9872-f5b97fb9e4e3/5764d1d1283299710895?ctid=30459df5-1e53-4d8b-a162-0ad2348546f1&experience=power-bi

// https://apps.powerapps.com/play/e/default-30459df5-1e53-4d8b-a162-0ad2348546f1/a/53efe29c-10a7-4ae2-b5bf-a353b2c9c49e?tenantId=30459df5-1e53-4d8b-a162-0ad2348546f1&hint=456b4cc0-9896-43c5-bd8e-df9b9f7faaef&sourcetime=1744217016711#

// Run in effort tool console to get token
// for (const [key, value] of Object.entries(localStorage)) {
//   if (
//     key.includes('accesstoken') &&
//     (key.includes('apihub.azure.com') || key.includes('service.flow.microsoft.com'))
//   ) {
//     try {
//       const parsed = JSON.parse(value);
//       if (parsed && parsed.secret && parsed.tokenType === 'Bearer') {
//         console.log('Bearer', parsed.secret);
//         break;
//       }
//     } catch {}
//   }
// }

// "days": [
//     { "date": "04-Nov-25", "payload": "kayo" },
//     { "date": "05-Nov-25", "payload": "kayo" },
//     { "date": "06-Nov-25", "payload": "kayo" },
//     { "date": "07-Nov-25", "payload": "kayo" },
//     { "date": "10-Nov-25", "payload": "onLeave" },
//     { "date": "12-Nov-25", "payload": "bau" },
//     { "date": "13-Nov-25", "payload": "bau" }
//   ],

// Load config
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
        field_1: config.name,
        field_2: config.email,
        field_3: dateStr,
        field_8: formatDate(),
        T1Comments: '',
        T2DEPID: '',
        T2ProjectName: '',
        T2Comments: '',
        T3DEPID: '',
        T3ProjectName: '',
        T3hours: '00:00',
        T3Comments: '',
        T4DEPID: '',
        T4ProjectName: '',
        T4Hours: '00:00',
        T4Comments: '',
        T5DEPID: '',
        T5ProjectName: '',
        T5Hours: '00:00',
        T5Comments: '',
        T6DEPID: '',
        T6ProjectName: '',
        T6Hours: '00:00',
        T6Comments: '',
        T7DEPID: '',
        T7ProjectName: '',
        T7Hours: '00:00',
        T7Comments: '',
        T8DEPID: '',
        T8ProjectName: '',
        T8Hours: '00:00',
        T8Comments: '',
    };
}

// Build payload for a single day
function buildPayload(dateStr, type) {
    const template = config.payloadTemplates[type];
    if (!template) throw new Error(`Unknown payload type: ${type}`);
    return { ...getBaseFields(dateStr), ...template };
}

// Send HTTPS request
function sendRequest(data) {
    const urlObj = new URL(config.endpoint);
    const options = {
        hostname: urlObj.hostname,
        path: urlObj.pathname,
        method: 'POST',
        headers: {
            Authorization: config.token,
            'Content-Type': 'application/json',
            'Content-Length': Buffer.byteLength(data),
            'x-ms-user-agent': 'PowerApps/3.25094.8 (Web Player; AppName=53efe29c-10a7-4ae2-b5bf-a353b2c9c49e)',
            'sec-ch-ua-platform': '"macOS"',
            'x-ms-client-environment-id':
                '/providers/Microsoft.PowerApps/environments/default-30459df5-1e53-4d8b-a162-0ad2348546f1',
            'sec-ch-ua': '"Google Chrome";v="141", "Not?A_Brand";v="8", "Chromium";v="141"',
            'x-ms-client-app-id': '/providers/Microsoft.PowerApps/apps/53efe29c-10a7-4ae2-b5bf-a353b2c9c49e',
            'x-ms-request-url':
                '/apim/sharepointonline/82103605c16148c1a2e7f83b9e8bc687/datasets/https%253A%252F%252Fperformgroup.sharepoint.com%252Fsites%252FTimesheetsTracker/tables/fe668198-afc4-4269-9a28-ae1d2ba32593/items',
            'sec-ch-ua-mobile': '?0',
            'x-ms-client-tenant-id': '30459df5-1e53-4d8b-a162-0ad2348546f1',
            Accept: '*/*',
            'x-ms-protocol-semantics': 'cdp',
            'x-ms-client-session-id': '4181f8ad-71db-488e-95fc-8953f7b51dab',
            'Cache-Control': 'no-cache, no-store',
            Referer: '',
            'Accept-Language': 'pl-PL',
            'x-ms-request-method': 'POST',
            'x-ms-client-request-id': 'd16bb26b-0e5a-41f2-a3c9-7e085f5a4152',
            'x-ms-client-object-id': '3be33023-6970-4a54-bd80-eb2b59fe211c',
            'User-Agent':
                'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/141.0.0.0 Safari/537.36',
            DNT: '1',
        },
    };

    const req = https.request(options, (res) => {
        let body = '';
        console.log(`Status: ${res.statusCode}`);
        // eslint-disable-next-line no-return-assign
        res.on('data', (chunk) => (body += chunk));
        res.on('end', () => console.log('Response:', body));
    });

    req.on('error', (err) => console.error('Request error:', err));
    req.write(data);
    req.end();
}

// Iterate through config.days
function run() {
    for (const day of config.days) {
        const { date, payload } = day;
        const builtPayload = buildPayload(date, payload);
        const data = JSON.stringify(builtPayload);
        console.log(`Sending ${payload} for date: ${date}.`);
        sendRequest(data);
    }
}

run();
