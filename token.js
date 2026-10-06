for (const [key, value] of Object.entries(localStorage)) {
  if (
    key.includes('accesstoken') &&
    (key.includes('apihub.azure.com') || key.includes('service.flow.microsoft.com'))
  ) {
    try {
      const parsed = JSON.parse(value);
      if (parsed && parsed.secret && parsed.tokenType === 'Bearer') {
        console.log('Bearer', parsed.secret);
        break;
      }
    } catch {}
  }
}

// ==UserScript==
// @name         Power Apps Bearer Token Interceptor
// @namespace    http://tampermonkey.net/
// @version      1.0
// @description  Intercept and log Bearer tokens from fetch and XHR calls at startup
// @author       You
// @match        https://*.powerapps.com/*
// @match        https://*.apps.powerapps.com/*
// @grant        none
// @run-at       document-start
// ==/UserScript==

(function () {
    'use strict';

    let lastCapturedToken = '';

    // Helper: Parse JWT without external libraries
    function parseJwt(token) {
        try {
            const base64Url = token.split('.')[1];
            const base64 = base64Url.replace(/-/g, '+').replace(/_/g, '/');
            const jsonPayload = decodeURIComponent(
                atob(base64)
                    .split('')
                    .map(c => '%' + ('00' + c.charCodeAt(0).toString(16)).slice(-2))
                    .join('')
            );
            return JSON.parse(jsonPayload);
        } catch (e) {
            return null;
        }
    }

    // Helper: Show UI Notification Toast
    function showToast(message) {
        let toast = document.getElementById('token-toast-notification');
        if (!toast) {
            toast = document.createElement('div');
            toast.id = 'token-toast-notification';
            toast.style.cssText = `
                position: fixed;
                bottom: 20px;
                right: 20px;
                background-color: #0078d4;
                color: white;
                padding: 12px 20px;
                border-radius: 6px;
                font-family: sans-serif;
                font-size: 14px;
                font-weight: bold;
                z-index: 999999;
                box-shadow: 0 4px 12px rgba(0,0,0,0.3);
                transition: opacity 0.3s ease-in-out;
            `;
            document.body.appendChild(toast);
        }
        toast.innerText = message;
        toast.style.opacity = '1';
        setTimeout(() => {
            toast.style.opacity = '0';
        }, 4000);
    }

    // Process & Filter Headers
    function processToken(authHeader) {
        if (!authHeader || !authHeader.startsWith('Bearer ')) return;

        const rawToken = authHeader.replace('Bearer ', '').trim();
        const payload = parseJwt(rawToken);

        if (!payload) return;

        // FILTER: Target ONLY Azure API Hub / Runtime Tokens
        const isApiHub = payload.aud === 'https://apihub.azure.com' || payload.scp === 'Runtime.All';

        if (isApiHub && authHeader !== lastCapturedToken) {
            lastCapturedToken = authHeader;

            // Log cleanly in Console
            console.clear();
            console.log(
                '%c[Captured API Hub Token]',
                'background: #107c41; color: white; padding: 4px 8px; border-radius: 4px; font-weight: bold;'
            );
            console.log(authHeader);

            // Auto-copy to Clipboard
            if (typeof GM_setClipboard !== 'undefined') {
                GM_setClipboard(authHeader);
            } else if (navigator.clipboard) {
                navigator.clipboard.writeText(authHeader);
            }

            showToast('📋 API Hub Token auto-copied to clipboard!');
        }
    }

    // Intercept fetch API
    const originalFetch = window.fetch;
    window.fetch = async function (...args) {
        const [resource, config] = args;

        if (config && config.headers) {
            let authHeader = null;

            if (config.headers instanceof Headers) {
                authHeader = config.headers.get('authorization') || config.headers.get('Authorization');
            } else if (typeof config.headers === 'object') {
                authHeader = config.headers['authorization'] || config.headers['Authorization'];
            }

            processToken(authHeader);
        }

        return originalFetch.apply(this, args);
    };

    // Intercept XMLHttpRequest
    const originalXHRSetRequestHeader = XMLHttpRequest.prototype.setRequestHeader;
    XMLHttpRequest.prototype.setRequestHeader = function (header, value) {
        if (header.toLowerCase() === 'authorization') {
            processToken(value);
        }
        return originalXHRSetRequestHeader.apply(this, arguments);
    };

    console.log('[Tampermonkey] API Hub Token Extractor Active.');
})();