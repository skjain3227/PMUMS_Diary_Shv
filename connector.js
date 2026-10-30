/**
 * ======================================================================================
 * PMUMS DIARY - GOOGLE SHEET LIVE CONNECTOR MODULE
 * ======================================================================================
 * Connects directly with user's Google Sheet via Google Apps Script Web App Endpoint.
 * Features:
 * - Live Query (searches member records on demand directly from Google Sheet)
 * - Live Data Extraction (extracts entire sheet data into active memory, no database.js needed)
 * - Dual Fetch Strategy: Standard fetch() with automatic JSONP fallback for zero CORS errors
 * - Connection health check & real-time statistics
 * ======================================================================================
 */

(function (window) {
  'use strict';

  const STORAGE_KEY_GAS_URL = 'PMUMS_GAS_URL';
  const STORAGE_KEY_CACHED_DATA = 'PMUMS_EXTRACTED_DATA';
  const STORAGE_KEY_LAST_SYNC = 'PMUMS_LAST_SYNC_TIME';

  // Default preset Google Sheet Web App if available
  let currentGasUrl = (localStorage.getItem(STORAGE_KEY_GAS_URL) || '').trim();

  // In-memory extracted database
  let liveDatabase = null;

  // Try restoring previously extracted data from localStorage if available
  try {
    const saved = localStorage.getItem(STORAGE_KEY_CACHED_DATA);
    if (saved) {
      liveDatabase = JSON.parse(saved);
      if (liveDatabase && typeof liveDatabase === 'object') {
        window.PMUMS_DATABASE = liveDatabase;
      }
    }
  } catch (e) {
    console.warn('Could not parse cached extracted data:', e);
  }

  // Ensure window.PMUMS_DATABASE exists even if database.js is omitted
  if (!window.PMUMS_DATABASE) {
    window.PMUMS_DATABASE = {
      beneficiaries: {},
      membersByMobile: {},
      donationsById: {}
    };
  }

  /**
   * Helper: Request with Fetch, fallback to JSONP on CORS/network errors
   */
  async function requestApi(url, params = {}) {
    if (!url || typeof url !== 'string' || !url.startsWith('http')) {
      throw new Error('अमान्य Google Apps Script URL (Invalid Connector URL)');
    }

    const queryParts = [];
    for (const [key, val] of Object.entries(params)) {
      if (val !== undefined && val !== null) {
        queryParts.push(`${encodeURIComponent(key)}=${encodeURIComponent(val)}`);
      }
    }

    const separator = url.includes('?') ? '&' : '?';
    const finalUrl = queryParts.length > 0 ? `${url}${separator}${queryParts.join('&')}` : url;

    // 1. Try standard modern fetch
    try {
      const response = await fetch(finalUrl, {
        method: 'GET',
        mode: 'cors',
        cache: 'no-cache'
      });

      if (response.ok) {
        const json = await response.json();
        return json;
      }
    } catch (fetchErr) {
      console.warn('Direct fetch failed, trying JSONP connector fallback:', fetchErr);
    }

    // 2. Fallback to JSONP (Guaranteed to bypass browser CORS)
    return new Promise((resolve, reject) => {
      const callbackName = 'pmums_cb_' + Math.random().toString(36).substring(2, 9);
      const jsonpUrl = `${finalUrl}${finalUrl.includes('?') ? '&' : '?'}callback=${callbackName}`;

      const script = document.createElement('script');
      script.src = jsonpUrl;
      script.async = true;

      const timeoutId = setTimeout(() => {
        cleanup();
        reject(new Error('Google Sheet से जवाब मिलने में समय अधिक लगा (Connection Timeout)'));
      }, 15000);

      function cleanup() {
        if (script.parentNode) script.parentNode.removeChild(script);
        delete window[callbackName];
        clearTimeout(timeoutId);
      }

      window[callbackName] = function (data) {
        cleanup();
        resolve(data);
      };

      script.onerror = function () {
        cleanup();
        reject(new Error('Google Sheet कनेक्टर से संपर्क नहीं हो सका। कृपया URL और इंटरनेट जांचें।'));
      };

      document.head.appendChild(script);
    });
  }

  const SheetConnector = {
    /**
     * Get configured connector URL
     */
    getUrl() {
      return currentGasUrl;
    },

    /**
     * Save new connector URL
     */
    setUrl(url) {
      currentGasUrl = (url || '').trim();
      if (currentGasUrl) {
        localStorage.setItem(STORAGE_KEY_GAS_URL, currentGasUrl);
      } else {
        localStorage.removeItem(STORAGE_KEY_GAS_URL);
      }
      this.notifyState();
      return currentGasUrl;
    },

    /**
     * Check if connector URL is set
     */
    isConnected() {
      return Boolean(currentGasUrl && currentGasUrl.length > 10);
    },

    /**
     * Test connection to Google Sheet
     */
    async testConnection(targetUrl) {
      const url = (targetUrl || currentGasUrl || '').trim();
      if (!url) {
        throw new Error('कृपया पहले Google Apps Script Web App URL दर्ज करें।');
      }

      const result = await requestApi(url, { action: 'status' });
      return result;
    },

    /**
     * Search member live directly from Google Sheet
     */
    async searchMember(cleanMobile) {
      if (!this.isConnected()) {
        throw new Error('NO_CONNECTOR_URL');
      }

      const result = await requestApi(currentGasUrl, {
        action: 'search',
        mobile: cleanMobile
      });

      return result;
    },

    /**
     * Extract full live database from Google Sheet into memory
     */
    async extractAllData(targetUrl) {
      const url = (targetUrl || currentGasUrl || '').trim();
      if (!url) {
        throw new Error('कृपया मान्य Google Apps Script URL दर्ज करें।');
      }

      const response = await requestApi(url, { action: 'extract' });

      if (response && response.success && response.data) {
        liveDatabase = response.data;
        window.PMUMS_DATABASE = liveDatabase;

        // Persist extracted data locally so subsequent page visits work immediately
        try {
          localStorage.setItem(STORAGE_KEY_CACHED_DATA, JSON.stringify(liveDatabase));
          localStorage.setItem(STORAGE_KEY_LAST_SYNC, new Date().toISOString());
        } catch (storageErr) {
          console.warn('LocalStorage limit reached for full extracted dataset, keeping in active memory:', storageErr);
        }

        window.dispatchEvent(new CustomEvent('pmums:dataSynced', { detail: response.data }));
        return response.data;
      } else {
        throw new Error(response.message || 'डेटा एक्सट्रेक्ट करने में विफलता।');
      }
    },

    /**
     * Get last sync timestamp
     */
    getLastSyncTime() {
      return localStorage.getItem(STORAGE_KEY_LAST_SYNC);
    },

    /**
     * Clear extracted data cache
     */
    clearCache() {
      localStorage.removeItem(STORAGE_KEY_CACHED_DATA);
      localStorage.removeItem(STORAGE_KEY_LAST_SYNC);
      liveDatabase = null;
    },

    /**
     * Emit connection state event
     */
    notifyState() {
      window.dispatchEvent(new CustomEvent('pmums:connectionChanged', {
        detail: {
          url: currentGasUrl,
          connected: this.isConnected()
        }
      }));
    }
  };

  window.PMUMS_CONNECTOR = SheetConnector;

})(window);
