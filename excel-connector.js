/**
 * ======================================================================================
 * PMUMS DIARY - EXCEL CONNECTOR & DATA EXTRACTOR MODULE
 * ======================================================================================
 * Directly extracts data from local/uploaded Excel spreadsheets (.xlsx / .xls).
 * Uses SheetJS (xlsx.full.min.js) for high-speed client-side parsing.
 * 
 * Capabilities:
 * 1. Auto-fetches local Excel file (e.g. 'PMUMS Diary.xlsx', 'diary.xlsx', etc.) if placed in folder
 * 2. Drag & Drop or File Picker for user to provide any updated Excel file directly in the browser
 * 3. Extracts 'Our_Mem', 'Donation', and 'donate' sheets directly into active memory
 * 4. Caches parsed dataset into localStorage/IndexedDB for instant zero-latency searches
 * ======================================================================================
 */

(function (window) {
  'use strict';

  const STORAGE_KEY_EXCEL_DATA = 'PMUMS_EXCEL_EXTRACTED_DATA';
  const STORAGE_KEY_EXCEL_META = 'PMUMS_EXCEL_META_INFO';

  // Common filenames to automatically check if present in project folder
  const AUTO_LOAD_CANDIDATES = [
    './PMUMS Diary.xlsx',
    './PMUMS_Diary.xlsx',
    './diary.xlsx',
    './data.xlsx'
  ];

  let currentExcelMeta = null;

  // Try restoring previously extracted data from localStorage if available
  try {
    const savedMeta = localStorage.getItem(STORAGE_KEY_EXCEL_META);
    const savedData = localStorage.getItem(STORAGE_KEY_EXCEL_DATA);
    if (savedMeta && savedData) {
      currentExcelMeta = JSON.parse(savedMeta);
      const parsedData = JSON.parse(savedData);
      if (parsedData && parsedData.membersByMobile) {
        window.PMUMS_DATABASE = parsedData;
      }
    }
  } catch (e) {
    console.warn('Could not restore cached Excel data:', e);
  }

  // Ensure window.PMUMS_DATABASE is always defined
  if (!window.PMUMS_DATABASE) {
    window.PMUMS_DATABASE = {
      beneficiaries: {},
      membersByMobile: {},
      donationsById: {}
    };
  }

  /**
   * Standardize mobile numbers to last 10 digits
   */
  function cleanPhoneNumber(val) {
    if (val == null) return '';
    let str = '';
    if (typeof val === 'number') {
      str = Math.floor(val).toString();
    } else {
      str = val.toString();
    }
    const digits = str.replace(/[^0-9]/g, '');
    return digits.length >= 10 ? digits.slice(-10) : digits;
  }

  /**
   * Format Excel dates or serial numbers to dd/mm/yyyy
   */
  function formatExcelDate(raw) {
    if (!raw) return '';
    if (raw instanceof Date) {
      const d = String(raw.getDate()).padStart(2, '0');
      const m = String(raw.getMonth() + 1).padStart(2, '0');
      const y = raw.getFullYear();
      return `${d}/${m}/${y}`;
    }

    if (typeof raw === 'number' && raw > 30000 && raw < 70000) {
      const date = new Date(Math.round((raw - 25569) * 86400 * 1000));
      const d = String(date.getUTCDate()).padStart(2, '0');
      const m = String(date.getUTCMonth() + 1).padStart(2, '0');
      const y = date.getUTCFullYear();
      return `${d}/${m}/${y}`;
    }

    const str = raw.toString().trim();

    // Check if ISO format YYYY-MM-DD
    if (/^\d{4}-\d{2}-\d{2}/.test(str)) {
      const parts = str.split('T')[0].split('-');
      return `${parts[2]}/${parts[1]}/${parts[0]}`;
    }

    return str;
  }

  /**
   * Case-insensitive sheet finder
   */
  function findSheet(workbook, targetNames) {
    const sheetNames = workbook.SheetNames || [];
    for (const target of targetNames) {
      const exactMatch = sheetNames.find(n => n.trim().toLowerCase() === target.toLowerCase());
      if (exactMatch) return workbook.Sheets[exactMatch];
    }
    for (const target of targetNames) {
      const partialMatch = sheetNames.find(n => n.trim().toLowerCase().includes(target.toLowerCase()));
      if (partialMatch) return workbook.Sheets[partialMatch];
    }
    return null;
  }

  /**
   * Core Excel Parser: Parses ArrayBuffer into PMUMS database structure
   */
  function parseExcelArrayBuffer(arrayBuffer, sourceName = 'PMUMS Diary.xlsx') {
    if (typeof XLSX === 'undefined') {
      throw new Error('SheetJS (XLSX) लाइब्रेरी लोड नहीं हुई। कृपया xlsx.full.min.js की जांच करें।');
    }

    const workbook = XLSX.read(arrayBuffer, {
      type: 'array',
      cellDates: true,
      cellNF: false,
      cellText: false
    });

    // 1. Parse 'donate' sheet (Beneficiaries Master)
    const donateSheet = findSheet(workbook, ['donate', 'campaign', 'beneficiaries']);
    const beneficiaries = {};
    if (donateSheet) {
      const donateRows = XLSX.utils.sheet_to_json(donateSheet, { header: 1, defval: '' });
      // Headers at row 0, data begins from row 1 (index 1)
      for (let i = 1; i < donateRows.length; i++) {
        const row = donateRows[i];
        const benName = (row[3] || '').toString().trim();
        const amount = parseFloat(row[9]) || 170;
        if (benName) {
          beneficiaries[benName] = {
            driveNo: (row[0] || '').toString().trim(),
            benNo: (row[1] || '').toString().trim(),
            district: (row[6] || '').toString().trim(),
            month: (row[8] || '').toString().trim(),
            amount: amount
          };
        }
      }
    }

    // 2. Parse 'Our_Mem' sheet (Members Master)
    const memSheet = findSheet(workbook, ['Our_Mem', 'OurMem', 'Members', 'Member']);
    if (!memSheet) {
      throw new Error(`एक्सेल फाइल में 'Our_Mem' शीट नहीं मिली। उपलब्ध शीट्स: ${workbook.SheetNames.join(', ')}`);
    }

    const memRows = XLSX.utils.sheet_to_json(memSheet, { header: 1, defval: '' });
    const membersByMobile = {};

    // Data starts around row 4 (index 3)
    for (let i = 3; i < memRows.length; i++) {
      const row = memRows[i];
      const cleanMob = cleanPhoneNumber(row[4]); // Column E: Mobile
      if (cleanMob && cleanMob.length === 10) {
        const pmumsId = (row[1] || '').toString().trim(); // Column B
        let name = (row[0] || '').toString().trim();      // Column A
        if (!name) {
          name = `${(row[2] || '').toString().trim()} ${(row[3] || '').toString().trim()}`.trim();
        }

        membersByMobile[cleanMob] = {
          pmumsId: pmumsId,
          name: name,
          mobile: cleanMob,
          department: (row[5] || 'शिक्षा विभाग').toString().trim(),
          state: (row[6] || 'Madhya Pradesh').toString().trim(),
          division: (row[7] || 'Gwalior').toString().trim(),
          district: (row[8] || 'Shivpuri').toString().trim(),
          block: (row[9] || '').toString().trim(),
          school: (row[10] || '').toString().trim(),
          regDate: formatExcelDate(row[11]),
          uniqueId: (row[12] || '').toString().trim()
        };
      }
    }

    // 3. Parse 'Donation' sheet (Donations Ledger)
    const donSheet = findSheet(workbook, ['Donation', 'Donations', 'sahayog']);
    const donationsById = {};
    let totalDonationEntries = 0;

    if (donSheet) {
      const donRows = XLSX.utils.sheet_to_json(donSheet, { header: 1, defval: '' });
      for (let j = 3; j < donRows.length; j++) {
        const row = donRows[j];
        const pmumsId = (row[1] || '').toString().trim(); // Column B: PMUMS ID
        if (!pmumsId) continue;

        const beneficiary = (row[9] || '').toString().trim(); // Column J: Beneficiary
        const donDate = formatExcelDate(row[10]);             // Column K: Date
        const benRate = beneficiaries[beneficiary];
        const amount = benRate ? benRate.amount : 170;

        if (!donationsById[pmumsId]) {
          donationsById[pmumsId] = [];
        }

        donationsById[pmumsId].push({
          beneficiary: beneficiary,
          date: donDate,
          amount: amount
        });
        totalDonationEntries++;
      }
    }

    const database = {
      beneficiaries: beneficiaries,
      membersByMobile: membersByMobile,
      donationsById: donationsById
    };

    const meta = {
      sourceName: sourceName,
      totalMembers: Object.keys(membersByMobile).length,
      totalDonations: totalDonationEntries,
      totalBeneficiaries: Object.keys(beneficiaries).length,
      extractedAt: new Date().toISOString()
    };

    // Update active memory
    window.PMUMS_DATABASE = database;
    currentExcelMeta = meta;

    // Cache locally for instant loading
    try {
      localStorage.setItem(STORAGE_KEY_EXCEL_DATA, JSON.stringify(database));
      localStorage.setItem(STORAGE_KEY_EXCEL_META, JSON.stringify(meta));
    } catch (storageErr) {
      console.warn('LocalStorage limit for Excel cache, keeping in active memory:', storageErr);
    }

    window.dispatchEvent(new CustomEvent('pmums:excelLoaded', { detail: { meta, database } }));

    return { meta, database };
  }

  const ExcelConnector = {
    /**
     * Check if Excel data is currently loaded
     */
    isLoaded() {
      return Boolean(
        window.PMUMS_DATABASE &&
        window.PMUMS_DATABASE.membersByMobile &&
        Object.keys(window.PMUMS_DATABASE.membersByMobile).length > 0
      );
    },

    /**
     * Get current metadata
     */
    getMeta() {
      return currentExcelMeta;
    },

    /**
     * Parse File object from file input or drag-and-drop
     */
    parseFile(file) {
      return new Promise((resolve, reject) => {
        if (!file) {
          return reject(new Error('कोई फाइल नहीं चुनी गई।'));
        }

        const reader = new FileReader();
        reader.onload = (e) => {
          try {
            const buffer = e.target.result;
            const result = parseExcelArrayBuffer(buffer, file.name);
            resolve(result);
          } catch (err) {
            reject(err);
          }
        };
        reader.onerror = () => reject(new Error('फाइल पढ़ने में त्रुटि हुई।'));
        reader.readAsArrayBuffer(file);
      });
    },

    /**
     * Try auto-fetching an Excel file placed in the same folder
     */
    async autoFetchLocalExcel() {
      for (const url of AUTO_LOAD_CANDIDATES) {
        try {
          const resp = await fetch(url, { method: 'GET', cache: 'no-cache' });
          if (resp.ok) {
            const buffer = await resp.arrayBuffer();
            const fileName = url.replace('./', '');
            const result = parseExcelArrayBuffer(buffer, fileName);
            console.log(`Successfully auto-loaded Excel: ${url}`);
            return result;
          }
        } catch (e) {
          // File not present at candidate URL, try next
        }
      }
      return null;
    }
  };

  window.PMUMS_EXCEL = ExcelConnector;

  // Auto-attempt loading local Excel file once DOM is ready
  document.addEventListener('DOMContentLoaded', () => {
    ExcelConnector.autoFetchLocalExcel().catch(() => {});

    // Silent window drag & drop handler for updating or loading Excel file
    window.addEventListener('dragover', (e) => {
      e.preventDefault();
      e.stopPropagation();
    });

    window.addEventListener('drop', (e) => {
      e.preventDefault();
      e.stopPropagation();
      const files = e.dataTransfer && e.dataTransfer.files;
      if (files && files.length > 0) {
        const file = files[0];
        if (file.name.endsWith('.xlsx') || file.name.endsWith('.xls')) {
          ExcelConnector.parseFile(file).then(() => {
            console.log(`Excel file updated from drop: ${file.name}`);
          }).catch(err => {
            console.error('Failed to parse dropped Excel:', err);
          });
        }
      }
    });

    // Allow clicking the status badge to browse/update Excel if needed
    const badge = document.getElementById('dataSourceBadge');
    if (badge) {
      badge.style.cursor = 'pointer';
      badge.title = 'Click to select or drop PMUMS Diary.xlsx';
      const fileInput = document.createElement('input');
      fileInput.type = 'file';
      fileInput.accept = '.xlsx,.xls';
      fileInput.style.display = 'none';
      document.body.appendChild(fileInput);

      badge.addEventListener('click', () => fileInput.click());
      fileInput.addEventListener('change', (e) => {
        if (e.target.files && e.target.files[0]) {
          ExcelConnector.parseFile(e.target.files[0]).catch(console.error);
        }
      });
    }
  });

})(window);
