/**
 * ======================================================================================
 * PMUMS DIARY - APPLICATION LOGIC & DATA CONTROLLER
 * ======================================================================================
 * Powered directly by local 'PMUMS Diary.xlsx' (via excel-connector.js)
 * Features:
 * 1. Instant offline search by 10-digit mobile number
 * 2. Complete member profile & contribution ledger rendering
 * 3. LocalStorage persistence for user payment modes, UTR numbers, and remarks
 * 4. A4 Print / PDF & CSV Export support
 * ======================================================================================
 */

(function () {
  'use strict';

  // --- DOM Elements ---
  const searchForm = document.getElementById('searchForm');
  const mobileInput = document.getElementById('mobileInput');
  const btnSearch = document.getElementById('btnSearch');
  const btnReset = document.getElementById('btnReset');
  const chipButtons = document.querySelectorAll('.chip-btn');

  const alertBox = document.getElementById('alertBox');
  const alertMsg = document.getElementById('alertMsg');
  const emptyState = document.getElementById('emptyState');

  const profileSection = document.getElementById('profileSection');
  const statsSection = document.getElementById('statsSection');
  const tableSection = document.getElementById('tableSection');

  const profName = document.getElementById('profName');
  const profPmumsId = document.getElementById('profPmumsId');
  const profUniqueId = document.getElementById('profUniqueId');
  const profMobile = document.getElementById('profMobile');
  const profSchool = document.getElementById('profSchool');
  const profBlock = document.getElementById('profBlock');
  const profDistrict = document.getElementById('profDistrict');
  const profRegDate = document.getElementById('profRegDate');

  const statTotalDonations = document.getElementById('statTotalDonations');
  const statTotalAmount = document.getElementById('statTotalAmount');
  const statLastDate = document.getElementById('statLastDate');

  const diaryTableBody = document.getElementById('diaryTableBody');
  const btnPrint = document.getElementById('btnPrint');
  const btnExportCsv = document.getElementById('btnExportCsv');

  const dataSourceBadge = document.getElementById('dataSourceBadge');
  const dataSourceText = document.getElementById('dataSourceText');

  // --- State ---
  let currentMember = null;
  let currentDonations = [];

  // --- Initialize ---
  function init() {
    updateDataSourceIndicator();
    bindEvents();

    // Auto-focus search input
    if (mobileInput) {
      mobileInput.focus();
    }
  }

  // --- Update Data Source Indicator ---
  function updateDataSourceIndicator() {
    const db = window.PMUMS_DATABASE;
    if (db && db.membersByMobile) {
      const count = Object.keys(db.membersByMobile).length;
      if (dataSourceBadge) {
        dataSourceBadge.className = 'data-source-badge badge-excel';
      }
      if (dataSourceText) {
        dataSourceText.textContent = `PMUMS Diary.xlsx (${count.toLocaleString('en-IN')} सदस्य)`;
      }
    }
  }

  // --- Event Bindings ---
  function bindEvents() {
    // Search form submission
    searchForm.addEventListener('submit', function (e) {
      e.preventDefault();
      handleSearch();
    });

    // Reset button
    btnReset.addEventListener('click', resetView);

    // Quick chips
    chipButtons.forEach(btn => {
      btn.addEventListener('click', function () {
        const mob = this.getAttribute('data-mobile');
        if (mob) {
          mobileInput.value = mob;
          handleSearch();
        }
      });
    });

    // Auto-format phone input (allow only digits)
    mobileInput.addEventListener('input', function () {
      this.value = this.value.replace(/[^0-9]/g, '');
    });

    // Print button
    btnPrint.addEventListener('click', function () {
      window.print();
    });

    // Export CSV button
    btnExportCsv.addEventListener('click', exportToCsv);

    // Listen for Excel auto-load event
    window.addEventListener('pmums:excelLoaded', updateDataSourceIndicator);
  }

  // --- Clean Phone Number ---
  function cleanPhoneNumber(str) {
    if (!str) return '';
    const digits = str.toString().replace(/[^0-9]/g, '');
    return digits.length >= 10 ? digits.slice(-10) : digits;
  }

  // --- Search Handler ---
  function handleSearch() {
    const rawVal = mobileInput.value.trim();
    const cleanMobile = cleanPhoneNumber(rawVal);

    if (cleanMobile.length !== 10) {
      showAlert('कृपया 10 अंकों का मान्य मोबाइल नंबर दर्ज करें। (Please enter valid 10-digit number)', 'error');
      return;
    }

    setLoading(true);
    hideAlert();

    try {
      const db = window.PMUMS_DATABASE;
      if (!db || !db.membersByMobile) {
        showAlert('डेटाबेस लोड नहीं हुआ। कृपया सुनिश्चित करें कि PMUMS Diary.xlsx फ़ाइल उपस्थित है।', 'error');
        resetDataViews();
        return;
      }

      const member = db.membersByMobile[cleanMobile];

      if (!member) {
        showAlert(`मोबाइल नंबर "${cleanMobile}" पंजीकृत सदस्यों की सूची में नहीं मिला। कृपया नंबर जांचें।`, 'error');
        resetDataViews();
        return;
      }

      const donations = (db.donationsById && db.donationsById[member.pmumsId]) || [];
      renderMemberDiary(member, donations);
    } catch (err) {
      console.warn('Search error:', err);
      showAlert(`खोज में समस्या: ${err.message}`, 'error');
    } finally {
      setLoading(false);
    }
  }

  // --- Render Member Diary View ---
  function renderMemberDiary(member, donations) {
    currentMember = member;
    currentDonations = donations;

    // 1. Populate Profile
    profName.textContent = member.name || '-';
    profPmumsId.textContent = member.pmumsId || '-';
    profUniqueId.textContent = member.uniqueId || 'लागू नहीं';
    profMobile.textContent = member.mobile || '-';
    profSchool.textContent = member.school || 'दर्ज नहीं';
    profBlock.textContent = member.block || '-';
    profDistrict.textContent = member.district || '-';
    profRegDate.textContent = member.regDate || '-';

    // 2. Populate Summary Stats
    const totalCount = donations.length;
    let totalAmt = 0;
    let lastDate = '-';

    donations.forEach(d => {
      totalAmt += (parseFloat(d.amount) || 0);
      const donationDate = normalizeDateString(d.date);
      if (donationDate && (lastDate === '-' || compareDateStrings(donationDate, lastDate) > 0)) {
        lastDate = donationDate;
      }
    });

    statTotalDonations.textContent = `${totalCount} सहयोग`;
    statTotalAmount.textContent = `₹${totalAmt.toLocaleString('en-IN')}`;
    statLastDate.textContent = lastDate || '-';

    // 3. Populate Table Rows
    renderTableRows(member.pmumsId, donations);

    // 4. Update Visibility
    emptyState.style.display = 'none';
    profileSection.classList.add('active');
    statsSection.classList.add('active');
    tableSection.classList.add('active');
  }

  function normalizeDateString(dateValue) {
    if (!dateValue) return null;

    const value = String(dateValue).trim();
    if (!value || value === '-') return null;

    const isoMatch = value.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
    if (isoMatch) {
      const [, year, month, day] = isoMatch;
      return `${day.padStart(2, '0')}/${month.padStart(2, '0')}/${year}`;
    }

    const standardMatch = value.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})$/);
    if (standardMatch) {
      const [, day, month, year] = standardMatch;
      const normalizedYear = year.length === 2 ? `20${year}` : year;
      return `${day.padStart(2, '0')}/${month.padStart(2, '0')}/${normalizedYear}`;
    }

    return value;
  }

  function compareDateStrings(left, right) {
    const leftDate = parseDateString(left);
    const rightDate = parseDateString(right);

    if (!leftDate || !rightDate) {
      return 0;
    }

    return leftDate.getTime() - rightDate.getTime();
  }

  function parseDateString(dateValue) {
    const normalized = normalizeDateString(dateValue);
    if (!normalized) return null;

    const parts = normalized.split('/');
    if (parts.length !== 3) {
      return null;
    }

    const [day, month, year] = parts.map(part => Number(part));
    if ([day, month, year].some(Number.isNaN)) {
      return null;
    }

    const parsed = new Date(year, month - 1, day);
    return Number.isNaN(parsed.getTime()) ? null : parsed;
  }

  function getDonationRowKey(donation, idx) {
    const suffix = `${donation.beneficiary || 'beneficiary'}_${donation.date || 'date'}_${donation.amount || 'amount'}`;
    return `row_${idx + 1}_${suffix}`.replace(/\s+/g, '_').replace(/[^a-zA-Z0-9_]/g, '_');
  }

  function getLegacyDonationRowKey(donation, idx) {
    return `row_${idx + 1}_${donation.beneficiary || ''}`;
  }

  // --- Render Table Rows with LocalStorage Persistence for Teacher's Notes ---
  function renderTableRows(pmumsId, donations) {
    diaryTableBody.innerHTML = '';

    if (donations.length === 0) {
      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td colspan="7" style="text-align: center; padding: 30px; color: #64748b;">
          इस सदस्य के लिए पोर्टल पर अभी कोई सहयोग प्रविष्टि दर्ज नहीं है।
        </td>
      `;
      diaryTableBody.appendChild(tr);
      return;
    }

    // Load saved teacher notes from localStorage
    const savedNotesKey = `PMUMS_USER_NOTES_${pmumsId}`;
    let savedNotes = {};
    try {
      savedNotes = JSON.parse(localStorage.getItem(savedNotesKey)) || {};
    } catch (e) {
      savedNotes = {};
    }

    donations.forEach((d, idx) => {
      const sno = idx + 1;
      const rowKey = getDonationRowKey(d, idx);
      const legacyRowKey = getLegacyDonationRowKey(d, idx);
      const savedRow = savedNotes[rowKey] || savedNotes[legacyRowKey] || {};

      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td class="col-sno">${sno}</td>
        <td class="col-ben">${escapeHtml(d.beneficiary)}</td>
        <td class="col-date">${escapeHtml(d.date || '-')}</td>
        <td class="col-amt">₹${d.amount}</td>
        <td>
          <select class="select-table-cell" data-row="${rowKey}" data-field="mode">
            <option value="">माध्यम चुनें...</option>
            <option value="PhonePe" ${savedRow.mode === 'PhonePe' ? 'selected' : ''}>PhonePe</option>
            <option value="Google Pay" ${savedRow.mode === 'Google Pay' ? 'selected' : ''}>Google Pay</option>
            <option value="Paytm" ${savedRow.mode === 'Paytm' ? 'selected' : ''}>Paytm</option>
            <option value="Net Banking" ${savedRow.mode === 'Net Banking' ? 'selected' : ''}>Net Banking</option>
            <option value="UPI Direct" ${savedRow.mode === 'UPI Direct' ? 'selected' : ''}>UPI Direct</option>
            <option value="Cash / अन्य" ${savedRow.mode === 'Cash / अन्य' ? 'selected' : ''}>Cash / अन्य</option>
          </select>
        </td>
        <td>
          <input 
            type="text" 
            class="input-table-cell" 
            placeholder="खाता / UPI ID..." 
            value="${escapeHtml(savedRow.account || '')}"
            data-row="${rowKey}" 
            data-field="account"
          >
        </td>
        <td>
          <input 
            type="text" 
            class="input-table-cell" 
            placeholder="UTR / रिमार्क..." 
            value="${escapeHtml(savedRow.remark || '')}"
            data-row="${rowKey}" 
            data-field="remark"
          >
        </td>
      `;

      diaryTableBody.appendChild(tr);
    });

    // Auto-save listeners for editable inputs
    diaryTableBody.querySelectorAll('input, select').forEach(input => {
      input.addEventListener('change', function () {
        saveUserRowNote(pmumsId, this.getAttribute('data-row'), this.getAttribute('data-field'), this.value);
      });
      input.addEventListener('keyup', function () {
        saveUserRowNote(pmumsId, this.getAttribute('data-row'), this.getAttribute('data-field'), this.value);
      });
    });
  }

  // --- Auto-Save Teacher Notes to LocalStorage ---
  function saveUserRowNote(pmumsId, rowKey, field, val) {
    const savedNotesKey = `PMUMS_USER_NOTES_${pmumsId}`;
    let savedNotes = {};
    try {
      savedNotes = JSON.parse(localStorage.getItem(savedNotesKey)) || {};
    } catch (e) {
      savedNotes = {};
    }

    if (!savedNotes[rowKey]) {
      savedNotes[rowKey] = {};
    }
    savedNotes[rowKey][field] = val;

    try {
      localStorage.setItem(savedNotesKey, JSON.stringify(savedNotes));
    } catch (e) {
      console.warn('LocalStorage save failed:', e);
    }
  }

  // --- Export CSV Handler ---
  function exportToCsv() {
    let csv = [];
    let rows = document.querySelectorAll("#diaryTable tr");

    if (rows.length === 0) {
        alert("एक्सपोर्ट करने के लिए कोई डेटा उपलब्ध नहीं है!");
        return;
    }

    for (let i = 0; i < rows.length; i++) {
        let row = [], cols = rows[i].querySelectorAll("td, th");
        
        for (let j = 0; j < cols.length; j++) {
            // टेक्स्ट से अतिरिक्त स्पेसेस और लाइन ब्रेक साफ़ करें
            let data = cols[j].innerText.replace(/(\r\n|\n|\r)/gm, "").replace(/(\s+)/gm, " ").trim();
            // डबल कोट्स को सुरक्षित रखें ताकि CSV फॉर्मेट खराब न हो
            row.push('"' + data.replace(/"/g, '""') + '"');
        }
        csv.push(row.join(","));
    }

    // Excel या अन्य सॉफ्टवेयर में हिंदी फॉन्ट सही दिखने के लिए BOM ('\uFEFF') जोड़ें
    let csvFile = new Blob(["\uFEFF" + csv.join("\n")], { type: 'text/csv;charset=utf-8;' });
    let downloadLink = document.createElement("a");
    
    // डाउनलोड होने वाली फाइल का नाम
    let fileName = currentMember ? `PMUMS_Diary_${currentMember.pmumsId}_${currentMember.mobile}.csv` : "PMUMS_Diary_Export.csv";
    downloadLink.download = fileName;
    
    downloadLink.href = window.URL.createObjectURL(csvFile);
    downloadLink.style.display = "none";
    
    document.body.appendChild(downloadLink);
    downloadLink.click();
    document.body.removeChild(downloadLink);
}

  // --- UI Helpers ---
  function setLoading(isLoading) {
    if (isLoading) {
      btnSearch.disabled = true;
      btnSearch.innerHTML = `
        <span class="status-dot"></span>
        <span>खोज जारी है...</span>
      `;
    } else {
      btnSearch.disabled = false;
      btnSearch.innerHTML = `
        <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
          <circle cx="11" cy="11" r="8"></circle>
          <line x1="21" y1="21" x2="16.65" y2="16.65"></line>
        </svg>
        <span>डायरी देखें</span>
      `;
    }
  }

  function showAlert(msg, type = 'error') {
    alertMsg.textContent = msg;
    alertBox.className = `alert-box alert-${type} active`;
  }

  function hideAlert() {
    alertBox.className = 'alert-box';
    alertBox.classList.remove('active');
  }

  function resetView() {
    mobileInput.value = '';
    hideAlert();
    resetDataViews();
    mobileInput.focus();
  }

  function resetDataViews() {
    currentMember = null;
    currentDonations = [];
    profileSection.classList.remove('active');
    statsSection.classList.remove('active');
    tableSection.classList.remove('active');
    emptyState.style.display = 'block';
  }

  function escapeHtml(str) {
    if (!str) return '';
    return str.toString()
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#039;');
  }

  // Run on page load
  document.addEventListener('DOMContentLoaded', init);

})();
