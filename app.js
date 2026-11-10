(function () {
  'use strict';

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
  const detailTimestamp = document.getElementById('detailTimestamp');

  let currentMember = null;
  let currentDonations = [];

  function init() {
    updateDataSourceIndicator();
    const tableActions = document.querySelector('.table-action-btns');
    if (tableActions && profileSection) profileSection.insertAdjacentElement('afterend', tableActions);
    bindEvents();
    if (mobileInput) mobileInput.focus();
  }

  function updateDataSourceIndicator() {
    const db = window.PMUMS_DATABASE;
    if (db && db.membersByMobile) {
      const count = Object.keys(db.membersByMobile).length;
      if (dataSourceBadge) dataSourceBadge.className = 'data-source-badge badge-excel';
      if (dataSourceText) dataSourceText.textContent = `PMUMS Diary.xlsx (${count.toLocaleString('en-IN')} सदस्य)`;
    }
  }

  function bindEvents() {
    searchForm.addEventListener('submit', function (e) {
      e.preventDefault();
      handleSearch();
    });
    btnReset.addEventListener('click', resetView);
    chipButtons.forEach(btn => btn.addEventListener('click', function () {
      const mob = this.getAttribute('data-mobile');
      if (mob) {
        mobileInput.value = mob;
        handleSearch();
      }
    }));
    mobileInput.addEventListener('input', function () {
      this.value = this.value.replace(/[^0-9]/g, '');
    });
    btnPrint.addEventListener('click', () => window.print());
    btnExportCsv.addEventListener('click', exportToCsv);
    window.addEventListener('pmums:excelLoaded', updateDataSourceIndicator);
  }

  function cleanPhoneNumber(str) {
    if (!str) return '';
    const digits = str.toString().replace(/[^0-9]/g, '');
    return digits.length >= 10 ? digits.slice(-10) : digits;
  }

  function handleSearch() {
    const cleanMobile = cleanPhoneNumber(mobileInput.value.trim());
    if (cleanMobile.length !== 10) {
      showAlert('कृपया 10 अंकों का मान्य मोबाइल नंबर दर्ज करें।', 'error');
      return;
    }
    setLoading(true);
    hideAlert();
    try {
      const db = window.PMUMS_DATABASE;
      if (!db || !db.membersByMobile) {
        showAlert('डेटाबेस लोड नहीं हुआ। कृपया PMUMS Diary.xlsx फ़ाइल उपलब्ध होने की जाँच करें।', 'error');
        resetDataViews();
        return;
      }
      const member = db.membersByMobile[cleanMobile];
      if (!member) {
        showAlert(`मोबाइल नंबर "${cleanMobile}" पंजीकृत सदस्यों की सूची में नहीं मिला।`, 'error');
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

  function updateDetailTimestamp() {
    if (!detailTimestamp) return;
    detailTimestamp.textContent = new Intl.DateTimeFormat('hi-IN', {
      timeZone: 'Asia/Kolkata',
      day: '2-digit',
      month: 'long',
      year: 'numeric',
      hour: '2-digit',
      minute: '2-digit',
      second: '2-digit',
      hour12: true,
      timeZoneName: 'short'
    }).format(new Date());
  }

  function renderMemberDiary(member, donations) {
    const sortedDonations = donations.slice().sort(compareDonationDates);
    currentMember = member;
    currentDonations = sortedDonations;
    profName.textContent = member.name || '-';
    profPmumsId.textContent = member.pmumsId || '-';
    profUniqueId.textContent = member.uniqueId || 'लागू नहीं';
    profMobile.textContent = member.mobile || '-';
    profSchool.textContent = member.school || 'दर्ज नहीं';
    profBlock.textContent = member.block || '-';
    profDistrict.textContent = member.district || '-';
    profRegDate.textContent = member.regDate || '-';
    updateDetailTimestamp();

    const totalCount = sortedDonations.length;
    let totalAmt = 0;
    let lastDate = '-';
    sortedDonations.forEach(d => {
      totalAmt += (parseFloat(d.amount) || 0);
      const donationDate = normalizeDateString(d.date);
      if (donationDate && (lastDate === '-' || compareDateStrings(donationDate, lastDate) > 0)) lastDate = donationDate;
    });
    statTotalDonations.textContent = `${totalCount} सहयोग`;
    statTotalAmount.textContent = `₹${totalAmt.toLocaleString('en-IN')}`;
    statLastDate.textContent = lastDate || '-';
    renderTableRows(member.pmumsId, sortedDonations);
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
      return `${day.padStart(2, '0')}/${month.padStart(2, '0')}/${year.length === 2 ? `20${year}` : year}`;
    }
    return value;
  }

  function compareDateStrings(left, right) {
    const leftDate = parseDateString(left);
    const rightDate = parseDateString(right);
    if (!leftDate || !rightDate) return 0;
    return leftDate.getTime() - rightDate.getTime();
  }

  function compareDonationDates(left, right) {
    const leftDate = parseDateString(left.date);
    const rightDate = parseDateString(right.date);
    if (!leftDate && !rightDate) return 0;
    if (!leftDate) return 1;
    if (!rightDate) return -1;
    return rightDate.getTime() - leftDate.getTime();
  }

  function parseDateString(dateValue) {
    const normalized = normalizeDateString(dateValue);
    if (!normalized) return null;
    const parts = normalized.split('/');
    if (parts.length !== 3) return null;
    const [day, month, year] = parts.map(Number);
    if ([day, month, year].some(Number.isNaN)) return null;
    const parsed = new Date(year, month - 1, day);
    if (parsed.getFullYear() !== year || parsed.getMonth() !== month - 1 || parsed.getDate() !== day) return null;
    return Number.isNaN(parsed.getTime()) ? null : parsed;
  }

  function getDonationRowKey(donation, idx) {
    const suffix = `${donation.beneficiary || 'beneficiary'}_${donation.date || 'date'}_${donation.amount || 'amount'}`;
    return `row_${idx + 1}_${suffix}`.replace(/\s+/g, '_').replace(/[^a-zA-Z0-9_]/g, '_');
  }

  function getLegacyDonationRowKey(donation, idx) {
    return `row_${idx + 1}_${donation.beneficiary || ''}`;
  }

  function renderTableRows(pmumsId, donations) {
    diaryTableBody.innerHTML = '';
    if (donations.length === 0) {
      const tr = document.createElement('tr');
      tr.innerHTML = '<td colspan="7" style="text-align:center;padding:30px;color:#64748b;">इस सदस्य के लिए पोर्टल पर अभी कोई सहयोग प्रविष्टि दर्ज नहीं है।</td>';
      diaryTableBody.appendChild(tr);
      return;
    }
    const savedNotesKey = `PMUMS_USER_NOTES_${pmumsId}`;
    let savedNotes = {};
    try { savedNotes = JSON.parse(localStorage.getItem(savedNotesKey)) || {}; } catch (e) { savedNotes = {}; }
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
        <td class="col-amt">₹${escapeHtml(d.amount)}</td>
        <td><select class="select-table-cell" data-row="${rowKey}" data-field="mode">
          <option value="">माध्यम चुनें...</option>
          <option value="PhonePe" ${savedRow.mode === 'PhonePe' ? 'selected' : ''}>PhonePe</option>
          <option value="Google Pay" ${savedRow.mode === 'Google Pay' ? 'selected' : ''}>Google Pay</option>
          <option value="Paytm" ${savedRow.mode === 'Paytm' ? 'selected' : ''}>Paytm</option>
          <option value="Net Banking" ${savedRow.mode === 'Net Banking' ? 'selected' : ''}>Net Banking</option>
          <option value="UPI Direct" ${savedRow.mode === 'UPI Direct' ? 'selected' : ''}>UPI Direct</option>
          <option value="Cash / अन्य" ${savedRow.mode === 'Cash / अन्य' ? 'selected' : ''}>Cash / अन्य</option>
        </select></td>
        <td><input type="text" class="input-table-cell" placeholder="खाता / UPI ID..." value="${escapeHtml(savedRow.account || '')}" data-row="${rowKey}" data-field="account"></td>
        <td><input type="text" class="input-table-cell" placeholder="UTR / रिमार्क..." value="${escapeHtml(savedRow.remark || '')}" data-row="${rowKey}" data-field="remark"></td>`;
      diaryTableBody.appendChild(tr);
    });
    diaryTableBody.querySelectorAll('input, select').forEach(input => {
      const save = function () { saveUserRowNote(pmumsId, this.getAttribute('data-row'), this.getAttribute('data-field'), this.value); };
      input.addEventListener('change', save);
      input.addEventListener('keyup', save);
    });
  }

  function saveUserRowNote(pmumsId, rowKey, field, val) {
    const savedNotesKey = `PMUMS_USER_NOTES_${pmumsId}`;
    let savedNotes = {};
    try { savedNotes = JSON.parse(localStorage.getItem(savedNotesKey)) || {}; } catch (e) { savedNotes = {}; }
    if (!savedNotes[rowKey]) savedNotes[rowKey] = {};
    savedNotes[rowKey][field] = val;
    try { localStorage.setItem(savedNotesKey, JSON.stringify(savedNotes)); } catch (e) { console.warn('LocalStorage save failed:', e); }
  }

  function exportToCsv() {
    const rows = document.querySelectorAll('#diaryTable tr');
    if (rows.length === 0) { alert('एक्सपोर्ट करने के लिए कोई डेटा उपलब्ध नहीं है!'); return; }
    const csv = [];
    rows.forEach(row => {
      const cols = row.querySelectorAll('td, th');
      const values = Array.from(cols, col => {
        const data = col.innerText.replace(/(\r\n|\n|\r)/gm, '').replace(/(\s+)/gm, ' ').trim();
        return `"${data.replace(/"/g, '""')}"`;
      });
      csv.push(values.join(','));
    });
    const csvFile = new Blob(['\uFEFF' + csv.join('\n')], { type: 'text/csv;charset=utf-8;' });
    const downloadLink = document.createElement('a');
    downloadLink.download = currentMember ? `PMUMS_Diary_${currentMember.pmumsId}_${currentMember.mobile}.csv` : 'PMUMS_Diary_Export.csv';
    downloadLink.href = window.URL.createObjectURL(csvFile);
    downloadLink.style.display = 'none';
    document.body.appendChild(downloadLink);
    downloadLink.click();
    document.body.removeChild(downloadLink);
    window.URL.revokeObjectURL(downloadLink.href);
  }

  function setLoading(isLoading) {
    if (isLoading) {
      btnSearch.disabled = true;
      btnSearch.innerHTML = '<span class="status-dot"></span><span>खोज जारी है...</span>';
    } else {
      btnSearch.disabled = false;
      btnSearch.innerHTML = '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="11" cy="11" r="8"></circle><line x1="21" y1="21" x2="16.65" y2="16.65"></line></svg><span>डायरी देखें</span>';
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
    return str.toString().replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#039;');
  }

  document.addEventListener('DOMContentLoaded', init);
})();
