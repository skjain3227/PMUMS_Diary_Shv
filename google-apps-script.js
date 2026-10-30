/**
 * ======================================================================================
 * PMUMS DIARY - GOOGLE APPS SCRIPT WEB APP ENDPOINT (LIVE CONNECTOR)
 * ======================================================================================
 * 
 * यह स्क्रिप्ट आपकी Google Sheet से लाइव डेटा सीधे PMUMS Diary वेब एप्लिकेशन से जोड़ती है।
 * (This script connects your live Google Sheet directly with PMUMS Diary Web App).
 * 
 * समर्थित फीचर्स (Supported Features):
 * 1. action=search&mobile=XXXXXXXXXX : सदस्य विवरण और उनके सभी सहयोग रिकॉर्ड खोजें
 * 2. action=status (या action=test) : कनेक्शन जांचें, कुल सदस्य व दान संख्या देखें
 * 3. action=extract (या action=all) : पूरी शीट का लाइव डेटा निकालें (Extract Live Data)
 * 4. CORS और JSONP दोनों का पूर्ण समर्थन
 * 
 * --------------------------------------------------------------------------------------
 * कैसे स्थापित करें (HOW TO DEPLOY IN 2 MINUTES):
 * 1. अपनी Google Sheet खोलें (जिसमें 'Our_Mem', 'Donation', 'donate' शीट्स हैं)।
 * 2. ऊपर मेनू में जाएं: "Extensions" (एक्सटेंशन) -> "Apps Script"।
 * 3. पहले से लिखे कोड को हटाकर यह पूरा कोड पेस्ट करें।
 * 4. ऊपर "Save" (💾) आइकॉन पर क्लिक करें।
 * 5. ऊपर दायें कोने में नीले बटन "Deploy" (तैनात करें) -> "New deployment" पर क्लिक करें।
 * 6. गियर आइकॉन ⚙️ पर क्लिक करके "Web app" चुनें।
 * 7. सेटिंग्स में चुनें:
 *    - Description: PMUMS Diary Live API
 *    - Execute as: "Me" (आपके खाते से)
 *    - Who has access: "Anyone" (कोई भी व्यक्ति - आवश्यक ताकि वेबसाइट डेटा लोड कर सके)
 * 8. "Deploy" पर क्लिक करें और परमिशन दें।
 * 9. प्राप्त "Web app URL" (https://script.google.com/macros/s/.../exec) कॉपी करें।
 * 10. वेबसाइट पर "Google Sheet कनेक्टर" में यह URL पेस्ट करके सेव करें!
 * ======================================================================================
 */

function doGet(e) {
  try {
    e = e || { parameter: {} };
    var params = e.parameter || {};
    var action = (params.action || "").toLowerCase();
    var mobile = params.mobile || "";
    var callback = params.callback || "";

    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // ----------------------------------------------------------------------------------
    // ACTION 1: STATUS / TEST CONNECTION (जांचें कि शीट ठीक से कनेक्टेड है)
    // ----------------------------------------------------------------------------------
    if (action === "status" || action === "test" || action === "ping" || (!mobile && !action)) {
      var sheets = ss.getSheets().map(function(s) { return s.getName(); });
      var memSheet = ss.getSheetByName("Our_Mem");
      var donSheet = ss.getSheetByName("Donation");
      var donateSheet = ss.getSheetByName("donate");

      var memCount = memSheet ? Math.max(0, memSheet.getLastRow() - 3) : 0;
      var donCount = donSheet ? Math.max(0, donSheet.getLastRow() - 3) : 0;
      var benCount = donateSheet ? Math.max(0, donateSheet.getLastRow() - 2) : 0;

      return respondJson({
        success: true,
        message: "Google Sheet कनेक्टर सफलतापूर्वक सक्रिय है! (Connected successfully)",
        spreadsheetTitle: ss.getName(),
        sheetsFound: {
          Our_Mem: !!memSheet,
          Donation: !!donSheet,
          donate: !!donateSheet
        },
        stats: {
          totalMembers: memCount,
          totalDonations: donCount,
          totalBeneficiaries: benCount
        },
        timestamp: new Date().toISOString()
      }, callback);
    }

    // ----------------------------------------------------------------------------------
    // ACTION 2: EXTRACT ALL DATA (पूरा डेटा लाइव निकालें ताकि स्थानीय फाइल की आवश्यकता न रहे)
    // ----------------------------------------------------------------------------------
    if (action === "extract" || action === "all") {
      var fullData = extractFullDatabase(ss);
      return respondJson({
        success: true,
        message: "लाइव डेटा सफलतापूर्वक एक्सट्रेक्ट किया गया।",
        data: fullData
      }, callback);
    }

    // ----------------------------------------------------------------------------------
    // ACTION 3: SEARCH MEMBER BY MOBILE (डिफ़ॉल्ट सर्च)
    // ----------------------------------------------------------------------------------
    var cleanMobile = formatPhoneNumber(mobile);
    if (!cleanMobile || cleanMobile.length < 10) {
      return respondJson({
        success: false,
        message: "कृपया 10 अंकों का वैध मोबाइल नंबर दर्ज करें (Please enter valid 10-digit mobile number)."
      }, callback);
    }

    // 1. Beneficiary rates from 'donate'
    var beneficiaryRates = getBeneficiaryRates(ss);

    // 2. Member lookup in 'Our_Mem'
    var memSheet = ss.getSheetByName("Our_Mem");
    if (!memSheet) {
      return respondJson({
        success: false,
        message: "Google Sheet में 'Our_Mem' शीट नहीं मिली।"
      }, callback);
    }

    var memData = memSheet.getDataRange().getValues();
    var matchedMember = null;

    // Data starts at row 4 (index 3)
    for (var i = 3; i < memData.length; i++) {
      var row = memData[i];
      var phoneDigits = formatPhoneNumber(row[4]);

      if (phoneDigits === cleanMobile) {
        var regDateVal = row[11] || "";
        var formattedDate = "";
        if (regDateVal instanceof Date) {
          formattedDate = Utilities.formatDate(regDateVal, Session.getScriptTimeZone() || "GMT+05:30", "dd/MM/yyyy");
        } else {
          formattedDate = (regDateVal || "").toString().trim();
        }

        matchedMember = {
          pmumsId: (row[1] || "").toString().trim(),
          name: (row[0] || "").toString().trim() || ((row[2] || "") + " " + (row[3] || "")).trim(),
          mobile: phoneDigits,
          department: (row[5] || "शिक्षा विभाग").toString().trim(),
          state: (row[6] || "Madhya Pradesh").toString().trim(),
          division: (row[7] || "Gwalior").toString().trim(),
          district: (row[8] || "Shivpuri").toString().trim(),
          block: (row[9] || "").toString().trim(),
          school: (row[10] || "").toString().trim(),
          regDate: formattedDate,
          uniqueId: (row[12] || "").toString().trim()
        };
        break;
      }
    }

    if (!matchedMember) {
      return respondJson({
        success: false,
        message: "मोबाइल नंबर (" + cleanMobile + ") पंजीकृत सदस्यों की सूची में नहीं मिला।"
      }, callback);
    }

    // 3. Member's donations from 'Donation'
    var donSheet = ss.getSheetByName("Donation");
    var donationsList = [];

    if (donSheet) {
      var donData = donSheet.getDataRange().getValues();
      var count = 1;

      for (var j = 3; j < donData.length; j++) {
        var donRow = donData[j];
        var donMemberId = (donRow[1] || "").toString().trim();

        if (donMemberId === matchedMember.pmumsId) {
          var beneficiary = (donRow[9] || "").toString().trim();
          var rawDate = donRow[10] || "";
          var donDate = "";

          if (rawDate instanceof Date) {
            donDate = Utilities.formatDate(rawDate, Session.getScriptTimeZone() || "GMT+05:30", "dd/MM/yyyy");
          } else if (rawDate) {
            donDate = rawDate.toString().trim();
          }

          var rateInfo = beneficiaryRates[beneficiary];
          var amount = (rateInfo && rateInfo.amount) ? rateInfo.amount : (parseFloat(beneficiaryRates[beneficiary]) || 170);

          donationsList.push({
            sno: count++,
            beneficiary: beneficiary,
            date: donDate,
            amount: amount
          });
        }
      }
    }

    return respondJson({
      success: true,
      member: matchedMember,
      donations: donationsList,
      totalCount: donationsList.length,
      totalAmount: donationsList.reduce(function(sum, item) { return sum + (Number(item.amount) || 0); }, 0),
      queriedAt: new Date().toISOString()
    }, callback);

  } catch (err) {
    return respondJson({
      success: false,
      message: "त्रुटि (Error): " + err.toString()
    }, (e && e.parameter && e.parameter.callback) || "");
  }
}

/**
 * Beneficiary catalog and mandated aid from 'donate'
 */
function getBeneficiaryRates(ss) {
  var donateSheet = ss.getSheetByName("donate");
  var beneficiaryRates = {};
  if (donateSheet) {
    var donateData = donateSheet.getDataRange().getValues();
    for (var i = 2; i < donateData.length; i++) {
      var row = donateData[i];
      var benName = (row[3] || "").toString().trim();
      var amount = parseFloat(row[9]) || 170;
      if (benName) {
        beneficiaryRates[benName] = {
          driveNo: (row[0] || "").toString().trim(),
          benNo: (row[2] || "").toString().trim(),
          district: (row[4] || "").toString().trim(),
          month: (row[5] || "").toString().trim(),
          amount: amount
        };
      }
    }
  }
  return beneficiaryRates;
}

/**
 * Extracts complete relational database from Google Sheet in one request
 */
function extractFullDatabase(ss) {
  var rates = getBeneficiaryRates(ss);

  // 1. Members
  var membersByMobile = {};
  var memSheet = ss.getSheetByName("Our_Mem");
  if (memSheet) {
    var memData = memSheet.getDataRange().getValues();
    for (var i = 3; i < memData.length; i++) {
      var row = memData[i];
      var phoneDigits = formatPhoneNumber(row[4]);
      if (phoneDigits && phoneDigits.length === 10) {
        var regDateVal = row[11] || "";
        var formattedDate = "";
        if (regDateVal instanceof Date) {
          formattedDate = Utilities.formatDate(regDateVal, Session.getScriptTimeZone() || "GMT+05:30", "dd/MM/yyyy");
        } else {
          formattedDate = (regDateVal || "").toString().trim();
        }

        membersByMobile[phoneDigits] = {
          pmumsId: (row[1] || "").toString().trim(),
          name: (row[0] || "").toString().trim() || ((row[2] || "") + " " + (row[3] || "")).trim(),
          mobile: phoneDigits,
          department: (row[5] || "शिक्षा विभाग").toString().trim(),
          state: (row[6] || "Madhya Pradesh").toString().trim(),
          division: (row[7] || "Gwalior").toString().trim(),
          district: (row[8] || "Shivpuri").toString().trim(),
          block: (row[9] || "").toString().trim(),
          school: (row[10] || "").toString().trim(),
          regDate: formattedDate,
          uniqueId: (row[12] || "").toString().trim()
        };
      }
    }
  }

  // 2. Donations
  var donationsById = {};
  var donSheet = ss.getSheetByName("Donation");
  if (donSheet) {
    var donData = donSheet.getDataRange().getValues();
    for (var j = 3; j < donData.length; j++) {
      var donRow = donData[j];
      var pmumsId = (donRow[1] || "").toString().trim();
      if (!pmumsId) continue;

      var beneficiary = (donRow[9] || "").toString().trim();
      var rawDate = donRow[10] || "";
      var donDate = "";

      if (rawDate instanceof Date) {
        donDate = Utilities.formatDate(rawDate, Session.getScriptTimeZone() || "GMT+05:30", "dd/MM/yyyy");
      } else if (rawDate) {
        donDate = rawDate.toString().trim();
      }

      var rate = rates[beneficiary];
      var amount = rate ? rate.amount : 170;

      if (!donationsById[pmumsId]) {
        donationsById[pmumsId] = [];
      }
      donationsById[pmumsId].push({
        beneficiary: beneficiary,
        date: donDate,
        amount: amount
      });
    }
  }

  return {
    beneficiaries: rates,
    membersByMobile: membersByMobile,
    donationsById: donationsById,
    extractedAt: new Date().toISOString()
  };
}

/**
 * Standardize phone numbers to last 10 digits
 */
function formatPhoneNumber(val) {
  if (val == null) return "";
  var str = "";
  if (typeof val === "number") {
    str = Math.floor(val).toString();
  } else {
    str = val.toString();
  }
  var digits = str.replace(/[^0-9]/g, "");
  return digits.length >= 10 ? digits.slice(-10) : digits;
}

/**
 * Return JSON or JSONP with proper MIME types
 */
function respondJson(data, callback) {
  var jsonString = JSON.stringify(data);
  if (callback && callback.trim().length > 0) {
    return ContentService.createTextOutput(callback + "(" + jsonString + ")")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(jsonString)
    .setMimeType(ContentService.MimeType.JSON);
}
