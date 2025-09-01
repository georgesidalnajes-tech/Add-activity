function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Company Data Entry')
    .addItem('Add Existing LMC', 'showForm')
    .addItem('Add Activity', 'showactivity')
    .addItem('Labor Disputes', 'showlabordispute')
    .addItem('Success Story', 'showsuccess')
    .addItem('Search Journey', 'showsearch')
    .addToUi();
}

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('form')
    .setTitle('Company Data Entry')
    .setWidth(900)
    .setHeight(650);
}

function showForm() {
  const html = HtmlService.createHtmlOutputFromFile('form')
    .setWidth(900)
    .setHeight(650);
  SpreadsheetApp.getUi().showModalDialog(html, 'Company Data Entry');
}

function showactivity() {
  const html = HtmlService.createHtmlOutputFromFile('activity')
    .setWidth(900)
    .setHeight(650);
  SpreadsheetApp.getUi().showModalDialog(html, 'Company Data Entry');
}

function showlabordispute() {
  const html = HtmlService.createHtmlOutputFromFile('labordispute')
    .setWidth(900)
    .setHeight(650);
  SpreadsheetApp.getUi().showModalDialog(html, 'Company Data Entry');
}

function showsuccess() {
  const html = HtmlService.createHtmlOutputFromFile('success')
    .setWidth(900)
    .setHeight(650);
  SpreadsheetApp.getUi().showModalDialog(html, 'Company Data Entry');
}

function showsearch() {
  const html = HtmlService.createHtmlOutputFromFile('search')
    .setWidth(900)
    .setHeight(650);
  SpreadsheetApp.getUi().showModalDialog(html, 'Company Data Entry');
}

function doPost(e) {
  Logger.log("Received POST request with parameters: " + JSON.stringify(e.parameter));

  try {
    submitForm(e.parameter);
    return ContentService.createTextOutput(JSON.stringify({ success: true, message: "Data successfully recorded!" }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    Logger.log("Error in doPost: " + error.toString());
    return ContentService.createTextOutput(JSON.stringify({ success: false, message: "Failed to record data: " + error.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function submitForm(formData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dbSheet = ss.getSheetByName("DATABASE");
  const preFaciSheet = ss.getSheetByName("PRE-FACI/ENH");

  const startRow = 4;

  // Convert all string values in formData to uppercase
  Object.keys(formData).forEach(key => {
    const value = formData[key];
    if (typeof value === 'string') {
      formData[key] = value.toUpperCase();
    }
  });

  const rtwpbChecked = !!formData.rtwpb ? 'Y' : '';
  const oshChecked = !!formData.osh ? 'Y' : '';
  const fwChecked = !!formData.fw ? 'Y' : '';
  const isOrganized = formData.category && formData.category.toUpperCase() === 'ORGANIZED';

  const baseCompanyData = [
    formData.region || '',                     // A: Region
    formData.category || '',                   // B: Category
    '', '',                                    // C, D: (Placeholders)
    formData.companyName || '',                // E: Company Name
    formData.acronym || '',                    // F: Acronym
    formData.filedCaseYear || '',              // G: Filed Case Year
    formData.filedRfaYear || '',               // H: Filed RFA Year
    formData.ecozone || '',                    // I: Ecozone
    formData.building || '',                   // J: Building No.
    formData.street || '',                     // K: Street
    formData.barangay || '',                   // L: Barangay
    formData.municipality || '',               // M: Municipality
    formData.province || '',                   // N: Province/City
    formData.section || '',                    // O: Industry Section
    formData.division || '',                   // P: Industry Division
    formData.group || '',                      // Q: Industry Group
    formData.class || '',                      // R: Industry Class
    formData.subclass || '',                   // S: Industry Subclass
    formData.ownership || '',                  // T: Ownership/Nationality
    formData.key || '',                         
    formData.male || '',                       // U: Male Employees
    formData.female || '',                     // V: Female Employees
    formData.totalEmployees || '',             // W: Total Employees
    formData.size || '',                       // X: Size
    rtwpbChecked,                              // Y: RTWPB/PIP
    oshChecked,                                // Z: OSH
    fwChecked                                  // AA: FW
  ];

  const numLMCs = parseInt(formData.numLMCs) || 0;
  const rowsToInsert = [];

  if (numLMCs > 0) {
    for (let i = 1; i <= numLMCs; i++) {
      const currentRow = [...baseCompanyData];

      if (isOrganized) {
        currentRow.push(
          formData[`unionName_${i}`] || '',      // AB: Union Name
          formData[`bargainingUnit_${i}`] || '', // AC: Bargaining Unit
          formData[`federation_${i}`] || '',     // AD: Federation
          formData[`cbaDuration_${i}`] || '',    // AE: CBA Duration
          formData[`unionMale_${i}`] || '',      // AF: Male Members
          formData[`unionFemale_${i}`] || '',    // AG: Female Members
          formData[`unionTotal_${i}`] || ''      // AH: Total Members
        );
      } else {
        currentRow.push(...new Array(7).fill('')); // Fill with blanks if not organized
      }

      // LMC Profile (AI-BG) - Re-aligned due to skipping AT
      currentRow.push(
        // Contact Persons - Head of Company (AI-AL)
        formData[`headName_${i}`] || '',          // AI: Head Name
        formData[`headPosition_${i}`] || '',      // AJ: Head Position
        formData[`headTel_${i}`] || '',           // AK: Head Telephone
        formData[`headEmail_${i}`] || '',         // AL: Head Email

        // Contact Persons - Management (AM-AP)
        formData[`mgmtName_${i}`] || '',          // AM: Mgmt Name
        formData[`mgmtPosition_${i}`] || '',      // AN: Mgmt Position
        formData[`mgmtTel_${i}`] || '',           // AO: Mgmt Telephone
        formData[`mgmtEmail_${i}`] || '',         // AP: Mgmt Email

        // Contact Persons - Labor (AQ-AT)
        formData[`laborName_${i}`] || '',         // AQ: Labor Name
        formData[`laborPosition_${i}`] || '',     // AR: Labor Position
        formData[`laborTel_${i}`] || '',          // AS: Labor Telephone
        formData[`laborEmail_${i}`] || '',        // AT: Labor Email

        '',                                       // AU: Placeholder (empty as requested)

        // LMC General Details (AV-BA) - Shifted left by one due to AT skip
        formData[`lmcName_${i}`] || '',           // AV: LMC Name
        formData[`facilitationDate_${i}`] || '',  // AW: Date of Facilitation
        formData[`facilitatedBy_${i}`] || '',     // AX: Facilitated By
        !!formData[`orgPhilamcop_${i}`] ? 'Y' : '', // AY: Org Philamcop
        !!formData[`orgRegional_${i}`] ? (formData[`regionalText_${i}`] || '') : '', // AZ: Org Regional Association (Specify)
        !!formData[`orgOthers_${i}`] ? (formData[`othersText_${i}`] || '') : '',     // BA: Org Others (Specify)

        // Background / History of LMC (BB-BD) - Shifted left by one
        formData[`lmcReasons_${i}`] || '',        // BB: Reasons for establishing LMC
        formData[`lmcObjectives_${i}`] || '',     // BC: Objectives of LMC
        formData[`lmcVisionMission_${i}`] || '',  // BD: Vision/Mission Statement

        // Organizational Structure (BE-BG) - Shifted left by one
        !!formData[`orgSteering_${i}`] ? 'Y' : '',  // BE: Org Steering Committee
        !!formData[`orgSubcommittees_${i}`] ? 'Y' : '', // BF: Org Sub-committees
        !!formData[`orgSecretariat_${i}`] ? 'Y' : ''   // BG: Org Secretariat
      );

      // Additional LMC Facilitation Details (BH-DB) - These now start at BG, and end at DA.
      currentRow.push(
        // LMC Subcommittees (BH-BP)
        !!formData[`subcProductivity_${i}`] ? "Y" : "",     // BH: Subc Productivity
        !!formData[`subcHealthSafety_${i}`] ? "Y" : "",     // BI: Subc Health & Safety
        !!formData[`subcLivelihood_${i}`] ? "Y" : "",       // BJ: Subc Livelihood
        !!formData[`subcSportsRec_${i}`] ? "Y" : "",        // BK: Subc Sports and Recreation
        !!formData[`subcCommunityEnv_${i}`] ? "Y" : "",     // BL: Subc Community & Environment Relations
        !!formData[`subcGenderDev_${i}`] ? "Y" : "",        // BM: Subc Gender & Development
        !!formData[`subcFamilyWelfare_${i}`] ? "Y" : "",    // BN: Subc Family Welfare
        !!formData[`subcGrievance_${i}`] ? "Y" : "",        // BO: Subc Grievance Machinery/Discipline Committee
        formData[`subcOthersText_${i}`] || "",              // BP: Subc Others (Specify)

        // Representation - Management (BQ-BU)
        !!formData[`repMgmtTop_${i}`] ? "Y" : "",           // BQ: Rep Mgmt Top Management
        !!formData[`repMgmtLocalTop_${i}`] ? "Y" : "",      // BR: Rep Mgmt Local Top Management
        !!formData[`repMgmtMiddle_${i}`] ? "Y" : "",        // BS: Rep Mgmt Middle Management
        !!formData[`repMgmtHR_${i}`] ? "Y" : "",            // BT: Rep Mgmt HR Department
        formData[`repMgmtOthersText_${i}`] || "",           // BU: Rep Mgmt Others (Specify)

        // Representation - Labor (BV-BY)
        !!formData[`repLaborOfficers_${i}`] ? "Y" : "",     // BV: Rep Labor Union Officers
        !!formData[`repLaborMembers_${i}`] ? "Y" : "",      // BW: Rep Labor Union Members
        !!formData[`repLaborRankFile_${i}`] ? "Y" : "",     // BX: Rep Labor Rank and File
        formData[`repLaborOthersText_${i}`] || "",          // BY: Rep Labor Others (Specify)

        // Meetings (BZ-CA)
        formData[`mtgFrequencyText_${i}`] || "",            // BZ: Mtg Frequency (Specify)
        formData[`mtgVenueText_${i}`] || "",                // CA: Mtg Venue (Specify)

        // Information Dissemination (CB-CF)
        !!formData[`infoNewsletter_${i}`] ? "Y" : "",       // CB: Info Newsletter
        !!formData[`infoBulletin_${i}`] ? "Y" : "",         // CC: Info Post to Bulletin Boards
        !!formData[`infoInstantMsg_${i}`] ? "Y" : "",       // CD: Info Instant Messaging Platforms
        !!formData[`infoSocialMedia_${i}`] ? "Y" : "",      // CE: Info Social Media Platforms
        formData[`infoOthersText_${i}`] || "",              // CF: Info Others (Specify)

        // Feedback (CG-CJ)
        !!formData[`fbSuggestionBox_${i}`] ? "Y" : "",      // CG: FB Suggestion Box
        !!formData[`fbAgendaPrep_${i}`] ? "Y" : "",         // CH: FB Part of Agenda Preparation
        !!formData[`fbInternalSurvey_${i}`] ? "Y" : "",     // CI: FB Internal Survey
        formData[`fbOthersText_${i}`] || "",                // CJ: FB Others (Specify)

        // Decision Making (CK-CL)
        !!formData[`decConsensus_${i}`] ? "Y" : "",         // CK: Dec Consensus Decision Making
        formData[`decOthersText_${i}`] || "",               // CL: Dec Others (Specify)

        // Nature of LMC Decision (CM-CO)
        !!formData[`natFinal_${i}`] ? "Y" : "",             // CM: Nat Final
        !!formData[`natSubjectTopMgt_${i}`] ? "Y" : "",     // CN: Nat Subject to Final Approval by Top Mgt
        formData[`natOthersText_${i}`] || "",               // CO: Nat Others (Specify)

        // Implementation (CP-CR)
        !!formData[`implResolution_${i}`] ? "Y" : "",       // CP: Impl Thru Resolution
        !!formData[`implMemos_${i}`] ? "Y" : "",            // CQ: Impl Thru Memos
        formData[`implOthersText_${i}`] || "",                // CR: Impl Others (Specify)

        // Implementors (CS-CV)
        !!formData[`implrSecretariat_${i}`] ? "Y" : "",     // CS: Implr Secretariat
        !!formData[`implrSubcommittee_${i}`] ? "Y" : "",    // CT: Implr Sub-committee
        !!formData[`implrHR_${i}`] ? "Y" : "",              // CU: Implr HR
        formData[`implrOthersText_${i}`] || "",               // CV: Implr Others (Specify)

        // Monitoring (CW-CZ)
        !!formData[`monSecretariat_${i}`] ? "Y" : "",       // CW: Mon Secretariat
        !!formData[`monSubcommittee_${i}`] ? "Y" : "",      // CX: Mon Sub-committee
        !!formData[`monHR_${i}`] ? "Y" : "",                // CY: Mon HR
        formData[`monOthersText_${i}`] || "",                 // CZ: Mon Others (Specify)

        // Other Joint Committees (DA-DB)
        !!formData[`jcCBA_${i}`] ? "Y" : "",                // DA: JC CBA
        formData[`jcOthersText_${i}`] || "",              // DB: JC Others (Specify)
      );
      rowsToInsert.push(currentRow);
    }
  } else {
    const emptyLmcRow = [...baseCompanyData];
    // Total LMC related columns (AA-DB) is 79.
    emptyLmcRow.push(...new Array(79).fill(''));
    rowsToInsert.push(emptyLmcRow);
  }

  if (rowsToInsert.length > 0) {
    const lastRowDb = dbSheet.getLastRow();
    const targetRow = Math.max(lastRowDb + 1, startRow);
    dbSheet.getRange(targetRow, 1, rowsToInsert.length, rowsToInsert[0].length).setValues(rowsToInsert);
    Logger.log("Data inserted into DATABASE sheet starting at row " + targetRow);
  } else {
    Logger.log("No rows to insert into DATABASE sheet.");
  }

  let actIndex = 0;
  while (formData[`activityType${actIndex}`]) {
    const preRow = Math.max(preFaciSheet.getLastRow() + 1, startRow);
    const activity = [
      formData.region || '',                    // A: Region
      formData.category || '',                  // B: Category
      '', '',                                   // C, D: (Placeholders)
      formData.companyName || '',               // E: Company Name
       '',                                       // F: Placeholder (empty as requested)     
      formData.filedCaseYear || '',             // G: Filed Case Year
      formData.filedRfaYear || '',              // H: Filed RFA Year
      formData.size || '',                      // I: Size
      formData[`activityType${actIndex}`] || '', // J: Activity Type
      formData[`activityTitle_${actIndex}`] || '', // K: Activity Title
      formData[`activityDetails_${actIndex}`] || '',  // L: Details 
      formData[`activityDate_${actIndex}`] || '', // M: Date of Activity
      formData[`activityMale_${actIndex}`] || '',  // N: Male Participants
      formData[`activityFemale_${actIndex}`] || '',// O: Female Participants
      formData[`activityTotal_${actIndex}`] || '', // P: Total Participants
      !!formData[`activityNWPC_${actIndex}`] ? 'Y' : '', // Q: In convergence with NWPC/RTWPB
      !!formData[`activityOSHC_${actIndex}`] ? 'Y' : '', // R: In convergence with OSHC/ECC
      !!formData[`activityDOLE_${actIndex}`] ? 'Y' : ''  // S: In convergence with DOLE-RO
    ];
    preFaciSheet.getRange(preRow, 1, 1, activity.length).setValues([activity]);
    Logger.log(`Activity data inserted into PRE-FACI/ENH sheet for index ${actIndex} at row ${preRow}`);
    actIndex++;
  }
}
function getCompanyList() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DATABASE');
  const data = sheet.getRange(4, 5, sheet.getLastRow() - 3, 4).getValues(); // Starts from row 4, columns E to H

  const result = [];

  data.forEach(row => {
    const [name, address, industry, contact] = row;
    if (name) {
      result.push({
        name,
        details: {
          address,
          industry,
          contact
        }
      });
    }
  });

  return result;
}
