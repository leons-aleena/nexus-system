function doGet(e) {
  try {
    return HtmlService.createTemplateFromFile('CRM')
      .evaluate()
      .setTitle('Nexus CRM')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (err) {
    return HtmlService.createHtmlOutput(
      'Error loading app: ' + err.message
    );
  }
}

/**
 * Validates employee ID against external spreadsheet - ENHANCED WITH DEBUG LOGGING
 */
function validateEmployee(employeeId) {
  try {
    console.log("=== VALIDATE EMPLOYEE CALLED ===");
    console.log("Employee ID received:", employeeId);
    
    // Validate TPEID format
    if (!employeeId.startsWith('TPEID')) {
      console.log("Validation failed: Doesn't start with TPEID");
      return {
        success: false,
        error: 'Employee ID must start with TPEID'
      };
    }
    
    if (employeeId.length !== 11) {
      console.log("Validation failed: Length not 11 characters");
      return {
        success: false,
        error: 'Employee ID must be 11 characters long'
      };
    }
    
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const teamNexusSheet = spreadsheet.getSheetByName('Team Nexus - Access');
    
    if (!teamNexusSheet) {
      console.log("Validation failed: Team Nexus sheet not found");
      return {
        success: false,
        error: 'Team Nexus sheet not found'
      };
    }
    
    // Get data from columns A (names) and B (employee IDs)
    const dataRange = teamNexusSheet.getRange('A:B');
    const data = dataRange.getValues();
    
    console.log("Total rows in Team Nexus sheet:", data.length);
    
    // Find employee ID in column B and get name from column A
    for (let i = 0; i < data.length; i++) {
      const sheetEmployeeId = data[i][1] ? data[i][1].toString().trim() : '';
      const sheetEmployeeName = data[i][0] ? data[i][0].toString().trim() : '';
      
      console.log(`Row ${i}: Name="${sheetEmployeeName}", ID="${sheetEmployeeId}"`);
      
      if (sheetEmployeeId === employeeId) {
        console.log("✅ Employee found! Name:", sheetEmployeeName);
        return {
          success: true,
          employeeName: sheetEmployeeName,
          employeeId: employeeId
        };
      }
    }
    
    console.log("❌ Employee ID not found in sheet");
    return {
      success: false,
      error: 'Employee ID not found. Please check your ID and try again.'
    };
    
  } catch (error) {
    console.error('❌ Error validating employee:', error);
    return {
      success: false,
      error: 'System error. Please try again later.'
    };
  }
}

/**
 * Get user leads with enhanced error handling
 */
function getUserLeads(employeeName) {
  try {
    console.log("=== SERVER: getUserLeads() started ===");
    console.log("Received employeeName:", employeeName);

    if (!employeeName || employeeName === "undefined" || employeeName === "") {
      console.error("❌ employeeName is missing or invalid");
      return JSON.stringify({ error: "Invalid employee name: " + employeeName });
    }

    // Use correct external database ID
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Try different sheet name patterns
    const sheetName = `${employeeName} - Leads`;
    console.log("Searching for sheet:", sheetName);

    let sheet = ss.getSheetByName(sheetName);
    
    // If not found, try alternative naming patterns
    if (!sheet) {
      console.log("Sheet not found with exact name, trying alternatives...");
      const sheets = ss.getSheets();
      
      // Try to find sheet that contains employee name
      for (let i = 0; i < sheets.length; i++) {
        const currentSheetName = sheets[i].getName();
        if (currentSheetName.includes(employeeName) && currentSheetName.includes('Leads')) {
          sheet = sheets[i];
          console.log("Found matching sheet:", currentSheetName);
          break;
        }
      }
    }

    if (!sheet) {
      console.error("❌ Sheet NOT FOUND:", sheetName);
      return JSON.stringify({ error: "Leads sheet not found for: " + employeeName });
    }

    console.log("✔ Sheet found:", sheet.getName());

    const data = sheet.getDataRange().getValues();
    console.log("Total rows in sheet:", data.length);

    if (data.length <= 1) {
      console.warn("⚠ No lead data rows in this sheet.");
      return JSON.stringify([]);
    }

    const leads = [];
    const headers = data[0];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      // Skip empty rows
      if (!row[0] || row[0].toString().trim() === '') continue;

      // Convert Date objects to strings for proper serialization
      const leadAssignedTimestamp = row[6] instanceof Date ? row[6].toISOString() : (row[6] || '');
      
      const lead = {
        profileId: row[0]?.toString() || '',
        fullName: row[1]?.toString() || '',
        phoneNumber: row[2]?.toString() || '',
        city: row[3]?.toString() || '',
        gender: row[4]?.toString() || '',
        appliedCompanies: row[5] ? row[5].toString().split(",").map(x => x.trim()) : [],
        leadAssignedTimestamp: leadAssignedTimestamp,
        leadStatus: row[7]?.toString() || 'Applied',
        interviewLevel: row[8]?.toString() || 'Not Scheduled',
        lastUpdatedTimestamp: row[9]?.toString() || '',
        generalNote: row[10]?.toString() || '',
        followupCount: parseInt(row[11]) || 0,
        notInterestedReason: row[12]?.toString() || '',
        experienced: row[13]?.toString() || '',
        experienceYears: row[14]?.toString() || '',
        languages: row[15]?.toString() || '',
        qualification: row[16]?.toString() || '',
        academicCourse: row[17]?.toString() || '',
        twoWheeler: row[18]?.toString() || '',
        workMode: row[19]?.toString() || ''  
        
      };

      leads.push(lead);
    }

    console.log("✔ Leads processed:", leads.length);
    
    // Return the array directly
    return leads;

  } catch (error) {
    console.error("❌ SERVER ERROR in getUserLeads:", error);
    return JSON.stringify({ error: error.toString() });
  }
}

/**
 * Gets available lead status options from Lead Status sheet
 */
function getLeadStatusOptions() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const leadStatusSheet = spreadsheet.getSheetByName('Lead Status');
    
    if (!leadStatusSheet) {
      return ['Applied', 'Contacted', 'Qualified', 'Proposal Sent', 'Negotiation', 'Closed Won', 'Closed Lost', 'RNR - Deadlead', 'Not Interested'];
    }
    
    const dataRange = leadStatusSheet.getRange('A:A');
    const data = dataRange.getValues();
    
    const statusOptions = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().trim() !== '') {
        statusOptions.push(data[i][0].toString().trim());
      }
    }
    
    return statusOptions;
    
  } catch (error) {
    console.error('Error getting lead status options:', error);
    return ['Applied', 'Contacted', 'Qualified', 'Proposal Sent', 'Negotiation', 'Closed Won', 'Closed Lost', 'RNR - Deadlead', 'Not Interested'];
  }
}

/**
 * Get status colors from Lead Status sheet
 */
function getStatusColors() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const leadStatusSheet = spreadsheet.getSheetByName('Lead Status');
    
    if (!leadStatusSheet) {
      // Default colors if sheet not found
      return {
        'Applied': '#43A047', // New tag color
        'Contacted': '#2196F3',
        'Qualified': '#9C27B0',
        'Proposal Sent': '#FF9800',
        'Negotiation': '#FF5722',
        'Closed Won': '#4CAF50',
        'Closed Lost': '#F44336',
        'RNR - Deadlead': '#795548',
        'Not Interested': '#607D8B',
        'Interview Scheduled': '#9C27B0',
        'Call Not Connected': '#FF9800',
        'Follow-up Required': '#FFC107',
        'Callback Requested': '#03A9F4'
      };
    }
    
    const dataRange = leadStatusSheet.getDataRange();
    const data = dataRange.getValues();
    
    const statusColors = {};
    
    // Column A: Status, Column B: Requires Follow-up, Column C: Color Code
    for (let i = 1; i < data.length; i++) {
      const status = data[i][0];
      const colorCode = data[i][2];
      
      if (status && colorCode) {
        statusColors[status.toString().trim()] = colorCode.toString().trim();
      }
    }
    
    // Add default for "Applied" status with new tag color
    if (!statusColors['Applied']) {
      statusColors['Applied'] = '#43A047';
    }
    
    console.log("Status colors loaded:", statusColors);
    return statusColors;
    
  } catch (error) {
    console.error("Error getting status colors:", error);
    return {
      'Applied': '#43A047',
      'Contacted': '#2196F3',
      'Qualified': '#9C27B0',
      'Proposal Sent': '#FF9800',
      'Negotiation': '#FF5722',
      'Closed Won': '#4CAF50',
      'Closed Lost': '#F44336',
      'RNR - Deadlead': '#795548',
      'Not Interested': '#607D8B'
    };
  }
}

/**
 * Gets statuses that require follow-up from Lead Status sheet
 */
function getFollowupRequiredStatuses() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const leadStatusSheet = spreadsheet.getSheetByName('Lead Status');
    
    if (!leadStatusSheet) {
      return ['Call Not Connected', 'Follow-up Required', 'Callback Requested'];
    }
    
    const dataRange = leadStatusSheet.getRange('A:B');
    const data = dataRange.getValues();
    
    const followupStatuses = [];
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] && data[i][1] && data[i][1].toString().trim().toLowerCase() === 'yes') {
        followupStatuses.push(data[i][0].toString().trim());
      }
    }
    
    return followupStatuses;
    
  } catch (error) {
    console.error('Error getting follow-up required statuses:', error);
    return ['Call Not Connected', 'Follow-up Required', 'Callback Requested'];
  }
}

// handleFollowupsAfterStatusChange

function handleFollowupsAfterStatusChange(profileId, newStatus) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const followupSheet = ss.getSheetByName('Follow-up History');

  if (!followupSheet) return;

  const data = followupSheet.getDataRange().getValues();

  // CONFIRMED COLUMN MAP
  const PROFILE_ID_COL = 2;     // B
  const STATUS_COL = 5;         // E
  const COMPLETED_DATE_COL = 8; // H

  //  Date + Time timestamp
  const timestamp = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    'dd/MM/yyyy HH:mm:ss'
  );

  for (let i = 1; i < data.length; i++) {
    const rowProfileId = data[i][PROFILE_ID_COL - 1];
    const followupStatus = data[i][STATUS_COL - 1];

    if (String(rowProfileId) !== String(profileId)) continue;

    // Only update active follow-ups
    if (followupStatus === 'Scheduled' || followupStatus === 'Pending') {
       //  Mark follow-up as completed
      followupSheet.getRange(i + 1, STATUS_COL).setValue('Completed');
      //  Save COMPLETED date + time
      followupSheet.getRange(i + 1, COMPLETED_DATE_COL).setValue(timestamp);
    }
  }
}


/**
 * Enhanced lead status update with reason validation
 */
function updateLeadStatusWithReason(employeeName, profileId, newStatus, reason) {
  try {
    console.log("=== SERVER: updateLeadStatusWithReason() started ===");
    console.log("Employee:", employeeName);
    console.log("Profile ID:", profileId);
    console.log("New Status:", newStatus);
    console.log("Reason:", reason);

    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

    // Find the user's lead sheet
    let sheetName = `${employeeName} - Leads`;
    let leadsSheet = spreadsheet.getSheetByName(sheetName);
    
    if (!leadsSheet) {
      console.log("Sheet not found with exact name, trying alternatives...");
      const sheets = spreadsheet.getSheets();
      for (let i = 0; i < sheets.length; i++) {
        const currentSheetName = sheets[i].getName();
        if (currentSheetName.includes(employeeName) && currentSheetName.includes('Leads')) {
          leadsSheet = sheets[i];
          sheetName = currentSheetName;
          console.log("Found matching sheet:", sheetName);
          break;
        }
      }
    }

    if (!leadsSheet) {
      console.error("❌ Leads sheet not found");
      return { success: false, error: "Leads sheet not found: " + sheetName };
    }

    const data = leadsSheet.getDataRange().getValues();
    console.log("Total rows in sheet:", data.length);

    for (let i = 1; i < data.length; i++) {
      const row = data[i];

      if (row[0] && row[0].toString().trim() === profileId) {
        console.log(`✅ Found lead at row ${i + 1}`);
        
        const currentStatus = row[7] || 'Applied';
        
        // Validate RNR - Deadlead: at least 6 completed follow-ups (Column M index 12)
        if (newStatus === "RNR - Deadlead") {
          const completedFollowups = row[11] ? parseInt(row[11]) : 0;
          console.log("Completed follow-ups count:", completedFollowups);
          
          if (completedFollowups < 6) {
            return {
              success: false,
              error: "At least 6 completed follow-ups required for RNR - Deadlead. Current: " + completedFollowups
            };
          }
        }

        // Validate Not Interested: Reason required and minimum 20 characters
        if (newStatus === "Not Interested") {
          if (!reason || reason.trim() === "") {
            return { success: false, error: "Reason is required for Not Interested" };
          }
          
        }

        // Validate Not Interested: Reason required and minimum 20 characters
        if (newStatus === "Over Age") {
          if (!reason || reason.trim() === "") {
            return { success: false, error: "Reason is required for Over AGe" };
          }
          
        }

        // Update lead status in Column J (index 9)
        leadsSheet.getRange(i + 1, 8).setValue(newStatus);
        
        // Update last updated timestamp in Column H (index 7)
        leadsSheet.getRange(i + 1, 10).setValue(new Date());

        // Update related follow-ups
        //  FIX: Only complete follow-ups IF NOT required
        // ===================================================
        const followupRequiredStatuses = getFollowupRequiredStatuses();
        const requiresFollowup = followupRequiredStatuses.includes(newStatus);

        if (!requiresFollowup && newStatus !== 'Interview Scheduled'&& newStatus !== 'Waiting for CV' ) {
          handleFollowupsAfterStatusChange(profileId, newStatus);
          }
        // If follow-up IS required → DO NOTHING here
        // Follow-up will be created via saveFollowupDetails()
        // ===================================================

        // Save reason in Column K (index 10) for Not Interested
        if (newStatus === "Not Interested" && reason) {
          leadsSheet.getRange(i + 1, 13).setValue(reason);
          console.log("✅ Saved reason in column K:", reason);
        }

        if (newStatus === "Over Age" && reason) {
          leadsSheet.getRange(i + 1, 13).setValue(reason);
          console.log("✅ Saved reason in column K:", reason);
        }
        
        console.log(`✅ Updated lead ${profileId} status to ${newStatus}`);
        return { success: true };
      }
    }

    return { success: false, error: "Lead not found" };

  } catch (error) {
    console.error("❌ Error in updateLeadStatusWithReason:", error);
    return { success: false, error: error.toString() };
  }
}

/**
 * Validate status change requirements before confirmation
 */
function validateStatusChangeRequirements(profileId, newStatus, employeeName) {
  try {
    console.log("=== SERVER: validateStatusChangeRequirements ===");
    console.log("Profile ID:", profileId, "New Status:", newStatus, "Employee:", employeeName);
    
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

    // Find the user's lead sheet
    let sheetName = `${employeeName} - Leads`;
    let leadsSheet = spreadsheet.getSheetByName(sheetName);
    
    if (!leadsSheet) {
      console.log("Sheet not found with exact name, trying alternatives...");
      const sheets = spreadsheet.getSheets();
      for (let i = 0; i < sheets.length; i++) {
        const currentSheetName = sheets[i].getName();
        if (currentSheetName.includes(employeeName) && currentSheetName.includes('Leads')) {
          leadsSheet = sheets[i];
          break;
        }
      }
    }

    if (!leadsSheet) {
      return { success: false, error: "Leads sheet not found" };
    }

    const data = leadsSheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      if (row[0] && row[0].toString().trim() === profileId) {
        if (newStatus === "RNR - Deadlead") {
          const completedFollowups = row[11] ? parseInt(row[11]) : 0;
          console.log("Completed follow-ups:", completedFollowups);
          
          if (completedFollowups < 6) {
            return {
              success: false,
              error: "At least 6 completed follow-ups required for RNR - Deadlead. Current: " + completedFollowups
            };
          }
        }
        
        return { success: true };
      }
    }
    
    return { success: false, error: "Lead not found" };
    
  } catch (error) {
    console.error("Error validating status change requirements:", error);
    return { success: false, error: error.toString() };
  }
}

/**
 * Generate unique 8-character alphanumeric ID
 */
function generateFollowupId() {
  const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789';
  let result = '';
  for (let i = 0; i < 8; i++) {
    result += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return 'FU-' + result;
}

/**
 * Enhanced: Saves a scheduled follow-up AND updates lead status if needed
 */
function saveFollowupDetails(profileId, followupDateTime, followupNote, employeeId, updateStatus = false, newStatus = null) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

    // Get or create Follow-up History sheet
    let historySheet = spreadsheet.getSheetByName('Follow-up History');
    if (!historySheet) {
      historySheet = spreadsheet.insertSheet('Follow-up History');
      historySheet.getRange(1, 1, 1, 7).setValues([[
        'Timestamp',
        'Profile ID',
        'Follow-up Date and Time',
        'Follow-up Note',
        'Status',
        'Employee ID',
        'Follow-up ID'
      ]]);
    }

    // Create Follow-up ID
    const followupId = generateFollowupId();

    // Add new entry
    const lastRow = historySheet.getLastRow() + 1;
    historySheet.getRange(lastRow, 1, 1, 7).setValues([[
      new Date(),             // Timestamp
      profileId,              // Profile ID
      followupDateTime,       // Follow-up Date & Time
      followupNote,           // Follow-up Note
      "Scheduled",            // Status
      employeeId,             // Employee ID
      followupId              // Follow-up ID
    ]]);

    // If this follow-up is for a status update, also update the lead status
    if (updateStatus && newStatus) {
      // Find employee name from ID
      const teamNexusSheet = spreadsheet.getSheetByName('Team Nexus - Access');
      let employeeName = '';
      
      if (teamNexusSheet) {
        const teamData = teamNexusSheet.getDataRange().getValues();
        for (let i = 0; i < teamData.length; i++) {
          if (teamData[i][1] && teamData[i][1].toString().trim() === employeeId) {
            employeeName = teamData[i][0] ? teamData[i][0].toString().trim() : '';
            break;
          }
        }
      }
      
      if (employeeName) {
        // Update lead status
        const updateResult = updateLeadStatusWithReason(employeeName, profileId, newStatus, '');
        if (!updateResult.success) {
          console.warn("Could not update lead status after follow-up creation:", updateResult.error);
        }
      }
    }

    return {
      success: true,
      followupId: followupId
    };

  } catch (error) {
    console.error("Error saving follow-up details:", error);
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * Adds a note to a lead
 */
function addLeadNote(profileId, note) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = spreadsheet.getSheets();
    
    // Find the leads sheet for current employee
    for (const sheet of sheets) {
      const sheetName = sheet.getName();
      if (sheetName.includes(' - Leads')) {
        const dataRange = sheet.getDataRange();
        const data = dataRange.getValues();
        
        // Find the row with matching profile ID (column A)
        for (let i = 1; i < data.length; i++) {
          if (data[i][0] && data[i][0].toString().trim() === profileId) {
            // Update general note (column I) and last updated timestamp (column H)
            const currentNote = data[i][10] ? data[i][10].toString() + '\n' : '';
            const timestamp = new Date().toLocaleDateString() + ' ' + new Date().toLocaleTimeString();
            sheet.getRange(i + 1, 11).setValue(currentNote + timestamp + ': ' + note);
            
            return { success: true };
          }
        }
      }
    }
    
    return { success: false, error: 'Lead not found' };
    
  } catch (error) {
    console.error('Error adding lead note:', error);
    return { success: false, error: error.toString() };
  }
}

/**
 * Includes HTML files
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Shows dashboard loading screen - FIXED TEMPLATE PROCESSING
 */
function showDashboardLoading(employeeName, employeeId) {
  try {
    const htmlTemplate = HtmlService.createTemplateFromFile('DashboardLoading');

    htmlTemplate.greeting = getGreeting();
    htmlTemplate.employeeName = employeeName;
    htmlTemplate.employeeId = employeeId;

    const html = htmlTemplate.evaluate()
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);

    return html.getContent();
  } catch (error) {
    console.error("Error in showDashboardLoading:", error);
    return HtmlService.createHtmlOutput(`
      <h1>Error</h1><p>${error}</p>
    `).getContent();
  }
}

/**
 * Shows main dashboard - FIXED TEMPLATE PROCESSING
 */
function showMainDashboard(employeeName, employeeId) {
  try {
    const htmlTemplate = HtmlService.createTemplateFromFile('MainDashboard');

    htmlTemplate.greeting = getGreeting();
    htmlTemplate.employeeName = employeeName;
    htmlTemplate.employeeId = employeeId;

    const html = htmlTemplate.evaluate()
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);

    return html.getContent();
  } catch (error) {
    console.error("Error in showMainDashboard:", error);
    return HtmlService.createHtmlOutput(`
      <h1>Error</h1><p>${error}</p>
    `).getContent();
  }
}

/**
 * Gets greeting based on current time - ENHANCED
 */
function getGreeting() {
  const hour = new Date().getHours();
  if (hour < 12) return 'Good Morning';
  if (hour < 18) return 'Good Afternoon';
  return 'Good Evening';
}

/**
 * Save additional details to both Final Database and User's Leads sheet
 */
function saveAdditionalDetailsToBothSheets(details) {
  const resultDB = saveToFinalDatabase(details);
  const resultUser = saveToUserLeadsSheet(details);

  if (resultDB.success && resultUser.success) {
    return {
      success: true,
      message: "Details saved successfully to both sheets"
    };
  }

  return {
    success: false,
    error: "Failed to save all details",
    dbResult: resultDB,
    userResult: resultUser
  };
}

/**
 * Saves additional details to Final Database sheet with corrected column mapping
 */
function saveToFinalDatabase(additionalDetails) {
  try {
    const finalDbSheetId = '1-wXQZzxHjgJEaHZtuBFe-WqALi1vfL5Z1WCchUCnqos';
    const spreadsheet = SpreadsheetApp.openById(finalDbSheetId);
    const finalDbSheet = spreadsheet.getSheetByName('Final Database');
    
    if (!finalDbSheet) {
      return { success: false, error: 'Final Database sheet not found' };
    }
    
    // Get all data to find the matching profile
    const dataRange = finalDbSheet.getDataRange();
    const data = dataRange.getValues();
    
    let foundRow = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().trim() === additionalDetails.profileId) {
        foundRow = i;
        break;
      }
    }
    
    if (foundRow === -1) {
      return { success: false, error: 'Profile ID not found in Final Database' };
    }
    
    // Update the specific columns as per requirements
    const row = foundRow + 1;
    
    // Home Town - Column G (index 6)
    if (additionalDetails.homeTown) {
      finalDbSheet.getRange(row, 7).setValue(additionalDetails.homeTown);
    }
    
    if (additionalDetails.gender) {
      finalDbSheet.getRange(row, 6).setValue(additionalDetails.gender);
    }

    // Sales/Marketing Experience - Column H (index 7)
    if (additionalDetails.salesExperience) {
      finalDbSheet.getRange(row, 8).setValue(additionalDetails.salesExperience);
    }
    
    // Experience or Fresher - Column I (index 8)
    if (additionalDetails.salesExperience === 'No') {
      finalDbSheet.getRange(row, 9).setValue('Fresher');
    } else if (additionalDetails.salesExperience === 'Yes' && additionalDetails.experienceYears) {
      finalDbSheet.getRange(row, 9).setValue(additionalDetails.experienceYears);
    }
    
    // Languages - Column J (index 9)
    if (additionalDetails.languages) {
      finalDbSheet.getRange(row, 10).setValue(additionalDetails.languages);
    }
    
    // Highest Qualification - Column K (index 10)
    if (additionalDetails.highestQualification) {
      finalDbSheet.getRange(row, 11).setValue(additionalDetails.highestQualification);
    }
    
    // Course Type - Column M (index 12)
    if (additionalDetails.highestQualification) {
      let courseType = '';
      switch(additionalDetails.highestQualification) {
        case 'Below 10th':
        case '10th/12th Pass':
          courseType = 'High School / Intermediate / 12th / PUC';
          break;
        case 'Diploma':
          courseType = 'Diploma';
          break;
        case "Graduate (Bachelor's Degree)":
        case "Postgraduate (Master's Degree)":
        case 'Doctorate / Professional':
          courseType = additionalDetails.academicCourse || '';
          break;
      }
      if (courseType) {
        finalDbSheet.getRange(row, 13).setValue(courseType);
      }
    }
    
    // Two-wheeler Availability - Column N (index 13)
    if (additionalDetails.twoWheeler) {
      finalDbSheet.getRange(row, 14).setValue(additionalDetails.twoWheeler);
    }
    
    // Preferred Work Mode - Column O (index 14)
    if (additionalDetails.workMode) {
      finalDbSheet.getRange(row, 15).setValue(additionalDetails.workMode);
    }
    
    console.log(`Additional details saved to Final Database for profile: ${additionalDetails.profileId}`);
    return { success: true };
    
  } catch (error) {
    console.error('Error saving to Final Database:', error);
    return { success: false, error: error.toString() };
  }
}

/**
 * Save additional details to user's leads sheet
 */
function saveToUserLeadsSheet(details) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

    const sheetName = details.employeeName + " - Leads";
    const leadsSheet = spreadsheet.getSheetByName(sheetName);

    if (!leadsSheet) return { success: false, error: "User leads sheet not found" };

    const data = leadsSheet.getDataRange().getValues();

    let rowIndex = -1;

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().trim() === details.profileId) {
        rowIndex = i;
        break;
      }
    }

    if (rowIndex === -1) return { success: false, error: "Profile not found" };

    const row = rowIndex + 1;

    console.log("Saving additional details to row:", row);
    console.log("Details to save:", details);
    
    // Map form fields to sheet columns
    leadsSheet.getRange(row, 14).setValue(details.salesExperience || "");       // Column O: Experienced?
    leadsSheet.getRange(row, 15).setValue(details.experienceYears || "");       // Column P: Year of Experience
    leadsSheet.getRange(row, 16).setValue(details.languages || "");             // Column Q: Languages
    leadsSheet.getRange(row, 17).setValue(details.highestQualification || "");  // Column R: Qualification
    leadsSheet.getRange(row, 18).setValue(details.academicCourse || "");        // Column S: Academic Course
    leadsSheet.getRange(row, 19).setValue(details.twoWheeler || "");            // Column T: 2-Wheeler Access
    leadsSheet.getRange(row, 20).setValue(details.workMode || "");              // Column U: Mode of Work

    // Also update home town in Column D (4) if provided
    if (details.homeTown) {
        leadsSheet.getRange(row, 4).setValue(details.homeTown); // Column D: City/Home Town
    }

    if (details.gender) {
        leadsSheet.getRange(row, 5).setValue(details.gender); // 
    }


    return { success: true };

  } catch (error) {
    console.error("Error in saveToUserLeadsSheet:", error);
    return { success: false, error: error.toString() };
  }
}

/**
 * Get existing additional details from user's leads sheet
 */
function getAdditionalDetailsFromUserSheet(employeeName, profileId) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

    const sheetName = employeeName + " - Leads";
    const leadsSheet = spreadsheet.getSheetByName(sheetName);

    if (!leadsSheet) {
      return { success: false, error: "User leads sheet not found: " + sheetName };
    }

    const data = leadsSheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      const row = data[i];

      if (row[0] && row[0].toString().trim() === profileId) {
        console.log("Found lead for additional details:", profileId);
        console.log("Row data:", row);
        
        return {
          success: true,
          data: {          
            homeTown: row[3] || '',
            gender: row[4] || '',
            salesExperience: row[13] || '',
            experienceYears: row[14] || '',
            languages: row[15] || '',
            highestQualification: row[16] || '',
            academicCourse: row[17] || '',
            twoWheeler: row[18] || '',
            workMode: row[19] || ''
          }
        };
      }
    }

    return { success: true, data: null };

  } catch (error) {
    console.error("Error in getAdditionalDetailsFromUserSheet:", error);
    return { success: false, error: error.toString() };
  }
}

/**
 * NEW: Get user statistics for dashboard with interviews today count
 */
function getUserStats(employeeName) {
    try {
        console.log("=== SERVER: getUserStats() started ===");
        console.log("Employee name:", employeeName);

        const ss = SpreadsheetApp.getActiveSpreadsheet();

        // Find user's lead sheet
        const sheetName = `${employeeName} - Leads`;
        const leadsSheet = ss.getSheetByName(sheetName);
        
        if (!leadsSheet) {
            console.error("Leads sheet not found:", sheetName);
            return { 
                success: false, 
                error: "Leads sheet not found",
                stats: {
                    totalLeads: 0,
                    newLeads: 0,
                    interviewsToday: 0,
                    detailsShared: 0,
                    cvRequested:0,
                    leadsByStatus: {}
                }
            };
        }

        const data = leadsSheet.getDataRange().getValues();
        
        // Skip header row
        const leads = data.slice(1);
        
        let totalLeads = 0;
        let newLeads = 0;
        let interviewsToday = 0;
        let detailsShared = 0;
        let cvRequested = 0;
        let leadsByStatus = {};

        const today = new Date();
        today.setHours(0, 0, 0, 0);

        leads.forEach(row => {
            if (row[0] && row[0].toString().trim() !== '') {
                totalLeads++;
                
                const status = row[7] ? row[7].toString().trim() : 'Applied';
                
                // Count new leads (status = "Applied")
                if (status === 'Applied') {
                    newLeads++;
                }
                
                // Count details shared leads (status = "Details Shared")
                if (status === 'Details Shared') {
                    detailsShared++;
                }

                if (status === 'Waiting for CV') {
                    cvRequested++;
                }
                
                // Count interviews scheduled for today
                if (status === 'Interview Scheduled') {
                    const lastUpdated = row[9]; // Column H
                    if (lastUpdated instanceof Date) {
                        const updateDate = new Date(lastUpdated);
                        updateDate.setHours(0, 0, 0, 0);
                        if (updateDate.getTime() === today.getTime()) {
                            interviewsToday++;
                        }
                    } else if (typeof lastUpdated === 'string') {
                        try {
                            const updateDate = new Date(lastUpdated);
                            updateDate.setHours(0, 0, 0, 0);
                            if (updateDate.getTime() === today.getTime()) {
                                interviewsToday++;
                            }
                        } catch (e) {
                            console.log("Error parsing date:", lastUpdated);
                        }
                    }
                }
                
                // Count leads by status
                leadsByStatus[status] = (leadsByStatus[status] || 0) + 1;
            }
        });

        console.log("Stats calculated:", { 
            totalLeads, 
            newLeads, 
            interviewsToday, 
            detailsShared, 
            cvRequested,
            leadsByStatus 
        });

        return {
            success: true,
            stats: {
                totalLeads,
                newLeads,
                interviewsToday,
                detailsShared,
                cvRequested,
                leadsByStatus
            }
        };

    } catch (error) {
        console.error("Error in getUserStats:", error);
        return {
            success: false,
            error: error.toString(),
            stats: {
                totalLeads: 0,
                newLeads: 0,
                interviewsToday: 0,
                detailsShared: 0,
                leadsByStatus: {}
            }
        };
    }
}

function getFollowupStats(employeeName, filterType = 'today') {
  try {
    console.log("=== SERVER: getFollowupStats() started ===");
    console.log("Employee:", employeeName);
    console.log("Filter type:", filterType);

    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Get employee ID from Team Nexus - Access sheet
    const teamNexusSheet = ss.getSheetByName("Team Nexus - Access");
    if (!teamNexusSheet) {
      return { success: false, error: "Team Nexus - Access sheet not found" };
    }

    const teamNexusData = teamNexusSheet.getDataRange().getValues();
    let employeeId = null;

    for (let i = 0; i < teamNexusData.length; i++) {
      const rowName = teamNexusData[i][0];
      const rowId = teamNexusData[i][1];
      
      if (rowName && rowName.toString().trim() === employeeName) {
        employeeId = rowId ? rowId.toString().trim() : null;
        break;
      }
    }

    if (!employeeId) {
      return { success: false, error: "Employee ID not found" };
    }

    console.log("Employee ID found:", employeeId);

    // Get follow-up history
    const followupSheet = ss.getSheetByName("Follow-up History");
    if (!followupSheet) {
      return { 
        success: true, 
        stats: {
          totalFollowups: 0,
          pendingFollowups: 0,
          completedFollowups: 0,
          overdueFollowups: 0
        }
      };
    }

    const followupData = followupSheet.getDataRange().getValues();
    
    // Calculate target date based on filter
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    
    let targetDate = today;
    
    if (filterType === 'tomorrow') {
      targetDate = new Date(today);
      targetDate.setDate(today.getDate() + 1);
    } else if (filterType === 'custom') {
      // Custom date is handled on client side
      targetDate = today;
    }

    console.log("Target date for filter:", targetDate);

    let totalFollowups = 0;
    let pendingFollowups = 0;
    let completedFollowups = 0;
    let overdueFollowups = 0;

    // Process follow-ups (skip header)
    for (let i = 1; i < followupData.length; i++) {
      const row = followupData[i];
      
      // Column indices:
      // 0: Timestamp, 1: Profile ID, 2: Follow-up Date, 3: Follow-up Note
      // 4: Status, 5: Employee ID, 6: Follow-up ID
      const empId = row[5]; // Column F - Employee ID
      const status = row[4]; // Column E - Status (Scheduled/Completed)
      const followupDateStr = row[2]; // Column C - Follow-up Date
      
      console.log(`Row ${i}: EmpID=${empId}, Status=${status}, Date=${followupDateStr}`);

      // Filter by employee ID - must match exactly
      if (empId !== employeeId) {
        continue;
      }

      // Parse follow-up date
      let followupDateTime;
      try {
        followupDateTime = new Date(followupDateStr);
        if (isNaN(followupDateTime.getTime())) {
          console.log("Invalid date format, skipping row");
          continue;
        }
      } catch (e) {
        console.log("Error parsing date, skipping row:", e);
        continue;
      }

      // Check if follow-up matches the target date (date only, ignoring time)
      const followupDateOnly = new Date(followupDateTime);
      followupDateOnly.setHours(0, 0, 0, 0);
      
      console.log(`Follow-up date only: ${followupDateOnly}, Target date: ${targetDate}`);
      
      if (followupDateOnly.getTime() !== targetDate.getTime()) {
        continue;
      }

      console.log("✅ Follow-up matches date filter");

      // Count the follow-up
      totalFollowups++;
      
      // Check status from Column E
      if (status === "Completed") {
        completedFollowups++;
        console.log("Added to completed count");
      } else {
        pendingFollowups++;
        console.log("Added to pending count");
        
        // Check if overdue (based on DATE only, not time)
        const todayDateOnly = new Date();
        todayDateOnly.setHours(0, 0, 0, 0);
        
        // Calculate days difference
        const timeDiff = todayDateOnly.getTime() - followupDateOnly.getTime();
        const daysDiff = Math.floor(timeDiff / (1000 * 3600 * 24));
        
        console.log(`Today: ${todayDateOnly}, Follow-up: ${followupDateOnly}, Days diff: ${daysDiff}`);
        
        if (daysDiff >= 1) {
          overdueFollowups++;
          console.log("Marked as overdue");
        }
      }
    }

    console.log("Final stats:", { 
      totalFollowups, 
      pendingFollowups, 
      completedFollowups, 
      overdueFollowups 
    });

    return {
      success: true,
      stats: {
        totalFollowups,
        pendingFollowups,
        completedFollowups,
        overdueFollowups
      }
    };

  } catch (error) {
    console.error("Error in getFollowupStats:", error);
    return {
      success: false,
      error: error.toString(),
      stats: {
        totalFollowups: 0,
        pendingFollowups: 0,
        completedFollowups: 0,
        overdueFollowups: 0
      }
    };
  }
}

/**
 * NEW: Get overdue follow-ups with date-only calculation
 */
function getOverdueFollowups(employeeName) {
  try {
    console.log("=== SERVER: getOverdueFollowups() started ===");
    console.log("Employee:", employeeName);

    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Get employee ID
    const teamNexusSheet = ss.getSheetByName("Team Nexus - Access");
    if (!teamNexusSheet) {
      return JSON.stringify({ overdue: [], pending: [], error: "Team Nexus - Access sheet not found" });
    }

    const teamNexusData = teamNexusSheet.getDataRange().getValues();
    let employeeId = null;

    for (let i = 0; i < teamNexusData.length; i++) {
      const rowName = teamNexusData[i][0];
      const rowId = teamNexusData[i][1];
      
      if (rowName && rowName.toString().trim() === employeeName) {
        employeeId = rowId ? rowId.toString().trim() : null;
        break;
      }
    }

    if (!employeeId) {
      return JSON.stringify({ overdue: [], pending: [], error: "Employee ID not found" });
    }

    // Get user's lead sheet for lead details
    const leadsSheetName = `${employeeName} - Leads`;
    const leadsSheet = ss.getSheetByName(leadsSheetName);
    
    // Create lead map for details
    const leadsMap = {};
    if (leadsSheet) {
      const leadsData = leadsSheet.getDataRange().getValues();
      for (let i = 1; i < leadsData.length; i++) {
        const row = leadsData[i];
        const pid = row[0];
        if (!pid) continue;
        
        leadsMap[pid] = {
          profileId: pid,
          fullName: row[1] || '',
          phoneNumber: row[2] || '',
          city: row[3] || '',
          gender: row[4] || '',
          appliedCompanies: row[5] ? row[5].toString().split(",").map(x => x.trim()) : [],
          leadStatus: row[7] || 'Applied',
          generalNote: row[10] || '',
          followupCount: parseInt(row[11]) || 0
        };
      }
    }

    // Get follow-up history
    const followupSheet = ss.getSheetByName("Follow-up History");
    if (!followupSheet) {
      return JSON.stringify({ overdue: [], pending: [] });
    }

    const followupData = followupSheet.getDataRange().getValues();
    const now = new Date();
    
    let overdueFollowups = [];
    let pendingFollowups = [];

    // Process follow-ups (skip header)
    for (let i = 1; i < followupData.length; i++) {
      const row = followupData[i];
      const profileId = row[1];
      const followupDate = row[2];
      const followupNote = row[3];
      const status = row[4];
      const empId = row[5];
      const followupId = row[6];


      console.log("lenght of data================",followupData.length)

      // Filter by employee ID
      if (empId !== employeeId) continue;

      // Parse date
      let followupDateTime;
      try {
        followupDateTime = new Date(followupDate);
      } catch (e) {
        continue;
      }

      // Check if follow-up is not completed
      if (status !== 'Completed') {
        // Calculate overdue based on DATE only (not time)
        const followupDateOnly = new Date(followupDateTime);
        followupDateOnly.setHours(0, 0, 0, 0);
        
        const todayDateOnly = new Date(now);
        todayDateOnly.setHours(0, 0, 0, 0);
        
        // Calculate days difference
        const timeDiff = todayDateOnly.getTime() - followupDateOnly.getTime();
        const daysDiff = Math.floor(timeDiff / (1000 * 3600 * 24));
        
        console.log(`Follow-up ${followupId}: Date=${followupDateOnly}, Today=${todayDateOnly}, Days diff=${daysDiff}`);
        
        const followupItem = {
          profileId: profileId,
          followupDate: followupDate,
          followupNote: followupNote,
          status: status,
          followupId: followupId,
          isOverdue: daysDiff >= 1,
          overdueDays: daysDiff
        };

        // Add lead details if available
        if (leadsMap[profileId]) {
          Object.assign(followupItem, leadsMap[profileId]);
        }

        if (daysDiff >= 1) {
          overdueFollowups.push(followupItem);
          console.log(`Added to overdue: ${profileId} (${daysDiff} days overdue)`);
        } else {
          pendingFollowups.push(followupItem);
        }
      }
    }

    console.log("Overdue follow-ups data:", overdueFollowups);

    console.log("Overdue follow-ups found:", overdueFollowups.length);
    console.log("Pending follow-ups found:", pendingFollowups.length);

    return JSON.stringify({
      overdue: overdueFollowups,
      pending: pendingFollowups
    });

  } catch (error) {
    console.error("Error in getOverdueFollowups:", error);
    return JSON.stringify({ overdue: [], pending: [], error: error.toString() });
  }
}

function getUserFollowups(filterType, employeeName, customDate = null) {
  try {
    console.log("=== SERVER: getUserFollowups() started ===");
    console.log("Filter type:", filterType);
    console.log("Employee name:", employeeName);
    console.log("Custom date:", customDate);

    if (!employeeName) {
      return JSON.stringify({ pending: [], completed: [], error: "Employee name is required" });
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    const followupSheet = ss.getSheetByName("Follow-up History");
    if (!followupSheet) {
      console.error("Follow-up History sheet not found");
      return JSON.stringify({ pending: [], completed: [], error: "Follow-up History sheet not found" });
    }

    // Get employee ID from Team Nexus - Access sheet
    const teamNexusSheet = ss.getSheetByName("Team Nexus - Access");
    if (!teamNexusSheet) {
      return JSON.stringify({ pending: [], completed: [], error: "Team Nexus - Access sheet not found" });
    }

    const teamNexusData = teamNexusSheet.getDataRange().getValues();
    let employeeId = null;

    // Find employee ID by name in column A
    for (let i = 0; i < teamNexusData.length; i++) {
      const rowName = teamNexusData[i][0];
      const rowId = teamNexusData[i][1];
      
      if (rowName && rowName.toString().trim() === employeeName) {
        employeeId = rowId ? rowId.toString().trim() : null;
        break;
      }
    }

    if (!employeeId) {
      console.error("Employee ID not found for:", employeeName);
      return JSON.stringify({ pending: [], completed: [], error: "Employee ID not found for: " + employeeName });
    }

    console.log("Employee:", employeeName, "ID:", employeeId);

    // Get user's leads sheet for lead details
    const leadsSheetName = `${employeeName} - Leads`;
    const leadsSheet = ss.getSheetByName(leadsSheetName);
    const leadsMap = {};
    
    if (leadsSheet) {
      const leadsData = leadsSheet.getDataRange().getValues();
      for (let i = 1; i < leadsData.length; i++) {
        const row = leadsData[i];
        const pid = row[0];
        if (!pid) continue;
        
        leadsMap[pid] = {
          profileId: pid,
          fullName: row[1] || '',
          phoneNumber: row[2] || '',
          city: row[3] || '',
          gender: row[4] || '',
          appliedCompanies: row[5] ? row[5].toString().split(",").map(x => x.trim()) : [],
          leadStatus: row[7] || 'Applied',
          generalNote: row[10] || '',
          followupCount: parseInt(row[11]) || 0,
          interviewLevel: row[8] || 'Not Scheduled'
        };
      }
    }

    // Read follow-up history
    const followupData = followupSheet.getDataRange().getValues();
    
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    
    const tomorrow = new Date(today);
    tomorrow.setDate(today.getDate() + 1);

    let targetDate = null;
    if (filterType === "today") {
      targetDate = today;
    } else if (filterType === "tomorrow") {
      targetDate = tomorrow;
    } else if (filterType === "custom" && customDate) {
      try {
        targetDate = new Date(customDate);
        targetDate.setHours(0, 0, 0, 0);
      } catch (e) {
        console.error("Error parsing custom date:", customDate);
        targetDate = today;
      }
    }

    let pendingFollowups = [];
    let completedFollowups = [];

    // Process follow-ups (skip header row)
    for (let i = 1; i < followupData.length; i++) {
      const row = followupData[i];
      
      // Correct column mapping:
      const timestamp = row[0];        // Column A: Timestamp
      const profileId = row[1];        // Column B: Profile ID
      const followupDateText = row[2]; // Column C: Follow-up Date and Time
      const followupNote = row[3];     // Column D: Follow-up Note
      const status = row[4];           // Column E: Status (Scheduled/Completed) <- THIS IS THE KEY
      const empId = row[5];            // Column F: Employee ID
      const followupId = row[6];       // Column G: Follow-up ID

      console.log(`Row ${i}: Profile=${profileId}, EmpID=${empId}, Status=${status}`);

      // CRITICAL: Verify employee ID matches
      if (empId !== employeeId) {
        continue; // Skip follow-ups that don't belong to this employee
      }

      if (!followupDateText) {
        continue; // Skip blanks
      }

      // Parse follow-up date
      let followupDate;
      try {
        followupDate = new Date(followupDateText);
        if (isNaN(followupDate.getTime())) {
          console.log("Invalid date, skipping:", followupDateText);
          continue;
        }
      } catch (e) {
        console.log("Error parsing date, skipping:", followupDateText);
        continue;
      }

      // Apply date filter if targetDate is set
      if (targetDate) {
        const followupDateOnly = new Date(followupDate);
        followupDateOnly.setHours(0, 0, 0, 0);
        
        if (followupDateOnly.getTime() !== targetDate.getTime()) {
          continue; // Not in selected date range
        }
      }

      const followupItem = {
        profileId: profileId,
        followupDate: followupDateText,
        followupNote: followupNote,
        status: status, // Column E status: "Scheduled" or "Completed"
        followupId: followupId || generateFollowupId()
      };

      // Add lead details if available
      if (leadsMap[profileId]) {
        Object.assign(followupItem, leadsMap[profileId]);
      }

      // Categorize by status in Column E
      if (status === "Completed") {
        completedFollowups.push(followupItem);
        console.log("Added to COMPLETED:", profileId);
      } else {
        pendingFollowups.push(followupItem);
        console.log("Added to PENDING:", profileId);
      }
    }

    console.log("Final counts - Pending:", pendingFollowups.length, "Completed:", completedFollowups.length);

    return JSON.stringify({
      pending: pendingFollowups,
      completed: completedFollowups
    });

  } catch (err) {
    console.error("Error in getUserFollowups:", err);
    return JSON.stringify({ pending: [], completed: [], error: err.toString() });
  }
}

function completeFollowup(profileId, followupId, employeeName) {
  try {
    console.log("=== SERVER: completeFollowup started ===");
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Update Follow-up History sheet
    const followupSheet = ss.getSheetByName("Follow-up History");
    if (!followupSheet) {
      return { success: false, error: "Follow-up History sheet not found" };
    }

    const followupData = followupSheet.getDataRange().getValues();
    let foundFollowup = false;
    let currentLeadStatus = '';

    const timestamp = Utilities.formatDate(
  new Date(),
  Session.getScriptTimeZone(),
  'dd/MM/yyyy HH:mm:ss'
);


    // Find and update the follow-up
    for (let i = 1; i < followupData.length; i++) {
      const row = followupData[i];


      // Column G (index 6) is Follow-up ID
      if (row[6] && row[6].toString().trim() === followupId) {
        // UPDATE COLUMN E (index 4) to "Completed" - THIS IS THE FIX
        followupSheet.getRange(i + 1, 5).setValue("Completed");

        followupSheet.getRange(i + 1, 8).setValue(timestamp);
        
        console.log("Updated follow-up status to Completed in Column E");
        foundFollowup = true;
        break;
      }
    }

    if (!foundFollowup) {
      return { success: false, error: "Follow-up ID not found in Follow-up History" };
    }

    // Update user's lead sheet follow-up count (same as before)
    const leadsSheetName = `${employeeName} - Leads`;
    const leadsSheet = ss.getSheetByName(leadsSheetName);
    
    if (!leadsSheet) {
      return { success: false, error: "Leads sheet not found: " + leadsSheetName };
    }

    const leadsData = leadsSheet.getDataRange().getValues();
    let foundLead = false;

    let updatedFollowupCount = 0;

    for (let i = 1; i < leadsData.length; i++) {

      const row = leadsData[i];
      if (row[0] && row[0].toString().trim() === profileId) {
        const currentCount = parseInt(row[11]) || 0;
        updatedFollowupCount = currentCount + 1;
        leadsSheet.getRange(i + 1, 12).setValue(updatedFollowupCount);
        currentLeadStatus = row[7] || 'Applied';
        console.log(`Updated follow-up count from ${currentCount} to ${currentCount + 1}`);
        foundLead = true;
        break;
      }
    }

    if (!foundLead) {
      return { success: false, error: "Profile ID not found in leads sheet" };
    }

    return { 
      success: true, 
      currentStatus: currentLeadStatus,
      followupCount: updatedFollowupCount,
      message: "Follow-up completed successfully" 
    };

  } catch (error) {
    console.error("Error in completeFollowup:", error);
    return { success: false, error: error.toString() };
  }
}

/**
 * Check if a status requires mandatory follow-up
 */
function checkStatusFollowupRequirement(status) {
  try {
    console.log("=== SERVER: checkStatusFollowupRequirement ===");
    console.log("Checking status:", status);
    
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const leadStatusSheet = spreadsheet.getSheetByName('Lead Status');
    
    if (!leadStatusSheet) {
      // Default statuses that require follow-up
      const defaultFollowupStatuses = ['Call Not Connected', 'Follow-up Required', 'Callback Requested'];
      return {
        success: true,
        requiresFollowup: defaultFollowupStatuses.includes(status)
      };
    }
    
    const dataRange = leadStatusSheet.getRange('A:B');
    const data = dataRange.getValues();
    
    for (let i = 0; i < data.length; i++) {
      const sheetStatus = data[i][0] ? data[i][0].toString().trim() : '';
      const requiresFollowup = data[i][1] ? data[i][1].toString().trim().toLowerCase() === 'yes' : false;
      
      if (sheetStatus === status) {
        console.log(`Found status "${status}": requiresFollowup = ${requiresFollowup}`);
        return {
          success: true,
          requiresFollowup: requiresFollowup
        };
      }
    }
    
    // Status not found in sheet, assume no follow-up required
    return {
      success: true,
      requiresFollowup: false
    };
    
  } catch (error) {
    console.error("Error checking status follow-up requirement:", error);
    return {
      success: false,
      error: error.toString()
    };
  }
}

function getLeadDetailsById(profileId, employeeName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetName = `${employeeName} - Leads`;
    const leadsSheet = ss.getSheetByName(sheetName);
    if (!leadsSheet) {
      return { success: false, error: "Leads sheet not found" };
    }
    
    const data = leadsSheet.getDataRange().getValues();
    const headers = data[0];
    
    // Find profileId column
    const profileIdCol = headers.indexOf("profileId") + 1;
    if (profileIdCol === 0) {
      return { success: false, error: "profileId column not found" };
    }
    
    // Find the row with matching profileId
    for (let i = 1; i < data.length; i++) {
      if (data[i][profileIdCol - 1] === profileId) {
        const lead = {};
        headers.forEach((header, index) => {
          lead[header] = data[i][index];
        });
        
        return { 
          success: true, 
          lead: lead 
        };
      }
    }
    
    return { success: false, error: "Lead not found" };
    
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

// ===============================================================================
function getCompanyModeOfWork(req_company, req_position) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Hiring Requirments");

  if (!sheet) return { success: false };

  const data = sheet.getDataRange().getValues();

  const searchCompany = req_company.toLowerCase().trim();
  const searchPosition = req_position.toLowerCase().trim();

  for (let i = 1; i < data.length; i++) {
    const company = (data[i][0] || "").toString().toLowerCase().trim();
    const role = (data[i][1] || "").toString().toLowerCase().trim();
    const modeOfWork = data[i][12];

    if (company === searchCompany && role === searchPosition) {
      return {
        success: true,
        modeOfWork: modeOfWork || 'Unknown'
      };
    }
  }

  return { success: false };
}

function companyJobLocation(req_company, req_position) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Hiring Requirments");

    if (!sheet) return { success: false };

    const data = sheet.getDataRange().getValues();

    const searchCompany = req_company.toLowerCase().trim();
    const searchPosition = req_position.toLowerCase().trim();

    for (let i = 1; i < data.length; i++) {
      const company = (data[i][0] || "").toString().toLowerCase().trim();
      const role = (data[i][1] || "").toString().toLowerCase().trim();
      const location = data[i][2];
      const flag = data[i][9];
      const priority = data[i][13];

      if (company === searchCompany && role === searchPosition) {
        return {
          success: true,
          location: location || 'Unknown',
          flag:flag || 'Unknown',
          priority:priority || ''
        };
      }
    }

    return { success: false };

  } catch (err) {
    return { success: false, error: err.toString() };
  }
}



function hiringRequirements(req_company, req_position) {
  try {
    const open_sheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheetName = "Hiring Requirments";
    const sheet = open_sheet.getSheetByName(sheetName);

    if (!sheet) {
      console.log("Hiring Requirements sheet not found");
      return {
        success: false,
        error: 'Hiring Requirements sheet not found'
      };
    }

    const data = sheet.getDataRange().getValues();

    if (data.length <= 1) {
      return {
        success: false,
        error: 'No data'
      };
    }

    // Convert search terms to lowercase for case-insensitive comparison
    const searchCompany = req_company.toString().toLowerCase().trim();
    const searchPosition = req_position.toString().toLowerCase().trim();

    for (let i = 1; i < data.length; i++) {
      const row = data[i];

      
        // Convert cell values to lowercase for comparison
        const rowCompany = row[0] ? row[0].toString().toLowerCase().trim() : "";
        const rowPosition = row[1] ? row[1].toString().toLowerCase().trim() : "";

        if (rowCompany === searchCompany && rowPosition === searchPosition) {
          const requirements = {
            gender: row[3],
            ageLimit:row[8],
            experience: row[4],
            yearsOfExperience: row[5],
            language: row[6],
            highestQualification: row[7]
          };
          return {
            success: true,
            requirements: requirements
          };
        }
        // Don't return here - continue searching other rows
      
    }

    
  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}
// =======================================================================================
function companyHighlights(req_company, req_position) {
  try {
    const open_sheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheetName = "Hiring Requirments";
    const sheet = open_sheet.getSheetByName(sheetName);

    if (!sheet) {
      console.log("Hiring Requirments sheet not found");
      return {
        success: false,
        error: 'Hiring Requirments sheet not found'
      };
    }

    const data = sheet.getDataRange().getValues();

    if (data.length <= 1) {
      return {
        success: false,
        error: 'No data'
      };
    }

    // Convert search terms to lowercase for case-insensitive comparison
    const searchCompany = req_company.toString().toLowerCase().trim();
    const searchPosition = req_position.toString().toLowerCase().trim();

    for (let i = 1; i < data.length; i++) {
      const row = data[i];

      
        // Convert cell values to lowercase for comparison
        const rowCompany = row[0] ? row[0].toString().toLowerCase().trim() : "";
        const rowPosition = row[1] ? row[1].toString().toLowerCase().trim() : "";


      

        if (rowCompany === searchCompany && rowPosition === searchPosition) {
          const listOfhighlights=row[11]
          
          console.log("highlightss===",listOfhighlights)
          return {
            success: true,
            listOfhighlights: listOfhighlights
          };
        }
      }
    

    return {
      success: false,
      error: 'No highlights found for the requested company and position'
    };
  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

// =======================================================================================
function jobDescription(req_company, req_position) {
  try {
    const open_sheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheetName = "Hiring Requirments";
    const sheet = open_sheet.getSheetByName(sheetName);

    if (!sheet) {
      console.log("Hiring Requirments sheet not found");
      return {
        success: false,
        error: 'Hiring Requirments sheet not found'
      };
    }

    const data = sheet.getDataRange().getValues();

    if (data.length <= 1) {
      return {
        success: false,
        error: 'No data'
      };
    }

    // Convert search terms to lowercase for case-insensitive comparison
    const searchCompany = req_company.toString().toLowerCase().trim();
    const searchPosition = req_position.toString().toLowerCase().trim();

    for (let i = 1; i < data.length; i++) {
      const row = data[i];

      
        // Convert cell values to lowercase for comparison
        const rowCompany = row[0] ? row[0].toString().toLowerCase().trim() : "";
        const rowPosition = row[1] ? row[1].toString().toLowerCase().trim() : "";


      

        if (rowCompany === searchCompany && rowPosition === searchPosition) {
          const jobDescriptions=row[10]
          
          console.log("jobDescriptions===",jobDescriptions)
          return {
            success: true,
            jobDescriptions: jobDescriptions
          };
        }
      }
    

    return {
      success: false,
      error: 'No job Descriptions found for the requested company and position'
    };
  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

//=============================================================================
function getEligibleCompanies(
  personGender,
  personExpType,
  personExpYears,
  personLanguages,
  personQualification
) {
  try {
    // ===== OPEN SHEET =====
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Hiring Requirments");

    const data = sheet.getDataRange().getValues();
    const companies = [];

    // ===== QUALIFICATION ORDER =====
    const qualificationRank = {
      "Below 10th":1,
      "10th/12th Pass": 1,
      "Diploma": 2,
      "Graduate (Bachelor's Degree)": 3,
      "Postgraduate (Master's Degree)": 4,
    };

    // ===== NORMALIZE PERSON INPUT =====
    personGender = personGender.toString().trim().toLowerCase();
    personExpType = personExpType.toString().trim().toLowerCase();
    personExpYears = personExpYears.toString().trim().toLowerCase() || 0;
    personQualification = personQualification.toString().trim();

    const personQualRank = qualificationRank[personQualification];
    if (!personQualRank) throw new Error("Invalid qualification");

    const personLangs = personLanguages
      .toString()
      .toLowerCase()
      .split(',')
      .map(l => l.trim());

    // ===== HELPER FUNCTIONS =====
    function normalizeQualification(q) {
      q = q.toLowerCase();
      if (q.includes("10")) return "10th/12th Pass";
      if (q.includes("diploma")) return "Diploma";
      if (q.includes("graduate")) return "Graduate (Bachelor's Degree)";
      if (q.includes("postgraduate")) return "Postgraduate (Master's Degree)";
      if (q.includes("any")) return "10th/12th Pass"; // lowest
      return null;
    }

   function parseExperience(exp) {
  if (!exp) return 0;
  exp = exp.toLowerCase();

  if (exp.includes("6 months")) return 0.5;
  if (exp.includes("1 - 2")) return 1;
  if (exp.includes("2 - 3")) return 2;
  if (exp.includes("3 - 4")) return 3;
  if (exp.includes("4 - 5")) return 4;
  if (exp.includes("5 - 6")) return 5;
  if (exp.includes("6 - 7")) return 6;
  if (exp.includes("7 - 8")) return 7;
  if (exp.includes("8 - 9")) return 8;
  if (exp.includes("10+")) return 10;

  return 0;
}


    // ===== MAIN LOOP =====
    for (let i = 1; i < data.length; i++) {
      const row = data[i];

      // 1️⃣ ACTIVE / DEACTIVATE CHECK (Column L)
      const status = row[9]?.toString().trim().toLowerCase();
      if (status !== "active") continue;

      // Company Name (Column C)
      const company = row[0]?.toString().trim();
      if (!company) continue;

      const jobPosition = row[1]?.toString().trim();
      if (!jobPosition) continue;

      // Gender Requirement (Column E)
      const reqGender = row[3]?.toString().trim().toLowerCase();

      // Experience Type (Column F)
      const reqExpType = row[4]?.toString().trim().toLowerCase();

      // Experience Years (Column G)
      const reqExpYears = parseExperience(row[5]?.toString() || "");
      const personYears = parseExperience(personExpYears);


      // Languages (Column H)
      const reqLangs = row[6]
        ?.toString()
        .toLowerCase()
        .split(',')
        .map(l => l.trim()) || [];

      // Qualification (Column I)
      const reqQualification = normalizeQualification(row[7]?.toString().trim().toLowerCase() || "");
      const reqQualRank = qualificationRank[reqQualification];
      if (!reqQualRank) continue;

      // ===== FILTERING LOGIC =====

      // Gender
      if (reqGender !== "any" && reqGender !== personGender) continue;

      // Qualification
      if (personQualRank < reqQualRank) continue;


// Experience matching
if (reqExpType === "freshers") {
  if (personExpType !== "no") continue;
}

if (reqExpType === "experienced") {
  if (personExpType !== "yes") continue;
  if (personYears < reqExpYears) continue;
}
// both → allowed


      // Language (person must know all required)
      if (!reqLangs.every(l => personLangs.includes(l))) continue;

      // ===== PASSED ALL CHECKS =====
      companies.push({
  companyName: company,
  jobPosition: jobPosition
});

    }

    console.log(companies)
    // ===== RETURN RESULT =====
    return {
      success: true,
      companies: companies,
      count: companies.length
    };

  } catch (error) {
    return {
      success: false,
      message: error.message,
      companies: [],
      count: 0
    };
  }
}
// =================================================================================
function saveEligibleCompaniesResult(profileId, employeeId, employeeName, companies) {
  console.log("=== saveEligibleCompaniesResult ===");

  try {
    let employeeUpdated = false;
    let compassUpdated = false;

    /* ================= EMPLOYEE LEADS SHEET ================= */

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const expectedSheetName = `${employeeName} - Leads`;

    let sheet = ss.getSheetByName(expectedSheetName);

    if (!sheet) {
      const sheets = ss.getSheets();
      for (let s of sheets) {
        const name = s.getName();
        if (name.includes(employeeName) && name.includes('Leads')) {
          sheet = s;
          break;
        }
      }
    }

    if (!sheet) {
      return {
        success: false,
        error: `Sheet not found for employee: ${employeeName} ${employeeId}`
      };
    }

    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(profileId).trim()) { // Column A
        sheet.getRange(i + 1, 6).setValue(companies); // Column F
        employeeUpdated = true;
        break;
      }
    }

    /* ================= COMPASS RESUMES SHEET ================= */

    const compassSheetId = '14L0Ku2MPYUJ6l8Nk5zNr4jpUJk_tAFHGz26Lns46k9k';
    const compassSheet = SpreadsheetApp.openById(compassSheetId);
    const compassSubSheet = compassSheet.getSheetByName('Resumes');

    const sheetData = compassSubSheet.getDataRange().getValues();

    for (let i = 1; i < sheetData.length; i++) {
      if (String(sheetData[i][10]).trim() === String(profileId).trim()) {// Column K
        const row = i + 1;

        const colY = 25; // Eligibility
        const colZ = 26; // Applied Companies
        const colAA = 27; // Applied Count

        // ✅ Correct columns
        compassSubSheet.getRange(row, colZ).setValue(companies);
        compassSubSheet.getRange(row, colY).setValue("Eligible");

        const currentCount = Number(sheetData[i][26]) || 0; // AA
        compassSubSheet.getRange(row, colAA).setValue(currentCount + 1);

        compassUpdated = true;
        break;
      }
    }
    console.log(`Employee sheet updated: ${employeeUpdated}, Compass sheet updated: ${compassUpdated}`)
    return {
      success: employeeUpdated && compassUpdated,
      message: `Employee sheet updated: ${employeeUpdated}, Compass sheet updated: ${compassUpdated}`
    };
    

  } catch (error) {
    return {
      success: false,
      message: error.message
    };
  }
}

//=============================================================================
function getInterviewSlots(level) {
  console.log("SERVER HIT, level =", level);

  const interviewSheetId = '1oBQERjpIcMutM9XFD6xo-3B2FCawqW1RgNOp-xbCDPM';

  try {
    const ss = SpreadsheetApp.openById(interviewSheetId);

    // Safe sheet selection
    const sheetName = level.toString().toLowerCase().includes("second")
      ? "Sales Level"
      : "HR Level";

    const sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      console.log("❌ Sheet not found:", sheetName);
      return { success: false, dates: [] };
    }

    const data = sheet.getDataRange().getValues();

    /*
      A → Date
      B → Time
      C → Max Slots
      D → Booked Slots
    */

    const dateMap = {};

    for (let i = 1; i < data.length; i++) {
      const [date, time, maxSlots, bookedSlots] = data[i];

      if (!(date instanceof Date) || !(time instanceof Date)) continue;

      const dateKey = Utilities.formatDate(
        date,
        Session.getScriptTimeZone(),
        "dd/MM/yyyy"
      );

      if (!dateMap[dateKey]) {
        dateMap[dateKey] = {
          date: dateKey,
          times: [],
          totalSlots: 0,
          fullSlots: 0
        };
      }

      const isAvailable = Number(bookedSlots) < Number(maxSlots);

      dateMap[dateKey].times.push({
        time: Utilities.formatDate(
          time,
          Session.getScriptTimeZone(),
          "hh:mm a"
        ),
        isAvailable: isAvailable
      });

      dateMap[dateKey].totalSlots++;
      if (!isAvailable) dateMap[dateKey].fullSlots++;
    }

    const dates = Object.values(dateMap).map(d => ({
      date: d.date,
      isAvailable: d.fullSlots < d.totalSlots,
      times: d.times
    }));

    return {
      success: true,
      dates: dates
    };

  } catch (err) {
    console.log("ERROR:", err);
    return { success: false, dates: [] };
  }
}

// =========================================================================================
function getLevelByProfileId(profileId, employeeName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const expectedSheetName = `${employeeName} - Leads`;

    let sheet = ss.getSheetByName(expectedSheetName);

    // If exact sheet name not found, try partial match
    if (!sheet) {
      const sheets = ss.getSheets();
      for (let i = 0; i < sheets.length; i++) {
        const name = sheets[i].getName();
        if (name.includes(employeeName) && name.includes('Leads')) {
          sheet = sheets[i];
          break;
        }
      }
    }

    if (!sheet) {
      return {
        success: false,
        error: `Sheet not found for employee: ${employeeName}`
      };
    }

    const data = sheet.getDataRange().getValues();

    const secondLevelRanges = [
      '3 - 4 Years',
      '4 - 5 Years',
      '5 - 6 Years',
      '6 - 7 Years',
      '7 - 8 Years',
      '8 - 9 Years',
      '10+ Years'
    ];

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == profileId) { // Column A
        const experience = data[i][14]; // Column P

        const interviewLevel = secondLevelRanges.includes(experience)
          ? 'Second Level'
          : 'First Level';

        // ✅ Update Column L (row index +1, column 12)
        sheet.getRange(i + 1, 9).setValue(interviewLevel);

        return {
          success: true,
          level: interviewLevel
        };
      }
    }

    return {
      success: false,
      error: `Profile ID ${profileId} not found`
    };

  } catch (err) {
    console.log("ERROR:", err);
    return {
      success: false,
      error: err.toString()
    };
  }
}

function saveReservationDetailsTosheet(details) {
  try {
    console.log("=== Starting saveReservationDetailsTosheet ===");
    console.log("Input details:", details);
    
    const reservationSheetId = '1oBQERjpIcMutM9XFD6xo-3B2FCawqW1RgNOp-xbCDPM';
    const ss = SpreadsheetApp.openById(reservationSheetId);

    // ===== Reservation Sheet =====
    const sheet = ss.getSheetByName("Reservation");
    if (!sheet) {
      throw new Error("Reservation sheet not found");
    }

    const row = [
      details.employeeId,
      details.profileId,
      details.leadName,
      details.phoneNumber,
      details.level,
      details.reservationDate,
      details.reservationTime,
      details.currentDateTime,
      "Reserved"
    ];

    sheet.appendRow(row);
    console.log("✓ Data appended to Reservation sheet");

    // ===== Format Conversion =====
    // Convert Reservation date from "DD-MM-YYYY" to "DD/MM/YYYY" format
    const formattedReservationDate = details.reservationDate.replace(/-/g, '/');
    console.log("Reservation date (original):", details.reservationDate);
    console.log("Reservation date (formatted):", formattedReservationDate);
    
    // Convert Reservation time to 12-hour format with AM/PM (lowercase)
    let formattedReservationTime = "";
    try {
      // Remove any AM/PM from the time first
      const timeWithoutAmPm = details.reservationTime.replace(/\s?(AM|PM|am|pm)/i, '').trim();
      const [hours, minutes] = timeWithoutAmPm.split(':');
      const hourNum = parseInt(hours);
      const minuteNum = minutes ? parseInt(minutes) : 0;
      
      // Check if original had AM or PM
      const hasAmPm = /\s?(AM|PM|am|pm)/i.test(details.reservationTime);
      const isPM = hasAmPm ? /PM|pm/i.test(details.reservationTime) : hourNum >= 12;
      
      let displayHour;
      if (hasAmPm) {
        displayHour = hourNum === 12 ? 12 : hourNum;
      } else {
        displayHour = hourNum === 0 ? 12 : (hourNum > 12 ? hourNum - 12 : hourNum);
      }
      
      formattedReservationTime = displayHour + ':' + (minuteNum < 10 ? '0' + minuteNum : minuteNum) + ' ' + (isPM ? 'pm' : 'am');
    } catch (e) {
      console.warn("Time conversion error:", e);
      formattedReservationTime = details.reservationTime;
    }
    
    console.log("Reservation time (original):", details.reservationTime);
    console.log("Reservation time (formatted):", formattedReservationTime);

    // ===== Slot Sheet Selection =====
    const sheetName = details.level
      .toString()
      .toLowerCase()
      .includes("second")
        ? "Sales Level"
        : "HR Level";

    console.log("Looking for slot sheet:", sheetName);
    const slotSheet = ss.getSheetByName(sheetName);
    if (!slotSheet) {
      throw new Error(`Slot sheet '${sheetName}' not found`);
    }

    const data = slotSheet.getDataRange().getValues();
    console.log(`Found ${data.length} rows in ${sheetName} sheet`);
    
    let matchFound = false;
    let matchRow = -1;

    for (let i = 1; i < data.length; i++) {
      const sheetDate = data[i][0]; // Column A
      const sheetTime = data[i][1]; // Column B
      const sheetCount = data[i][3]; // Column D
      
      // Format sheet date to "DD/MM/YYYY"
      let sheetDateStr;
      if (sheetDate instanceof Date) {
        sheetDateStr = Utilities.formatDate(sheetDate, Session.getScriptTimeZone(), "dd/MM/yyyy");
      } else if (sheetDate) {
        // If it's already a string, clean it up
        sheetDateStr = sheetDate.toString().split(' ')[0]; // Take only the date part
      }
      
      // Format sheet time to "h:mm a" (e.g., "11:45 am")
      let sheetTimeStr = "";
      if (sheetTime instanceof Date) {
        // Extract just the time portion
        sheetTimeStr = Utilities.formatDate(sheetTime, Session.getScriptTimeZone(), "h:mm a").toLowerCase();
      } else if (sheetTime) {
        // If it's already a string, try to parse it
        sheetTimeStr = sheetTime.toString().toLowerCase();
        // If it looks like a full date string, extract time
        if (sheetTimeStr.includes(':')) {
          const timeMatch = sheetTimeStr.match(/(\d{1,2}:\d{2})(?::\d{2})?\s*(am|pm)?/i);
          if (timeMatch) {
            let hour = timeMatch[1];
            let ampm = timeMatch[2] ? timeMatch[2].toLowerCase() : '';
            
            // If no AM/PM in the string but we have a Date object, infer it
            if (!ampm && sheetTime instanceof Date) {
              const hours = sheetTime.getHours();
              ampm = hours >= 12 ? 'pm' : 'am';
            }
            sheetTimeStr = hour + ' ' + ampm;
          }
        }
      }
      
      // Clean up the strings for comparison
      sheetDateStr = sheetDateStr ? sheetDateStr.trim() : '';
      sheetTimeStr = sheetTimeStr ? sheetTimeStr.trim() : '';
      
      // Debug for first few rows and matching date
      if (i <= 5 || sheetDateStr === formattedReservationDate) {
        console.log(`Row ${i+1} - Sheet Date: ${sheetDateStr}, Sheet Time: ${sheetTimeStr}, Reservation Date: ${formattedReservationDate}, Reservation Time: ${formattedReservationTime}`);
      }
      
      if (sheetDateStr === formattedReservationDate && sheetTimeStr === formattedReservationTime.toLowerCase()) {
        console.log(`✓ MATCH FOUND at row ${i+1}!`);
        console.log(`Sheet count: ${sheetCount}`);
        matchFound = true;
        matchRow = i;
        
        const currentCount = Number(sheetCount) || 0;
        const newCount = currentCount + 1;
        
        console.log(`Updating count from ${currentCount} to ${newCount} at cell R${i+1}C4`);
        
        slotSheet.getRange(i + 1, 4).setValue(newCount);
        console.log(`✓ Count updated successfully`);
        break;
      }
    }

    if (!matchFound) {
      console.error("❌ No matching slot found!");
      console.error("Looking for:", formattedReservationDate, formattedReservationTime);
      
      // Check what's actually in the sheet for that date
      console.log("Checking all slots for date:", formattedReservationDate);
      for (let i = 1; i < data.length; i++) {
        const sheetDate = data[i][0];
        let sheetDateStr;
        if (sheetDate instanceof Date) {
          sheetDateStr = Utilities.formatDate(sheetDate, Session.getScriptTimeZone(), "dd/MM/yyyy");
        } else if (sheetDate) {
          sheetDateStr = sheetDate.toString().split(' ')[0];
        }
        
        if (sheetDateStr === formattedReservationDate) {
          const sheetTime = data[i][1];
          let sheetTimeStr = "";
          if (sheetTime instanceof Date) {
            sheetTimeStr = Utilities.formatDate(sheetTime, Session.getScriptTimeZone(), "h:mm a").toLowerCase();
          }
          console.log(`Row ${i+1}: Date ${sheetDateStr}, Time ${sheetTimeStr}`);
        }
      }
    }

    return {
      success: true,
      message: matchFound 
        ? "Reservation saved and slot count updated successfully" 
        : "Reservation saved but no matching slot found to update count"
    };

  } catch (err) {
    console.error("❌ ERROR:", err);
    return {
      success: false,
      error: err.message
    };
  }
}

// ===========================================================================
function checkReservation(profileId) {
  try {
    // Google Sheet ID
    const reservationSheetId = '1oBQERjpIcMutM9XFD6xo-3B2FCawqW1RgNOp-xbCDPM';

    // Open spreadsheet
    const ss = SpreadsheetApp.openById(reservationSheetId);
    const sheet = ss.getSheetByName("Reservation");

    if (!sheet) {
      throw new Error("Reservation sheet not found");
    }

    console.log("✅ Sheet found");

    // Get all data
    const data = sheet.getDataRange().getValues();

    // Loop through rows (skip header)
    for (let i = 1; i < data.length; i++) {
      const row = data[i];

      // Match profileId (Column B)
      if (String(row[1]) === String(profileId)) {

        console.log("✅ Match found:", row);

        const rawDate = row[5]; // Column F - Date
        const rawTime = row[6]; // Column G - Time

        // Format date → DD/MM/YYYY
        const formattedDate = Utilities.formatDate(
          new Date(rawDate),
          "Asia/Kolkata",
          "dd/MM/yyyy"
        );

        // Format time → HH:mm (24-hour)
        const formattedTime = Utilities.formatDate(
          new Date(rawTime),
          "Asia/Kolkata",
          "HH:mm"
        );

        console.log(
          "Formatted Date:",
          formattedDate,
          "Formatted Time:",
          formattedTime,
          "ProfileId:",
          profileId
        );

        // Return clean backend response
        return {
          success: true,
          profileId: profileId,
          date: formattedDate,
          time: formattedTime
        };
      }
      
    }

    // No reservation found
    console.log("⚠️ No reservation found for:", profileId);

    return {
      success: true,
      message: "No reservation found for this profileId"
    };

  } catch (err) {
    console.error("❌ ERROR:", err);

    return {
      success: false,
      error: err.message
    };
  }
}

//=============================main function==============================================
function processInterviewSubmission(data) {

  try {
    // 🔹 CONFIG
    const folderId = '1aH1J4A7BQuECv3llHW2ATYs9lSi55deeSRACuIqJ8fgQwcq_AsYSq77co6mKqunAxOc7tvgS';

    // 🔹 Date + time
    const today = Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      'dd/MM/yyyy'
    );

    const timestamp = Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      'dd/MM/yyyy HH:mm:ss'
    );

    // 🔹 File name
    const safeLeadName = data.leadName.replace(/[^\w\s]/gi, '');
    const fileName = `${safeLeadName}_${today}`;

    // 🔹 Save resume to Drive
    const folder = DriveApp.getFolderById(folderId);

    const blob = Utilities.newBlob(
      Utilities.base64Decode(data.resumeBase64),
      data.resumeMimeType,
      fileName
    );

    const file = folder.createFile(blob);
    const fileUrl = file.getUrl();

    // 🔹 Save interview row
    saveInterviewData(data, fileUrl, timestamp);

    // 🔹 Call other flows
    saveAndUpdateHRdatabase(data, fileUrl);
    updateWabisSheet(data);
    updateLeadsSheet(data, fileUrl);
    updateInterviewBookingSlot(data);

    return {
      success: true,
      fileUrl
    };

  } catch (err) {
    console.error('❌ processInterviewSubmission error:', err);
    return { success: false, message: err.message };
  }
}


// =============================Horizon sheet===============================================
function saveInterviewData(data, fileUrl, timestamp) {
  const sheetId = '1EYYy6B1dtokoWMOhTAO_liHECcAZUrGZBoH2WPPQNvw';
  const sheetName = 'Interviews';

  const sheet = SpreadsheetApp
    .openById(sheetId)
    .getSheetByName(sheetName);

  sheet.appendRow([
    data.employeeId,
    data.profileId,
    fileUrl,
    data.leadName,
    data.leadPhone,
    data.interviewLevel,
    data.selectedCompanies.join(', '),
    data.interviewDate,
    data.interviewTime,
    data.interviewNotes,
    timestamp
  ]);
}



// =========================HR database==================================================================
function saveAndUpdateHRdatabase(data,fileUrl) {
  try {
    const hrSheetId = '1-wXQZzxHjgJEaHZtuBFe-WqALi1vfL5Z1WCchUCnqos'; 
    const sheetName = 'Final Database';       
    const ss = SpreadsheetApp.openById(hrSheetId);
    const sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      throw new Error('HR sheet not found!');
    }

    const values = sheet.getDataRange().getValues();

    // Loop through rows starting from row 2 (skip header)
    for (let i = 1; i < values.length; i++) {
      const rowProfileId = values[i][0]; // Column A

      if (String(rowProfileId).trim() === String(data.profileId).trim()) {
        // Update column B (index 1)
        sheet.getRange(i + 1, 2).setValue(fileUrl);
        console.log(`Updated resume URL for profileId: ${data.profileId}`);
        return; // Exit after updating
      }
    }

    console.log(`Profile ID ${data.profileId} not found in HR sheet.`);

  } catch (err) {
    console.error('Error in saveAndUpdateHRdatabase:', err);
  }
}

function updateWabisSheet(data) {
  try {
    const wabisSheetId = '1Amfgfix1o60L6tqWwO7Xc9IFijAxmLFWUwyXO67yOAc';  
    const sheetName = 'Interview Schedule 1';             

    const ss = SpreadsheetApp.openById(wabisSheetId);
    const sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      throw new Error('Wabis sheet not found!');
    }

    // Current date (no time)
    const today = Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      'dd/MM/yyyy'
    );

    data.leadPhone = '91' + data.leadPhone;


    // Prepare row
    const row = [
      data.profileId,                         // Profile ID
      data.leadName,                          // Name
      data.leadPhone,                         // Phone
      data.selectedCompanies.join(', '),      // Companies
      data.interviewDate,                     // Interview Date
      data.interviewTime,                     // Interview Time
      today,                                  // Current Date
      `Team Nexus - ${data.employeeId}`          // Data Source
    ];

    sheet.appendRow(row);

    console.log('Wabis sheet updated successfully');

  } catch (err) {
    console.error('Error in updateWabisSheet:', err);
  }
}

function updateLeadsSheet(data,fileUrl) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const employeeMapSheetName = 'Team Nexus - Access'; 

    /* --------------------------------------------------
       STEP 1: Get employeeName using employeeId
    -------------------------------------------------- */
    const empSheet = ss.getSheetByName(employeeMapSheetName);
    if (!empSheet) throw new Error('Employees sheet not found');

    const empData = empSheet.getDataRange().getValues();
    let employeeName = null;

    for (let i = 1; i < empData.length; i++) {
      if (String(empData[i][1]).trim() === String(data.employeeId).trim()) {
        employeeName = empData[i][0]; // Column A = Employee Name
        break;
      }
    }

    if (!employeeName) {
      throw new Error(`Employee name not found for ID: ${data.employeeId}`);
    }

    /* --------------------------------------------------
       STEP 2: Find Leads sheet for this employee
    -------------------------------------------------- */
    const expectedSheetName = `${employeeName} - Leads`;
    let sheet = ss.getSheetByName(expectedSheetName);

    // Fallback: partial match
    if (!sheet) {
      const sheets = ss.getSheets();
      for (let s of sheets) {
        const name = s.getName();
        if (name.includes(employeeName) && name.includes('Leads')) {
          sheet = s;
          break;
        }
      }
    }

    if (!sheet) {
      throw new Error(`Leads sheet not found for employee: ${employeeName}`);
    }

    /* --------------------------------------------------
       STEP 3: Update resume link (ProfileId match)
    -------------------------------------------------- */
    const values = sheet.getDataRange().getValues();

    for (let i = 1; i < values.length; i++) {
  const rowProfileId = values[i][0];

  console.log('Row profileId:', rowProfileId, ' | Incoming:', data.profileId);

  if (
    rowProfileId !== '' &&
    String(rowProfileId).trim().toUpperCase() ===
    String(data.profileId).trim().toUpperCase()
  ) {
    sheet.getRange(i + 1, 21).setValue(fileUrl);
    console.log('✅ Resume URL updated at row:', i + 1,' and  col V : ',21);
    console.log('data.fileurl',fileUrl);
    return;
  }
}


    console.log(`Profile ID ${data.profileId} not found in Leads sheet`);

  } catch (err) {
    console.error('❌ updateLeadsSheet error:', err.message);
  }
}


function updateInterviewBookingSlot(data) {
  try {
    const bookingSheetId = '1oBQERjpIcMutM9XFD6xo-3B2FCawqW1RgNOp-xbCDPM'; 
    const sheetName = 'Reservation'; 

    const ss = SpreadsheetApp.openById(bookingSheetId);
    const sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      throw new Error('Interview booking sheet not found');
    }

    const values = sheet.getDataRange().getValues();

    // Start from row 2 (skip header)
    for (let i = 1; i < values.length; i++) {
      const rowProfileId = values[i][1]; // Column B

      if (
        rowProfileId &&
        String(rowProfileId).trim().toUpperCase() ===
        String(data.profileId).trim().toUpperCase()
      ) {
        sheet.deleteRow(i + 1); // delete entire row
        console.log(`🗑️ Booking slot deleted for profileId: ${data.profileId}`);
        return;
      }
    }

    console.log(` No booking slot found for profileId: ${data.profileId}`);

  } catch (err) {
    console.error('❌ updateInterviewBookingSlot error:', err.message);
  }
}

// ========================================================================
// function checkExistingFollowUp(employeeName, employeeId, profileId) {
//   try {
//     const ss = SpreadsheetApp.openById('1CoecJqKpv55_nQ5p7qOhiWZsGOszzcxJrg27DLw-6iU');
//     const sheet = ss.getSheetByName('Follow-up History');

//     if (!sheet) {
//       return { found: false, error: 'Sheet not found' };
//     }

//     const data = sheet.getDataRange().getValues();

//     for (let i = 1; i < data.length; i++) {
//       const row = data[i];

//       if (
//         row[1] === profileId &&
//         row[5] === employeeId &&
//         String(row[4]).toLowerCase() === 'scheduled'
//       ) {
//         return {
//           found: true,
//           rowNumber: i + 1,
//           rowData: row
//         };
//       }
//     }

//     return { found: false };

//   } catch (err) {
//     return {
//       found: false,
//       error: err.message
//     };
//   }
// }
