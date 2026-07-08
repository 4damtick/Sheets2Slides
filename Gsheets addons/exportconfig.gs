/**
 * Creates custom menu when spreadsheet opens
 */
function onOpen(e) {
  SpreadsheetApp.getUi()
    .createMenu('Sheets2Slides')
    .addItem('⚙️Configure & Export', 'showSettingsDialog')
    .addToUi();
}

function onInstall(e) {
    SpreadsheetApp.getUi()
    .createMenu('Sheets2Slides')
    .addItem('⚙️Configure & Export', 'showSettingsDialog')
    .addToUi();
}

/**
 * Shows the settings dialog
 */
function showSettingsDialog() {
  try {
    const html = HtmlService.createHtmlOutputFromFile('SettingsDialog')
    .setTitle('Sheets2Slides Sync Config Export');
    SpreadsheetApp.getUi().showSidebar(html);
  } catch (e) {
    SpreadsheetApp.getUi().alert('Error: ' + e.toString() + '\n\nMake sure you created the SettingsDialog.html file!');
  }
}

/**
 * Gets the current active sheet name
 */
function getCurrentSheetName() {
  return SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
}

/**
 * Auto-save config for generated file
 * Automatically creates a config entry each time a PDF/PNG is exported
 */
function autoSaveConfig(fileName, settings, sheetTab, cellRange) {
  try {
    // Create a specific config for this exact screenshot
    var configSettings = JSON.parse(JSON.stringify(settings)); // Deep copy
    
    // Override to be specific to this single screenshot
    configSettings.sheetTabs = sheetTab;
    configSettings.cellRanges = cellRange;
    
    // Use the filename (without extension) as config name
    var configName = fileName.replace(/\.pdf$/i, '').replace(/\.png$/i, '');
    
    // Save the config
    var result = saveConfig(configName, configSettings);
    
    if (result.success) {
      Logger.log('✅ Auto-saved config: "' + configName + '" (ID: ' + result.configId + ')');
      return result.configId;
    } else {
      Logger.log('⚠️ Failed to auto-save config for "' + configName + '": ' + result.message);
      return null;
    }
  } catch (e) {
    Logger.log('⚠️ Error auto-saving config: ' + e.toString());
    return null;
  }
}

var CONFIGS_SHEET_NAME = 'Configs';
var SHARED_SETTINGS_SHEET_NAME = 'Settings';
var SHARED_SETTINGS_HEADERS = ['Key', 'Value', 'Updated'];
var SHARED_SETTINGS_DEFAULTS = {
  gcsInputBucket: '',
  gcsOutputBucket: '',
  cloudConvertApiKey: ''
};

function ensureSharedSettingsSheet_(ss) {
  var sheet = ss.getSheetByName(SHARED_SETTINGS_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHARED_SETTINGS_SHEET_NAME);
  }

  sheet.getRange(1, 1, 1, SHARED_SETTINGS_HEADERS.length).setValues([SHARED_SETTINGS_HEADERS]);
  sheet.getRange(1, 1, 1, SHARED_SETTINGS_HEADERS.length)
    .setBackground('#4285f4')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setWrap(true);
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 180);
  sheet.setColumnWidth(2, 320);
  sheet.setColumnWidth(3, 180);

  var rowMap = {};
  if (sheet.getLastRow() >= 2) {
    var existingRows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
    for (var i = 0; i < existingRows.length; i++) {
      if (existingRows[i][0]) {
        rowMap[String(existingRows[i][0])] = i + 2;
      }
    }
  }

  Object.keys(SHARED_SETTINGS_DEFAULTS).forEach(function(key) {
    if (!rowMap[key]) {
      sheet.appendRow([key, SHARED_SETTINGS_DEFAULTS[key], '']);
    }
  });

  return sheet;
}

function readSharedConversionSettings_(sheet) {
  var settings = JSON.parse(JSON.stringify(SHARED_SETTINGS_DEFAULTS));
  if (sheet.getLastRow() < 2) {
    return settings;
  }

  var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  for (var i = 0; i < rows.length; i++) {
    var key = rows[i][0];
    if (key && Object.prototype.hasOwnProperty.call(settings, key)) {
      settings[key] = rows[i][1] || '';
    }
  }

  settings.gcsInputBucket = String(settings.gcsInputBucket || '').trim();
  settings.gcsOutputBucket = String(settings.gcsOutputBucket || '').trim();
  settings.cloudConvertApiKey = String(settings.cloudConvertApiKey || '').trim();
  return settings;
}

function writeSharedConversionSettings_(sheet, settings) {
  var normalizedSettings = {
    gcsInputBucket: String(settings.gcsInputBucket || '').trim(),
    gcsOutputBucket: String(settings.gcsOutputBucket || '').trim(),
    cloudConvertApiKey: String(settings.cloudConvertApiKey || '').trim()
  };
  var updatedAt = new Date().toISOString();
  var rowMap = {};

  if (sheet.getLastRow() >= 2) {
    var existingKeys = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
    for (var i = 0; i < existingKeys.length; i++) {
      if (existingKeys[i][0]) {
        rowMap[String(existingKeys[i][0])] = i + 2;
      }
    }
  }

  Object.keys(SHARED_SETTINGS_DEFAULTS).forEach(function(key) {
    var rowIndex = rowMap[key];
    if (!rowIndex) {
      sheet.appendRow([key, normalizedSettings[key], updatedAt]);
      return;
    }

    sheet.getRange(rowIndex, 2).setValue(normalizedSettings[key]);
    sheet.getRange(rowIndex, 3).setValue(updatedAt);
  });

  return normalizedSettings;
}

function getSharedConversionSettings() {
  try {
    var configSheetId = getConfigSheetId();
    if (!configSheetId) {
      return {
        success: false,
        message: 'No config sheet linked. Please create or link one first.'
      };
    }

    var ss = SpreadsheetApp.openById(configSheetId);
    if (!ss.getSheetByName(CONFIGS_SHEET_NAME)) {
      return {
        success: false,
        message: 'Configs tab not found in the linked sheet.'
      };
    }

    return {
      success: true,
      settings: readSharedConversionSettings_(ensureSharedSettingsSheet_(ss))
    };
  } catch (e) {
    Logger.log('Error loading shared conversion settings: ' + e.toString());
    return {
      success: false,
      message: 'Error loading shared conversion settings: ' + e.toString()
    };
  }
}

function saveSharedConversionSettings(settings) {
  try {
    var configSheetId = getConfigSheetId();
    if (!configSheetId) {
      return {
        success: false,
        message: 'No config sheet linked. Please create or link one first.'
      };
    }

    var ss = SpreadsheetApp.openById(configSheetId);
    if (!ss.getSheetByName(CONFIGS_SHEET_NAME)) {
      return {
        success: false,
        message: 'Configs tab not found in the linked sheet.'
      };
    }

    return {
      success: true,
      message: 'Shared conversion settings saved.',
      settings: writeSharedConversionSettings_(ensureSharedSettingsSheet_(ss), settings || {})
    };
  } catch (e) {
    Logger.log('Error saving shared conversion settings: ' + e.toString());
    return {
      success: false,
      message: 'Error saving shared conversion settings: ' + e.toString()
    };
  }
}


// ============================================
// CENTRALIZED CONFIG SHEET FUNCTIONS
// ============================================

/**
 * Get the Config Sheet ID from user properties
 */
function getConfigSheetId() {
  var props = PropertiesService.getUserProperties();
  return props.getProperty('configSheetId');
}

/**
 * Extract Spreadsheet ID from either a raw ID or a full Google Sheets URL
 */
function extractSpreadsheetId(sheetInput) {
  if (!sheetInput) {
    return '';
  }

  var trimmedInput = String(sheetInput).trim();

  // Match IDs from standard Google Sheets URLs: /spreadsheets/d/{ID}
  var urlMatch = trimmedInput.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  if (urlMatch && urlMatch[1]) {
    return urlMatch[1];
  }

  // Fallback: assume the user pasted a raw spreadsheet ID
  return trimmedInput;
}

/**
 * Set the Config Sheet ID
 */
function setConfigSheetId(sheetId) {
  try {
    var normalizedSheetId = extractSpreadsheetId(sheetId);

    if (!normalizedSheetId) {
      return {
        success: false,
        message: 'Please provide a valid Google Sheet ID or URL.'
      };
    }

    // Validate the sheet exists and is accessible
    var sheet = SpreadsheetApp.openById(normalizedSheetId);
    ensureSharedSettingsSheet_(sheet);
    
    var props = PropertiesService.getUserProperties();
    props.setProperty('configSheetId', normalizedSheetId);
    
    return { 
      success: true, 
      message: 'Config sheet linked successfully!',
      sheetId: normalizedSheetId,
      sheetName: sheet.getName(),
      sheetUrl: sheet.getUrl()
    };
  } catch (e) {
    return { 
      success: false, 
      message: 'Error: Cannot access sheet. Make sure you have edit permission. ' + e.toString() 
    };
  }
}

/**
 * Create a new centralized config sheet
 */
function createConfigSheet() {
  try {
    // Create new spreadsheet
    var ss = SpreadsheetApp.create('Screenshot Configs Database');
    var sheet = ss.getActiveSheet();
    sheet.setName(CONFIGS_SHEET_NAME);
    
    // Set up headers
    var headers = [
      'Config ID',
      'Name', 
      'Spreadsheet ID',
      'Spreadsheet Name',
      'Sheet Tab',
      'Cell Range',
      'Created',
      'Updated',
      'Settings JSON'
    ];
    
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // Format header row
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground('#4285f4')
      .setFontColor('#ffffff')
      .setFontWeight('bold')
      .setWrap(true);
    
    // Set column widths
    sheet.setColumnWidth(1, 150); // Config ID
    sheet.setColumnWidth(2, 200); // Name
    sheet.setColumnWidth(3, 200); // Spreadsheet ID
    sheet.setColumnWidth(4, 200); // Spreadsheet Name
    sheet.setColumnWidth(5, 120); // Sheet Tab
    sheet.setColumnWidth(6, 120); // Cell Range
    sheet.setColumnWidth(7, 150); // Created
    sheet.setColumnWidth(8, 150); // Updated
    sheet.setColumnWidth(9, 300); // Settings JSON
    
    // Freeze header row
    sheet.setFrozenRows(1);

    ensureSharedSettingsSheet_(ss);
    
    // Auto-link this sheet
    setConfigSheetId(ss.getId());
    
    Logger.log('Created config sheet: ' + ss.getUrl());
    
    return {
      success: true,
      message: 'Config sheet created and linked successfully!',
      sheetId: ss.getId(),
      sheetUrl: ss.getUrl(),
      sheetName: ss.getName()
    };
    
  } catch (e) {
    Logger.log('Error creating config sheet: ' + e.toString());
    return {
      success: false,
      message: 'Error creating config sheet: ' + e.toString()
    };
  }
}

/**
 * Check if config sheet is linked and accessible
 */
function checkConfigSheet() {
  var configSheetId = getConfigSheetId();
  
  if (!configSheetId) {
    return {
      linked: false,
      message: 'No config sheet linked. Please create or link one.'
    };
  }
  
  try {
    var ss = SpreadsheetApp.openById(configSheetId);
    var sheet = ss.getSheetByName(CONFIGS_SHEET_NAME);
    
    if (!sheet) {
      return {
        linked: false,
        sheetId: configSheetId, // <-- Added this
        message: 'Config sheet found but "Configs" tab is missing. Please recreate.'
      };
    }
    
    ensureSharedSettingsSheet_(ss);

    return {
      linked: true,
      sheetId: configSheetId, // <-- Added this
      message: 'Config sheet is linked and accessible',
      sheetName: ss.getName(),
      sheetUrl: ss.getUrl(),
      configCount: Math.max(0, sheet.getLastRow() - 1)
    };
    
  } catch (e) {
    return {
      linked: false,
      sheetId: configSheetId, // <-- Added this
      message: 'Cannot access config sheet. Check permissions or recreate.',
      error: e.toString()
    };
  }
}

/**
 * Save config to centralized sheet (replaces old saveConfig)
 */
function saveConfig(configName, settings) {
  try {
    var configSheetId = getConfigSheetId();
    
    if (!configSheetId) {
      return { 
        success: false, 
        message: 'No config sheet linked. Please create or link one first.' 
      };
    }
    
    var configSS = SpreadsheetApp.openById(configSheetId);
    var sheet = configSS.getSheetByName(CONFIGS_SHEET_NAME);
    
    if (!sheet) {
      return { 
        success: false, 
        message: 'Configs tab not found in the linked sheet.' 
      };
    }
    
    var currentSS = SpreadsheetApp.getActiveSpreadsheet();
    var spreadsheetId = currentSS.getId();
    var spreadsheetName = currentSS.getName();
    
    // Check for existing config with same tab, range, spreadsheet AND NAME
    var data = sheet.getDataRange().getValues();
    var existingRow = -1;
    
    for (var i = 1; i < data.length; i++) {
      var rowConfigName = data[i][1];      // Column B
      var rowSpreadsheetId = data[i][2];   // Column C
      var rowSheetTab = data[i][4];        // Column E
      var rowRange = data[i][5];           // Column F
      
      // FIX: Added rowConfigName === configName so filters don't overwrite each other
      if (rowSpreadsheetId === spreadsheetId &&
          rowSheetTab === settings.sheetTabs &&
          rowRange === settings.cellRanges &&
          rowConfigName === configName) {
        existingRow = i + 1;
        break;
      }
    }
    
    var now = new Date().toISOString();
    var settingsJson = JSON.stringify(settings);
    
    if (existingRow > 0) {
      // Update existing config
      var configId = data[existingRow - 1][0];
      
      sheet.getRange(existingRow, 2).setValue(configName);
      sheet.getRange(existingRow, 4).setValue(spreadsheetName);
      sheet.getRange(existingRow, 8).setValue(now);
      sheet.getRange(existingRow, 9).setValue(settingsJson);
      
      Logger.log('Updated existing config: ' + configName + ' (ID: ' + configId + ')');
      
      return {
        success: true,
        configId: configId,
        message: 'Config "' + configName + '" updated (duplicate prevented)',
        updated: true
      };
    } else {
      // FIX: Added a random number to the ID to ensure fast loops don't generate identical IDs in the same millisecond
      var configId = 'config_' + new Date().getTime() + '_' + Math.floor(Math.random() * 1000);
      
      var newRow = [
        configId,
        configName,
        spreadsheetId,
        spreadsheetName,
        settings.sheetTabs,
        settings.cellRanges,
        now,
        now,
        settingsJson
      ];
      
      sheet.appendRow(newRow);
      
      Logger.log('Created new config: ' + configName + ' (ID: ' + configId + ')');
      
      return {
        success: true,
        configId: configId,
        message: 'Config "' + configName + '" saved successfully!',
        updated: false
      };
    }
    
  } catch (e) {
    Logger.log('Error saving config: ' + e.toString());
    return { 
      success: false, 
      message: 'Error saving config: ' + e.toString() 
    };
  }
}

/**
 * Load config by ID from centralized sheet
 */
function loadConfig(configId) {
  try {
    var configSheetId = getConfigSheetId();
    
    if (!configSheetId) {
      return { 
        success: false, 
        message: 'No config sheet linked.' 
      };
    }
    
    var configSS = SpreadsheetApp.openById(configSheetId);
    var sheet = configSS.getSheetByName(CONFIGS_SHEET_NAME);
    
    if (!sheet) {
      return { 
        success: false, 
        message: 'Configs tab not found.' 
      };
    }
    
    var data = sheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === configId) {
        var config = {
          id: data[i][0],
          name: data[i][1],
          spreadsheetId: data[i][2],
          spreadsheetName: data[i][3],
          settings: JSON.parse(data[i][8]),
          createdAt: data[i][6],
          updatedAt: data[i][7]
        };
        
        return { 
          success: true, 
          config: config 
        };
      }
    }
    
    return { 
      success: false, 
      message: 'Config not found' 
    };
    
  } catch (e) {
    Logger.log('Error loading config: ' + e.toString());
    return { 
      success: false, 
      message: 'Error loading config: ' + e.toString() 
    };
  }
}

/**
 * Get list of all configs from centralized sheet
 */
function getConfigList() {
  try {
    var configSheetId = getConfigSheetId();
    
    if (!configSheetId) {
      return [];
    }
    
    var configSS = SpreadsheetApp.openById(configSheetId);
    var sheet = configSS.getSheetByName(CONFIGS_SHEET_NAME);
    
    if (!sheet || sheet.getLastRow() < 2) {
      return [];
    }
    
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getValues();
    var configs = [];
    
    for (var i = 0; i < data.length; i++) {
      if (data[i][0]) {
        configs.push({
          id: data[i][0],
          name: data[i][1],
          spreadsheetId: data[i][2],
          spreadsheetName: data[i][3],
          sheetTab: data[i][4],
          range: data[i][5],
          createdAt: data[i][6],
          updatedAt: data[i][7]
        });
      }
    }
    
    return configs;
    
  } catch (e) {
    Logger.log('Error getting config list: ' + e.toString());
    return [];
  }
}

/**
 * Delete config from centralized sheet
 */
function deleteConfig(configId) {
  try {
    var configSheetId = getConfigSheetId();
    
    if (!configSheetId) {
      return { 
        success: false, 
        message: 'No config sheet linked.' 
      };
    }
    
    var configSS = SpreadsheetApp.openById(configSheetId);
    var sheet = configSS.getSheetByName(CONFIGS_SHEET_NAME);
    
    if (!sheet) {
      return { 
        success: false, 
        message: 'Configs tab not found.' 
      };
    }
    
    var data = sheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === configId) {
        sheet.deleteRow(i + 1);
        return { 
          success: true, 
          message: 'Config deleted successfully!' 
        };
      }
    }
    
    return { 
      success: false, 
      message: 'Config not found' 
    };
    
  } catch (e) {
    Logger.log('Error deleting config: ' + e.toString());
    return { 
      success: false, 
      message: 'Error deleting config: ' + e.toString() 
    };
  }
}


/**
 * Reads data validation dropdown options from a specified cell.
 * Returns the list of valid values so the UI can render checkboxes.
 */
function getFilterDropdownOptions(cellRef) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();
    const range = sheet.getRange(cellRef);
    const validation = range.getDataValidation();

    if (!validation) {
      return { success: false, error: 'No data validation found on cell ' + cellRef };
    }

    const criteria = validation.getCriteriaType();
    if (criteria !== SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST &&
        criteria !== SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE) {
      return { success: false, error: 'Cell ' + cellRef + ' does not have a dropdown validation. Found: ' + criteria };
    }

    var values = [];
    if (criteria === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE) {
      var rangeValue = validation.getCriteriaValues()[0];
      var valuesRange = ss.getRange(rangeValue.getA1Notation());
      var valuesData = valuesRange.getValues();
      for (var i = 0; i < valuesData.length; i++) {
        if (valuesData[i][0] && valuesData[i][0].toString().trim()) {
          values.push(valuesData[i][0].toString().trim());
        }
      }
    } else {
      var criteriaValues = validation.getCriteriaValues();
      values = criteriaValues[0].map(function(v) { return v.toString().trim(); }).filter(function(v) { return v; });
    }

    return { success: true, values: values, cell: cellRef };
  } catch (e) {
    Logger.log('Error reading dropdown options: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

function runExportWithSettings(settings) {
  try {
    const ss = settings.sheetId 
      ? SpreadsheetApp.openById(settings.sheetId) 
      : SpreadsheetApp.getActiveSpreadsheet();
    
    // Parse sheet tabs
    const sheetTabs = settings.sheetTabs.split(",").map(s => s.trim()).filter(s => s);
    if (sheetTabs.length === 0) {
      return "❌ Error: Please specify at least one sheet tab";
    }
    
    // Parse cell ranges
    const cellRanges = settings.cellRanges.split(",").map(r => r.trim()).filter(r => r);
    if (cellRanges.length === 0) {
      return "❌ Error: Please specify at least one cell range";
    }

    // NEW: Parse custom names
    const customNames = settings.customNames ? settings.customNames.split(",").map(n => n.trim()).filter(n => n) : [];
    
    // Validate all sheets exist
    for (const tabName of sheetTabs) {
      const sheet = ss.getSheetByName(tabName);
      if (!sheet) {
        return `❌ Error: Sheet tab '${tabName}' not found!`;
      }
    }
    
    let configCount = 0;
    let configErrors = [];

    // Parse filter settings for metadata-only config generation.
    // We intentionally do not write these values into sheet cells.
    // Supports N dynamic filter cells via filterCells array, with fallback to
    // legacy filter1Cell/filter1Values/filter2Cell/filter2Values for backward compat.
    var filterCells = [];

    if (settings.filterEnabled) {
      if (settings.filterCells && settings.filterCells.length > 0) {
        // New format: array of { cell, values: [...] }
        for (var fc = 0; fc < settings.filterCells.length; fc++) {
          var fcObj = settings.filterCells[fc];
          if (fcObj.cell && fcObj.values && fcObj.values.length > 0) {
            filterCells.push(fcObj);
          }
        }
      } else {
        // Legacy format: build from filter1Cell/filter1Values, filter2Cell/filter2Values
        if (settings.filter1Values && settings.filter1Values.trim()) {
          var vals1 = settings.filter1Values.split(",").map(function(v) { return v.trim(); }).filter(function(v) { return v; });
          if (vals1.length > 0) {
            filterCells.push({ cell: settings.filter1Cell || 'Filter1', values: vals1 });
          }
        }
        if (settings.filter2Values && settings.filter2Values.trim()) {
          var vals2 = settings.filter2Values.split(",").map(function(v) { return v.trim(); }).filter(function(v) { return v; });
          if (vals2.length > 0) {
            filterCells.push({ cell: settings.filter2Cell || 'Filter2', values: vals2 });
          }
        }
      }

      if (filterCells.length === 0) {
        return "❌ Error: At least one filter cell with values is required when filtering is enabled";
      }
    }

    // Generate Cartesian product of all filter cell values
    function buildFilterCombinations(cells) {
      if (!cells || cells.length === 0) return [[]];
      return cells.reduce(function(acc, cell) {
        var result = [];
        acc.forEach(function(existing) {
          cell.values.forEach(function(val) {
            result.push(existing.concat([val]));
          });
        });
        return result;
      }, [[]]);
    }

    var combinations = settings.filterEnabled ? buildFilterCombinations(filterCells) : [[]];
    
    // Iterate through each sheet tab
    sheetTabs.forEach((tabName, tabIndex) => {
      // Determine which range to use for this tab
      const rangeIndex = cellRanges.length === 1 ? 0 : Math.min(tabIndex, cellRanges.length - 1);
      
      // Keep track of the original input so we know if it was a named range
      let originalRangeInput = cellRanges[rangeIndex]; 
      let cellRange = originalRangeInput;
      let isNamedRange = false;
      
      // Check if it's a named range and resolve it
      const resolvedRange = resolveNamedRange(ss, cellRange);
      if (resolvedRange) {
        Logger.log(`Resolved named range "${cellRange}" → ${resolvedRange}`);
        cellRange = resolvedRange;
        isNamedRange = true;
      } else if (originalRangeInput.indexOf(':') === -1) {
        // If it failed to resolve but doesn't have a colon, it's likely intended as a named range or single cell
        isNamedRange = true;
      }
  
      // ==========================================
      // NEW: DETERMINE BASE NAME
      // ==========================================
      let baseName = tabName; // Default to just the tab name
      let providedCustomName = customNames.length > 0 ? (customNames[tabIndex] || customNames[0]) : null;

      if (providedCustomName) {
        baseName = providedCustomName; // User provided a custom name
      } else if (isNamedRange) {
        baseName = `${originalRangeInput} - ${tabName}`; // It's a named range, e.g., "SalesData - Sheet1"
      }

      Logger.log(`Processing config for tab: ${tabName} with baseName: ${baseName} (Range: ${cellRange})`);

      combinations.forEach(function(combination) {
        var fileName = baseName;

        if (settings.filterEnabled && combination.length > 0) {
          fileName = combination.join(' - ') + ' - ' + baseName;
        }

        var configSettings = JSON.parse(JSON.stringify(settings));
        if (configSettings.filterEnabled) {
          // Store per-combination filterCells array with single values
          configSettings.filterCells = filterCells.map(function(fc, idx) {
            return { cell: fc.cell, values: combination[idx] || '' };
          });
          // Backward compat: continue populating legacy fields
          configSettings.filter1Cell = filterCells[0] ? filterCells[0].cell : '';
          configSettings.filter1Values = combination[0] || '';
          configSettings.filter2Cell = filterCells[1] ? filterCells[1].cell : '';
          configSettings.filter2Values = combination[1] || '';
        }

        try {
          // Keep originalRangeInput so named ranges stay readable in the config sheet.
          const result = autoSaveConfig(fileName, configSettings, tabName, originalRangeInput);

          if (result) {
            configCount++;
            Logger.log("✓ Config " + configCount + ": " + fileName);
          } else {
            configErrors.push("Failed to save config: " + fileName);
          }
        } catch (err) {
          const errorMsg = "Failed: " + fileName + " - " + err.message;
          Logger.log("✗ " + errorMsg);
          configErrors.push(errorMsg);
        }
      });
    });
    
    let message = "✅ Success! Saved " + configCount + " configuration(s).";
    
    if (configErrors.length > 0) {
      message += "\n\n⚠️ " + configErrors.length + " error(s):";
      message += "\n" + configErrors.slice(0, 5).join("\n");
      if (configErrors.length > 5) {
        message += "\n... and " + (configErrors.length - 5) + " more (check logs)";
      }
    }
    
    return message;
    
  } catch (e) {
    Logger.log("FATAL ERROR: " + e.toString());
    Logger.log("Stack: " + e.stack);
    return "❌ Error: " + e.toString();
  }
}

// ============================================
// CONFIG MANAGEMENT FUNCTIONS
// ============================================

/**
 * Save current configuration with a name
 */

/**
 * Load a saved configuration
 */

/**
 * Get list of all saved configs
 */

/**
 * Delete a saved configuration
 */



/**
 * Resolves named range to actual cell coordinates
 * Returns null if not a named range, or the A1 notation if it is
 */
function resolveNamedRange(ss, rangeName) {
  Logger.log('🔍 resolveNamedRange called with: "' + rangeName + '"');
  
  try {
    // First check if it looks like a cell range (contains ':')
    if (rangeName.indexOf(':') > -1) {
      Logger.log('  → Contains ":", treating as regular cell range');
      return null;
    }
    
    Logger.log('  → Attempting to get named range...');
    const namedRange = ss.getRangeByName(rangeName);
    
    if (!namedRange) {
      Logger.log('  → Not a named range, returning null');
      return null; // Not a named range, treat as regular cell reference
    }
    
    // Return the A1 notation (e.g., "D10:AD62")
    const a1Notation = namedRange.getA1Notation();
    Logger.log('  ✅ Named range found! Resolved to: ' + a1Notation);
    return a1Notation;
    
  } catch (e) {
    Logger.log('  ⚠️ Error resolving named range "' + rangeName + '": ' + e.toString());
    return null;
  }
}


/**
 * Get Google Identity Token for Cloud Run
 * This exchanges OAuth token for Identity token
 */
function getCloudRunIdentityToken(targetAudience) {
  try {
    // Method 1: Try built-in identity token (if available)
    try {
      var identityToken = ScriptApp.getIdentityToken();
      if (identityToken) {
        Logger.log("Using ScriptApp.getIdentityToken()");
        return identityToken;
      }
    } catch (e) {
      Logger.log("ScriptApp.getIdentityToken() not available: " + e.toString());
    }
    
    // Method 2: Exchange OAuth token for Identity token
    var oauthToken = ScriptApp.getOAuthToken();
    
    // Use Google's token exchange endpoint
    var payload = {
      'grant_type': 'urn:ietf:params:oauth:grant-type:token-exchange',
      'subject_token_type': 'urn:ietf:params:oauth:token-type:access_token',
      'requested_token_type': 'urn:ietf:params:oauth:token-type:id_token',
      'subject_token': oauthToken,
      'audience': targetAudience
    };
    
    var options = {
      'method': 'post',
      'contentType': 'application/x-www-form-urlencoded',
      'payload': payload,
      'muteHttpExceptions': true
    };
    
    var response = UrlFetchApp.fetch('https://sts.googleapis.com/v1/token', options);
    var responseCode = response.getResponseCode();
    
    Logger.log("STS response code: " + responseCode);
    
    if (responseCode === 200) {
      var result = JSON.parse(response.getContentText());
      if (result.access_token) {
        Logger.log("Successfully exchanged for identity token");
        return result.access_token;
      }
    }
    
    Logger.log("STS response: " + response.getContentText());
    
    // Method 3: Try using OAuth token directly (fallback)
    Logger.log("Falling back to OAuth token");
    return oauthToken;
    
  } catch (e) {
    Logger.log("Error getting identity token: " + e.toString());
    throw new Error("Failed to get authentication token: " + e.toString());
  }
}

/**
 * Returns all sheet tab names in the active spreadsheet.
 */
function getSheetTabs() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheets().map(function(s) { return s.getName(); });
}

/**
 * Returns named ranges that belong to a specific sheet tab.
 */
function getSheetNamedRanges(sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var namedRanges = ss.getNamedRanges();
  return namedRanges
    .filter(function(nr) {
      return nr.getRange().getSheet().getName() === sheetName;
    })
    .map(function(nr) {
      return {
        name: nr.getName(),
        a1Notation: nr.getRange().getA1Notation()
      };
    });
}

/**
 * Returns the active sheet name and selected range A1 notation.
 */
function getActiveCellRange() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var range = sheet.getActiveRange();
  return {
    sheetName: sheet.getName(),
    rangeA1: range.getA1Notation()
  };
}


