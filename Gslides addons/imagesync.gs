/**

 * STANDALONE: Google Slides Image Replacement Tool

 * Creates a sidebar to sync/replace images in slides

 */

// ============================================
// SHARED CONVERSION SETTINGS
// ============================================
var CONFIGS_SHEET_NAME = 'Configs';
var SHARED_SETTINGS_SHEET_NAME = 'Settings';
var SHARED_SETTINGS_HEADERS = ['Key', 'Value', 'Updated'];
var SHARED_SETTINGS_DEFAULTS = {
  gcsInputBucket: '',
  gcsOutputBucket: '',
  cloudConvertApiKey: ''
};
var SYNC_METADATA_TAG = '##SYNC##';
var CONFIG_LINK_TAG = '##S2S_CONFIG##';

// Milliseconds to wait after SpreadsheetApp.flush() when a filter is applied,
// to allow Google Sheets formulas to fully recalculate before the PDF export.
// Increase this if you still see stale data in exported PDFs.
var FILTER_FLUSH_DELAY_MS = 3000;

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

function escapeRegex(text) {
  return String(text).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function getElementDescription(elementOrDescription) {
  try {
    if (!elementOrDescription) {
      return '';
    }

    if (typeof elementOrDescription === 'string') {
      return elementOrDescription;
    }

    if (typeof elementOrDescription.getDescription === 'function') {
      return elementOrDescription.getDescription() || '';
    }
  } catch (e) {
    Logger.log('Error reading description: ' + e.toString());
  }

  return '';
}

function extractTaggedPayload(elementOrDescription, tag) {
  var description = getElementDescription(elementOrDescription);
  if (!description) {
    return null;
  }

  var match = description.match(new RegExp(escapeRegex(tag) + '([\\s\\S]*?)' + escapeRegex(tag)));
  return match && match[1] ? match[1] : null;
}

function stripTaggedPayload(description, tag) {
  if (!description) {
    return '';
  }

  return String(description)
    .replace(new RegExp(escapeRegex(tag) + '[\\s\\S]*?' + escapeRegex(tag), 'g'), '')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

function upsertTaggedPayload(description, tag, payload) {
  var cleanedDescription = stripTaggedPayload(description, tag);
  var taggedPayload = tag + payload + tag;

  return cleanedDescription ? cleanedDescription + '\n' + taggedPayload : taggedPayload;
}

function findImageContextByObjectId(objectId, presentation) {
  var activePresentation = presentation || SlidesApp.getActivePresentation();
  var slides = activePresentation.getSlides();

  for (var i = 0; i < slides.length; i++) {
    var images = slides[i].getImages();
    for (var j = 0; j < images.length; j++) {
      if (images[j].getObjectId() === objectId) {
        return {
          presentation: activePresentation,
          slide: slides[i],
          slideId: slides[i].getObjectId(),
          image: images[j],
          objectId: objectId
        };
      }
    }
  }

  return null;
}

function getLegacyConfigReference(presentationId, slideId, objectId) {
  try {
    var userProps = PropertiesService.getUserProperties();
    var directLink = userProps.getProperty('imageLink_' + presentationId + '_' + objectId);
    if (directLink) {
      return String(directLink);
    }

    var metadataLink = userProps.getProperty(presentationId + '_' + slideId + '_' + objectId + '_config');
    if (metadataLink) {
      try {
        var parsedMetadataLink = JSON.parse(metadataLink);
        if (parsedMetadataLink && parsedMetadataLink.configId) {
          return String(parsedMetadataLink.configId);
        }
      } catch (parseError) {
        return String(metadataLink);
      }
    }

    var documentProps = PropertiesService.getDocumentProperties();
    var documentLink = documentProps.getProperty('imageConfig_' + objectId);
    if (documentLink) {
      return String(documentLink);
    }
  } catch (e) {
    Logger.log('Error reading legacy config reference: ' + e.toString());
  }

  return null;
}

function clearLegacyConfigReference(presentationId, slideId, objectId) {
  try {
    var userProps = PropertiesService.getUserProperties();
    userProps.deleteProperty('imageLink_' + presentationId + '_' + objectId);
    if (slideId) {
      userProps.deleteProperty(presentationId + '_' + slideId + '_' + objectId + '_config');
    }

    var documentProps = PropertiesService.getDocumentProperties();
    documentProps.deleteProperty('imageConfig_' + objectId);
  } catch (e) {
    Logger.log('Error clearing legacy config reference: ' + e.toString());
  }
}

function getConfigIdFromTagPayload(payload) {
  if (!payload) {
    return null;
  }

  try {
    var parsedPayload = JSON.parse(payload);
    if (parsedPayload && parsedPayload.configId) {
      return String(parsedPayload.configId);
    }
  } catch (e) {
    return String(payload);
  }

  return null;
}


function onOpen(e) {

  SlidesApp.getUi()

    .createMenu('Sheets2Slides')

    .addItem('Open Sync Panel', 'showImageSyncSidebar')

    .addToUi();

}



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
  var urlMatch = trimmedInput.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);

  if (urlMatch && urlMatch[1]) {
    return urlMatch[1];
  }

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

        message: 'Config sheet found but "Configs" tab is missing. Please recreate.'

      };

    }

    

    ensureSharedSettingsSheet_(ss);

    return {

      linked: true,

      message: 'Config sheet is linked and accessible',

      sheetName: ss.getName(),

      sheetUrl: ss.getUrl(),

      configCount: Math.max(0, sheet.getLastRow() - 1)

    };

    

  } catch (e) {

    return {

      linked: false,

      message: 'Cannot access config sheet. Check permissions or recreate.',

      error: e.toString()

    };

  }

}





function showImageSyncSidebar() {

  const html = HtmlService.createHtmlOutputFromFile('ImageSyncSidebar')

    .setTitle('Sheets2Slides Sync Manager')

    .setWidth(350);

  SlidesApp.getUi().showSidebar(html);

}

// ============================================

// IMAGE SYNC FUNCTIONS - FIXED FOR PNG

// ============================================



/**

 * Sync image from config - CORRECTED VERSION with PNG conversion

 */

/**

 * Sync image from config with explicit workflow:

 * 1. Export PDF from Sheets

 * 2. Send to API and get PNG

 * 3. Save PNG temporarily to Google Drive

 * 4. Replace image in Slides

 * 5. Delete temp file from Drive

 */

function syncImageFromConfig(objectId, configId) {
  try {
    Logger.log('=== STARTING IMAGE SYNC ===');
    Logger.log('Old Object ID: ' + objectId);
    Logger.log('Config ID: ' + configId);
    
    var presentation = SlidesApp.getActivePresentation();
    var presentationId = presentation.getId();
    
    // Find the image across all slides
    var slides = presentation.getSlides();
    var targetSlide = null;
    var targetImage = null;
    var slideIndex = -1;
    
    for (var i = 0; i < slides.length; i++) {
      var images = slides[i].getImages();
      for (var j = 0; j < images.length; j++) {
        if (images[j].getObjectId() === objectId) {
          targetSlide = slides[i];
          targetImage = images[j];
          slideIndex = i;
          Logger.log('✓ Found image on slide ' + (i + 1));
          break;
        }
      }
      if (targetImage) break;
    }
    
    if (!targetImage) {
      throw new Error('Image not found with ID: ' + objectId);
    }
    
    // Load config
    Logger.log('Loading config: ' + configId);
    var configResult = loadConfigById(configId);
    if (!configResult.success) {
      throw new Error('Failed to load config: ' + configResult.message);
    }
    
    var config = configResult.config;
    Logger.log('✓ Config loaded: ' + config.name);
    
    // Get image properties before removal
    var width = targetImage.getWidth();
    var height = targetImage.getHeight();
    var left = targetImage.getLeft();
    var top = targetImage.getTop();
    var title = targetImage.getTitle() || '';
    
    Logger.log('Image position: (' + left + ', ' + top + ')');
    Logger.log('Image size: ' + width + ' x ' + height);
    
    // === STEP 1: Export PDF from Google Sheets ===
    Logger.log('\n--- STEP 1: Exporting PDF from Sheets ---');
    var settings = config.settings;
    var ss = SpreadsheetApp.openById(config.spreadsheetId);
    var sheet = ss.getSheetByName(settings.sheetTabs);
    
    if (!sheet) {
      throw new Error('Sheet not found: ' + settings.sheetTabs);
    }

    // ==========================================
    // NEW: APPLY FILTERS BEFORE EXPORTING
    // ==========================================
    if (settings.filterEnabled) {
      Logger.log('Applying config filters to sheet...');
      if (settings.filter1Cell && settings.filter1Values) {
        sheet.getRange(settings.filter1Cell).setValue(settings.filter1Values);
        Logger.log('Set ' + settings.filter1Cell + ' = ' + settings.filter1Values);
      }
      if (settings.filter2Cell && settings.filter2Values) {
        sheet.getRange(settings.filter2Cell).setValue(settings.filter2Values);
        Logger.log('Set ' + settings.filter2Cell + ' = ' + settings.filter2Values);
      }
      
      // Force Google Sheets to recalculate formulas with the new filter values,
      // then wait for FILTER_FLUSH_DELAY_MS to let slow/array formulas settle.
      SpreadsheetApp.flush();
      Utilities.sleep(FILTER_FLUSH_DELAY_MS);
      Logger.log('Filter flush complete, waited ' + FILTER_FLUSH_DELAY_MS + 'ms for recalculation');
    }
    
    // Parse cell range
    var cellRange = settings.cellRanges;
    Logger.log('📍 Original cellRange from config: "' + cellRange + '"');

    var resolvedRange = resolveNamedRange(ss, cellRange);
    if (resolvedRange) {
      Logger.log('✅ Resolved named range "' + cellRange + '" → ' + resolvedRange);
      cellRange = resolvedRange;
    } else {
      Logger.log('ℹ️ Not a named range, using as-is: ' + cellRange);
    }
    var rangeParts = cellRange.split(':');
    if (rangeParts.length !== 2) {
      throw new Error('Invalid range format: ' + cellRange);
    }
    
    var startCell = parseCellReference(rangeParts[0]);
    var endCell = parseCellReference(rangeParts[1]);
    
    var pdfRange = {
      c1: startCell.col,
      r1: startCell.row,
      c2: endCell.col,
      r2: endCell.row
    };
    
    var pdfBlob = generateSnapshotBlob(ss.getId(), sheet.getSheetId(), pdfRange, config.name);
    Logger.log('✓ PDF exported: ' + pdfBlob.getBytes().length + ' bytes');
    
    // === STEP 2: Send PDF to API and get PNG ===
    Logger.log('\n--- STEP 2: Converting PDF to PNG via API ---');
    var pngSettings = settings.pngSettings || {
      dpi: 300,
      quality: 95,
      background: 'white',
      autoCrop: false,
      cropLeft: 0,
      cropTop: 0,
      cropRight: 0,
      cropBottom: 0
    };

    var sharedSettings = getSharedConversionSettings();
    if (!sharedSettings.success) {
      throw new Error(sharedSettings.message || 'Shared conversion settings are unavailable');
    }

    var runtimeSettings = sharedSettings.settings || {};
    
    var conversionMethod = settings.pngConversionMethod || 'internal';
    Logger.log('Conversion method: ' + conversionMethod);
    
    var pngBlob;
    if (conversionMethod === 'cloudconvert') {
      if (!runtimeSettings.cloudConvertApiKey) {
        throw new Error('CloudConvert API key is missing in shared conversion settings.');
      }
      Logger.log('Using CloudConvert API...');
      pngBlob = convertPdfToPngCloudConvert(pdfBlob, config.name, runtimeSettings.cloudConvertApiKey, pngSettings);
    } else {
      if (!runtimeSettings.gcsInputBucket || !runtimeSettings.gcsOutputBucket) {
        throw new Error('Input and output buckets are missing in shared conversion settings.');
      }

      Logger.log('Using GCS Pub/Sub pipeline (internal)...');
      pngBlob = convertPdfToPngViaGcs(
        pdfBlob,
        config.name,
        pngSettings,
        runtimeSettings.gcsInputBucket,
        runtimeSettings.gcsOutputBucket
      );
    }
    
    Logger.log('✓ PNG received: ' + pngBlob.getBytes().length + ' bytes');
    Logger.log('  Content type: ' + pngBlob.getContentType());
    
    // Verify it's actually PNG
    if (pngBlob.getContentType().indexOf('pdf') !== -1) {
      throw new Error('ERROR: API returned PDF instead of PNG!');
    }
    
    // === STEP 3: Save PNG temporarily to Google Drive ===
    Logger.log('\n--- STEP 3: Saving PNG to Google Drive ---');
    var tempFileName = 'temp_sync_' + new Date().getTime() + '.png';
    var tempPngFile = DriveApp.createFile(pngBlob.setName(tempFileName));
    Logger.log('✓ PNG saved to Drive');
    Logger.log('  File ID: ' + tempPngFile.getId());
    
    var newObjectId = null;
    
    try {
      // === STEP 4: Replace image in Slides ===
      Logger.log('\n--- STEP 4: Replacing image in Slides ---');
      
      // Remove old image
      Logger.log('Removing old image (ID: ' + objectId + ')...');
      targetImage.remove();
      
      // Insert new PNG from Drive
      Logger.log('Inserting new PNG from Drive...');
      var newImage = targetSlide.insertImage(tempPngFile, left, top, width, height);
      newImage.setTitle(title);
      newObjectId = newImage.getObjectId();
      Logger.log('✓ Image replaced with new ID: ' + newObjectId);
      
      // Store metadata in image description
      var metadata = {
        configId: configId,
        configName: config.name,
        lastSync: new Date().toISOString(),
        sourceType: 'config',
        spreadsheetId: config.spreadsheetId,
        sheetTab: config.settings.sheetTabs,
        cellRange: config.settings.cellRanges,
        slideIndex: slideIndex,
        position: { left: left, top: top }
      };
      
      storeConfigReference(newImage, configId);
      storeSyncMetadata(newImage, metadata);
      clearLegacyConfigReference(presentationId, targetSlide.getObjectId(), objectId);
      Logger.log('✓ Preserved config link in image description for new object ID: ' + newObjectId);
      
    } finally {
      // === STEP 5: Delete temporary PNG from Drive ===
      Logger.log('\n--- STEP 5: Cleaning up Drive ---');
      tempPngFile.setTrashed(true);
      Logger.log('✓ Temporary file deleted');
    }
    
    Logger.log('\n=== SYNC COMPLETE ===');
    Logger.log('Config linked to new object ID: ' + newObjectId);
    
    return {
      success: true,
      message: 'Successfully synced: ' + config.name,
      oldObjectId: objectId,
      newObjectId: newObjectId
    };
    
  } catch (e) {
    Logger.log('\n❌ SYNC FAILED');
    Logger.log('Error: ' + e.toString());
    Logger.log('Stack: ' + e.stack);
    return {
      success: false,
      message: 'Sync failed: ' + e.toString()
    };
  }
}



function recoverConfigLinksFromMetadata() {

  try {

    Logger.log('=== RECOVERING CONFIG LINKS ===');

    

    var presentation = SlidesApp.getActivePresentation();

    var slides = presentation.getSlides();

    var recovered = 0;

    var failed = 0;

    

    for (var i = 0; i < slides.length; i++) {

      var images = slides[i].getImages();

      

      for (var j = 0; j < images.length; j++) {

        var image = images[j];

        var objectId = image.getObjectId();

        

        // Try to get metadata from description

        var metadata = getSyncMetadata(image);

        

        if (metadata && metadata.configId) {

          // Check if config link exists

          var existingLink = getConfigReference(image);

          

          if (!existingLink || String(existingLink) !== String(metadata.configId)) {

            // Recover the link

            Logger.log('Recovering link for image on slide ' + (i + 1));

            Logger.log('  Object ID: ' + objectId);

            Logger.log('  Config ID: ' + metadata.configId);

            

            storeConfigReference(image, metadata.configId);

            clearLegacyConfigReference(presentation.getId(), slides[i].getObjectId(), objectId);

            

            recovered++;

          }

        }

      }

    }

    

    Logger.log('\n=== RECOVERY COMPLETE ===');

    Logger.log('Recovered: ' + recovered + ' links');

    

    return {

      success: true,

      recovered: recovered,

      message: 'Recovered ' + recovered + ' config links'

    };

    

  } catch (e) {

    Logger.log('Error in recovery: ' + e.toString());

    return {

      success: false,

      message: 'Recovery failed: ' + e.toString()

    };

  }

}

/**

 * Refresh all images that have config links

 * This syncs every image in the presentation that is linked to a config

 */

function refreshAllImages() {

  try {

    Logger.log('=== REFRESHING ALL IMAGES ===');

    

    var presentation = SlidesApp.getActivePresentation();

    var slides = presentation.getSlides();

    

    // Collect all images with config links

    var imagesToSync = [];

    

    for (var i = 0; i < slides.length; i++) {

      var images = slides[i].getImages();

      

      for (var j = 0; j < images.length; j++) {

        var image = images[j];

        var objectId = image.getObjectId();

        

        var configId = getConfigReference(image);

        

        if (configId) {

          imagesToSync.push({

            objectId: objectId,

            configId: configId,

            slideIndex: i,

            slideNumber: i + 1

          });

          Logger.log('Found linked image on slide ' + (i + 1) + ': ' + configId);

        }

      }

    }

    

    Logger.log('\nFound ' + imagesToSync.length + ' images with config links');

    

    if (imagesToSync.length === 0) {

      return {

        success: true,

        message: 'No images with config links found',

        synced: 0,

        failed: 0

      };

    }

    

    // Sync each image

    var synced = 0;

    var failed = 0;

    var errors = [];

    

    for (var k = 0; k < imagesToSync.length; k++) {

      var item = imagesToSync[k];

      

      Logger.log('\n--- Syncing image ' + (k + 1) + '/' + imagesToSync.length + ' ---');

      Logger.log('Slide: ' + item.slideNumber);

      Logger.log('Object ID: ' + item.objectId);

      Logger.log('Config ID: ' + item.configId);

      

      try {

        var result = syncImageFromConfig(item.objectId, item.configId);

        

        if (result.success) {

          synced++;

          Logger.log('✓ Synced successfully');

        } else {

          failed++;

          errors.push('Slide ' + item.slideNumber + ': ' + result.message);

          Logger.log('✗ Sync failed: ' + result.message);

        }

      } catch (e) {

        failed++;

        errors.push('Slide ' + item.slideNumber + ': ' + e.toString());

        Logger.log('✗ Exception: ' + e.toString());

      }

      

      // Small delay to avoid rate limiting

      if (k < imagesToSync.length - 1) {

        Utilities.sleep(500);

      }

    }

    

    Logger.log('\n=== REFRESH COMPLETE ===');

    Logger.log('Synced: ' + synced);

    Logger.log('Failed: ' + failed);

    

    var message = 'Refreshed ' + synced + ' image(s)';

    if (failed > 0) {

      message += '\n\nFailed: ' + failed + ' image(s)';

      if (errors.length > 0) {

        message += '\n' + errors.slice(0, 3).join('\n');

        if (errors.length > 3) {

          message += '\n... and ' + (errors.length - 3) + ' more errors';

        }

      }

    }

    

    return {

      success: true,

      message: message,

      synced: synced,

      failed: failed,

      total: imagesToSync.length,

      errors: errors

    };

    

  } catch (e) {

    Logger.log('\n❌ REFRESH ALL FAILED');

    Logger.log('Error: ' + e.toString());

    return {

      success: false,

      message: 'Refresh failed: ' + e.toString(),

      synced: 0,

      failed: 0

    };

  }

}


/**

 * Helper function to get config reference by object ID

 */

function getConfigReferenceByObjectId(objectId) {

  try {

    var context = findImageContextByObjectId(objectId);

    if (!context) {

      return null;

    }

    return getConfigReference(context.image);

  } catch (e) {

    Logger.log('Error getting config reference for ' + objectId + ': ' + e.toString());

    return null;

  }

}

function loadConfigById(configId) {

  try {

    Logger.log('Loading config by ID: ' + configId);

    

    var props = PropertiesService.getUserProperties();

    var configSheetId = props.getProperty('configSheetId');

    

    if (!configSheetId) {

      return { success: false, message: 'No config sheet linked' };

    }

    

    var configSS = SpreadsheetApp.openById(configSheetId);

    var sheet = configSS.getSheetByName(CONFIGS_SHEET_NAME);

    

    if (!sheet) {

      return { success: false, message: 'Configs tab not found' };

    }

    

    var data = sheet.getDataRange().getValues();

    Logger.log('Total rows in config sheet: ' + data.length);

    

    for (var i = 1; i < data.length; i++) {

      Logger.log('Checking row ' + i + ', ID: ' + data[i][0] + ' vs ' + configId);

      

      if (data[i][0] == configId || data[i][0].toString() === configId.toString()) {

        Logger.log('Found matching config!');

        Logger.log('Raw settings data (column I): ' + data[i][8]);

        

        var settingsJson = data[i][8];

        var parsedSettings;

        

        try {

          parsedSettings = JSON.parse(settingsJson);

          Logger.log('Successfully parsed settings: ' + JSON.stringify(parsedSettings));

        } catch (parseError) {

          Logger.log('ERROR parsing settings JSON: ' + parseError.toString());

          throw new Error('Invalid settings JSON in config: ' + parseError.toString());

        }

        

        var config = {

          id: data[i][0],

          name: data[i][1],

          spreadsheetId: data[i][2],

          spreadsheetName: data[i][3],

          settings: parsedSettings,

          createdAt: data[i][6],

          updatedAt: data[i][7]

        };

        

        Logger.log('Config loaded successfully: ' + config.name);

        

        return { success: true, config: config };

      }

    }

    

    Logger.log('Config not found with ID: ' + configId);

    return { success: false, message: 'Config not found' };

    

  } catch (e) {

    Logger.log('Error loading config: ' + e.toString());

    Logger.log('Stack: ' + e.stack);

    return { success: false, message: 'Error loading config: ' + e.toString() };

  }

}

/**

 * Store sync metadata in image description

 */

function storeSyncMetadata(image, fileNameOrMetadata, folderId, referenceMode) {

  try {

    var metadata = null;

    if (fileNameOrMetadata && typeof fileNameOrMetadata === 'object' && !Array.isArray(fileNameOrMetadata)) {

      metadata = fileNameOrMetadata;

      if (!metadata.lastSync && !metadata.syncedAt) {

        metadata.lastSync = new Date().toISOString();

      }

    } else {

      metadata = {

        syncedAt: new Date().toISOString(),

        fileName: fileNameOrMetadata,

        folderId: folderId,

        referenceMode: referenceMode,

        lastModified: null

      };

      if (folderId && fileNameOrMetadata) {

        var fileResult = getImageIdByName(folderId, fileNameOrMetadata);

        if (fileResult.success) {

          var file = DriveApp.getFileById(fileResult.id);

          metadata.lastModified = file.getLastUpdated().toISOString();

        }

      }

    }

    var existingDesc = image.getDescription() || '';
    var newDesc = upsertTaggedPayload(existingDesc, SYNC_METADATA_TAG, JSON.stringify(metadata));

    image.setDescription(newDesc);

  } catch (e) {

    Logger.log('Error storing metadata: ' + e.toString());

  }

}



/**

 * Extract sync metadata from image description

 */

function getSyncMetadata(imageOrDescription) {

  try {

    var payload = extractTaggedPayload(imageOrDescription, SYNC_METADATA_TAG);

    if (payload) {

      return JSON.parse(payload);

    }

  } catch (e) {

    Logger.log('Error parsing metadata: ' + e.toString());

  }

  return null;

}



/**

 * Check if image is up-to-date with Drive file

 */

function checkSyncStatus(metadata, folderId) {

  if (!metadata || !metadata.fileName || !metadata.folderId) {

    return 'unknown';

  }

  

  try {

    const fileResult = getImageIdByName(folderId || metadata.folderId, metadata.fileName);

    if (!fileResult.success) {

      return 'missing';

    }

    

    const file = DriveApp.getFileById(fileResult.id);

    const currentModified = file.getLastUpdated().toISOString();

    

    if (!metadata.lastModified) {

      return 'unknown';

    }

    

    if (currentModified > metadata.lastModified) {

      return 'outdated';

    }

    

    return 'synced';

    

  } catch (e) {

    Logger.log('Error checking sync status: ' + e.toString());

    return 'error';

  }

}



/**

 * Get all images from the presentation with their metadata

 */

function getAllImagesFromPresentation() {

  try {

    var presentation = SlidesApp.getActivePresentation();

    var slides = presentation.getSlides();

    var presentationId = presentation.getId();

    

    var allImages = [];

    

    slides.forEach(function(slide, index) {

      var images = slide.getImages();

      

      images.forEach(function(img) {

        var description = '';

        try {

          description = img.getDescription() || '';

        } catch (e) {

          description = '';

        }

        allImages.push({

          presentationId: presentationId,

          slideId: slide.getObjectId(),

          slideNumber: index + 1,

          objectId: img.getObjectId(),

          title: img.getTitle() || 'Untitled',

          description: description,

          width: img.getWidth(),

          height: img.getHeight()

        });

      });

    });

    

    return {

      success: true,

      images: allImages,

      slideCount: slides.length

    };

    

  } catch (e) {

    Logger.log('Error getting images: ' + e.toString());

    return {

      success: false,

      message: 'Error: ' + e.toString()

    };

  }

}



/**

 * Save image-to-config link mapping

 */

function saveImageConfigLink(objectId, configId) {

  try {

    var context = findImageContextByObjectId(objectId);

    if (!context) {

      throw new Error('Image not found with ID: ' + objectId);

    }

    storeConfigReference(context.image, configId);

    clearLegacyConfigReference(context.presentation.getId(), context.slideId, objectId);

    

    return { success: true };

  } catch (e) {

    Logger.log('Error saving link: ' + e.toString());

    return { success: false, message: e.toString() };

  }

}



/**

 * Remove image-to-config link mapping

 */

function removeImageConfigLink(objectId) {

  try {

    var context = findImageContextByObjectId(objectId);

    if (!context) {

      return { success: true };

    }

    storeConfigReference(context.image, null);

    clearLegacyConfigReference(context.presentation.getId(), context.slideId, objectId);

    

    return { success: true };

  } catch (e) {

    Logger.log('Error removing link: ' + e.toString());

    return { success: false, message: e.toString() };

  }

}



/**

 * Get all image-to-config link mappings for current presentation

 */

function getAllImageConfigLinks() {

  try {

    var links = {};

    var presentation = SlidesApp.getActivePresentation();

    var slides = presentation.getSlides();

    for (var i = 0; i < slides.length; i++) {

      var slide = slides[i];

      var images = slide.getImages();

      for (var j = 0; j < images.length; j++) {

        var image = images[j];

        var objectId = image.getObjectId();

        var configId = getConfigReference(image);

        if (!configId) {

          configId = getLegacyConfigReference(presentation.getId(), slide.getObjectId(), objectId);

          if (configId) {

            storeConfigReference(image, configId);

            clearLegacyConfigReference(presentation.getId(), slide.getObjectId(), objectId);

          }

        }

        if (configId) {

          links[objectId] = String(configId);

        }

      }

    }

    

    return links;

    

  } catch (e) {

    Logger.log('Error getting links: ' + e.toString());

    return {};

  }

}

/**

 * Sync selected slides - syncs ALL images with values (not just outdated)

 */

/**

 * Sync all images on currently selected slides

 * FIXED: Properly reads config links from your existing storage

 */

function syncSelectedSlides() {

  try {

    Logger.log('=== SYNCING SELECTED SLIDES ===');

    

    var presentation = SlidesApp.getActivePresentation();

    var selection = presentation.getSelection();

    

    if (!selection) {

      return {

        success: false,

        message: 'No selection found. Please select slides first.'

      };

    }

    

    var selectionType = selection.getSelectionType();

    Logger.log('Selection type: ' + selectionType);

    

    var selectedSlides = [];

    

    // Get selected slides based on selection type

    if (selectionType === SlidesApp.SelectionType.PAGE) {

      // User selected entire slides (filmstrip view)

      var pageRange = selection.getPageRange();

      selectedSlides = pageRange.getPages();

      Logger.log('Selected ' + selectedSlides.length + ' slides from page range');

      

    } else if (selectionType === SlidesApp.SelectionType.PAGE_ELEMENT) {

      // User selected elements on a slide - use current slide

      var currentPage = selection.getCurrentPage();

      if (currentPage && currentPage.getPageType() === SlidesApp.PageType.SLIDE) {

        selectedSlides = [currentPage];

        Logger.log('Using current slide (element selected)');

      }

      

    } else if (selectionType === SlidesApp.SelectionType.CURRENT_PAGE) {

      // Current page is selected

      var currentPage = selection.getCurrentPage();

      if (currentPage && currentPage.getPageType() === SlidesApp.PageType.SLIDE) {

        selectedSlides = [currentPage];

        Logger.log('Using current page');

      }

    }

    

    if (selectedSlides.length === 0) {

      return {

        success: false,

        message: 'No slides selected. Please select one or more slides in the filmstrip.'

      };

    }

    

    Logger.log('Processing ' + selectedSlides.length + ' slide(s)');

    

    // Collect all images with config links on selected slides

    var imagesToSync = [];

    var allSlides = presentation.getSlides();

    

    for (var i = 0; i < selectedSlides.length; i++) {

      var slide = selectedSlides[i];

      var slideId = slide.getObjectId();

      

      // Find slide index

      var slideIndex = -1;

      for (var j = 0; j < allSlides.length; j++) {

        if (allSlides[j].getObjectId() === slideId) {

          slideIndex = j;

          break;

        }

      }

      

      Logger.log('\nProcessing slide ' + (slideIndex + 1));

      

      var images = slide.getImages();

      Logger.log('Found ' + images.length + ' image(s) on slide');

      

      for (var k = 0; k < images.length; k++) {

        var image = images[k];

        var objectId = image.getObjectId();

        

        var configId = getConfigReference(image);

        

        if (configId) {

          imagesToSync.push({

            objectId: objectId,

            configId: configId,

            slideIndex: slideIndex,

            slideNumber: slideIndex + 1

          });

          Logger.log('  ✓ Found linked image: ' + configId);

        } else {

          Logger.log('  - Image has no config link (skipped)');

        }

      }

    }

    

    Logger.log('\nFound ' + imagesToSync.length + ' image(s) with config links');

    

    if (imagesToSync.length === 0) {

      return {

        success: true,

        message: 'No images with config links found on selected slides',

        synced: 0,

        failed: 0

      };

    }

    

    // Sync each image

    var synced = 0;

    var failed = 0;

    var errors = [];

    

    for (var m = 0; m < imagesToSync.length; m++) {

      var item = imagesToSync[m];

      

      Logger.log('\n--- Syncing image ' + (m + 1) + '/' + imagesToSync.length + ' ---');

      Logger.log('Slide: ' + item.slideNumber);

      Logger.log('Config ID: ' + item.configId);

      

      try {

        var result = syncImageFromConfig(item.objectId, item.configId);

        

        if (result.success) {

          synced++;

          Logger.log('✓ Synced successfully');

        } else {

          failed++;

          errors.push('Slide ' + item.slideNumber + ': ' + result.message);

          Logger.log('✗ Sync failed: ' + result.message);

        }

      } catch (e) {

        failed++;

        errors.push('Slide ' + item.slideNumber + ': ' + e.toString());

        Logger.log('✗ Exception: ' + e.toString());

      }

      

      // Small delay

      if (m < imagesToSync.length - 1) {

        Utilities.sleep(500);

      }

    }

    

    Logger.log('\n=== SYNC COMPLETE ===');

    Logger.log('Synced: ' + synced);

    Logger.log('Failed: ' + failed);

    

    var message = 'Synced ' + synced + ' image(s) on ' + selectedSlides.length + ' slide(s)';

    if (failed > 0) {

      message += '\n\nFailed: ' + failed + ' image(s)';

      if (errors.length > 0) {

        message += '\n' + errors.slice(0, 3).join('\n');

      }

    }

    

    return {

      success: true,

      message: message,

      synced: synced,

      failed: failed,

      total: imagesToSync.length,

      selectedSlides: selectedSlides.length

    };

    

  } catch (e) {

    Logger.log('\n❌ SYNC SELECTED FAILED');

    Logger.log('Error: ' + e.toString());

    return {

      success: false,

      message: 'Sync failed: ' + e.toString(),

      synced: 0,

      failed: 0

    };

  }

}



/**

 * Replace a specific image - supports both ID and filename

 */

function replaceImage(presentationId, slideId, imageObjectId, newImageReference, referenceFolderId, referenceMode) {

  try {

    const presentation = SlidesApp.openById(presentationId);

    const slide = presentation.getSlides().find(s => s.getObjectId() === slideId);

    

    if (!slide) {

      throw new Error("Slide not found");

    }

    

    const images = slide.getImages();

    const targetImage = images.find(img => img.getObjectId() === imageObjectId);

    

    if (!targetImage) {

      throw new Error("Image not found on slide");

    }

    

    let fileId = newImageReference;

    let fileName = newImageReference;

    

    // If reference mode is 'name', look up the file by name

    if (referenceMode === 'name' && referenceFolderId) {

      const result = getImageIdByName(referenceFolderId, newImageReference);

      if (!result.success) {

        throw new Error(result.error);

      }

      fileId = result.id;

      fileName = newImageReference;

    }

    

    // Get the position and size of the old image

    const left = targetImage.getLeft();

    const top = targetImage.getTop();

    const width = targetImage.getWidth();

    const height = targetImage.getHeight();

    const title = targetImage.getTitle();

    

    // Get the new image from Drive

    const driveFile = DriveApp.getFileById(fileId);

    const blob = driveFile.getBlob();

    

    // Remove the old image

    targetImage.remove();

    

    // Insert the new image at the same position

    const newImage = slide.insertImage(blob, left, top, width, height);

    

    // Restore metadata

    if (title) newImage.setTitle(title);

    var existingConfigId = getConfigReference(targetImage);

    

    // Store sync metadata

    if (existingConfigId) {

      storeConfigReference(newImage, existingConfigId);

    }

    if (referenceMode === 'name') {

      storeSyncMetadata(newImage, fileName, referenceFolderId, referenceMode);

    }

    

    return {

      success: true,

      message: "Image replaced successfully",

      newObjectId: newImage.getObjectId()

    };

    

  } catch (error) {

    return {

      success: false,

      error: error.toString()

    };

  }

}







/**

 * Get list of available configs from GSheets script

 * This requires the GSheets script functions to be accessible

 */

  function getAvailableConfigs() {

    try {

      var props = PropertiesService.getUserProperties();

      var configSheetId = props.getProperty('configSheetId');

      

      Logger.log('Config Sheet ID: ' + configSheetId);

      

      if (!configSheetId) {

        Logger.log('No config sheet linked');

        return [];

      }

      

      var configSS = SpreadsheetApp.openById(configSheetId);

      var sheet = configSS.getSheetByName(CONFIGS_SHEET_NAME);

      

      Logger.log('Sheet found: ' + (sheet ? 'Yes' : 'No'));

      

      if (!sheet) {

        Logger.log('Configs sheet not found. Available sheets: ' + configSS.getSheets().map(function(s) { return s.getName(); }).join(', '));

        return [];

      }

      

      var lastRow = sheet.getLastRow();

      Logger.log('Last row: ' + lastRow);

      

      if (lastRow < 2) {

        Logger.log('No data rows found');

        return [];

      }

      

      // Get all data from row 2 onwards

      var numRows = lastRow - 1;

      var data = sheet.getRange(2, 1, numRows, 9).getValues();

      

      Logger.log('Data rows retrieved: ' + data.length);

      

      var configs = [];

      

      for (var i = 0; i < data.length; i++) {

        Logger.log('Processing row ' + (i + 2) + ': ID=' + data[i][0]);

        

        if (data[i][0]) { // Check if ID exists

          var settings = {};

          

          // Try to parse settings JSON

          try {

            if (data[i][8] && data[i][8] !== '') {

              settings = JSON.parse(data[i][8]);

            }

          } catch (e) {

            Logger.log('Error parsing settings for row ' + (i + 2) + ': ' + e);

            settings = {

              convertToPng: false,

              sheetTabs: data[i][4] || '',

              cellRanges: data[i][5] || ''

            };

          }

          

          var config = {

            id: data[i][0].toString(),

            name: data[i][1] || 'Unnamed Config',

            spreadsheetId: data[i][2] || '',

            spreadsheetName: data[i][3] || 'Unknown',

            sheetTab: data[i][4] || '',

            range: data[i][5] || '',

            createdAt: data[i][6] || '',

            updatedAt: data[i][7] || '',

            settings: settings

          };

          

          configs.push(config);

          Logger.log('Added config: ' + config.name + ' (ID: ' + config.id + ')');

        }

      }

      

      Logger.log('Total configs loaded: ' + configs.length);

      return configs;

      

    } catch (e) {

      Logger.log('ERROR in getAvailableConfigs: ' + e.toString());

      Logger.log('Stack trace: ' + e.stack);

      return [];

    }

  }

/**

 * Parses cell reference like "D10" into {col: 3, row: 10}

 * Same as Google Sheets script

 */

function parseCellReference(cell) {

  const match = cell.toUpperCase().match(/^([A-Z]+)(\d+)$/);

  if (!match) {

    throw new Error("Invalid cell reference: " + cell);

  }

  return {

    col: columnToNumber(match[1]),

    row: parseInt(match[2])

  };

}






/**

 * Converts A1 notation to column number (A=1, B=2, Z=26, AA=27, etc.)

 * Returns 0-based index for API

 */

function columnToNumber(column) {

  let num = 0;

  for (let i = 0; i < column.length; i++) {

    num = num * 26 + (column.charCodeAt(i) - 64);

  }

  return num - 1; // 0-based for API

}



/**

 * Generate PDF snapshot blob from Google Sheets range

 */

function generateSnapshotBlob(spreadsheetId, sheetId, pdfRange, fileName) {

  try {

    // Build export URL with correct range parameters

    // NOTE: Google Sheets export API uses 0-based indexing for rows

    var url = "https://docs.google.com/spreadsheets/d/" + spreadsheetId + "/export?" +

      "format=pdf&" +

      "size=letter&" +

      "portrait=false&" +

      "fitw=true&" +

      "gridlines=false&" +

      "top_margin=0.00&" +

      "bottom_margin=0.00&" +

      "left_margin=0.00&" +

      "right_margin=0.00&" +

      "gid=" + sheetId +

      "&c1=" + (pdfRange.c1) +

      "&r1=" + (pdfRange.r1-1) +  // Convert to 0-based

      "&c2=" + (pdfRange.c2+1) +

      "&r2=" + (pdfRange.r2);   // Convert to 0-based

    

    Logger.log('Export URL: ' + url);

    

    var token = ScriptApp.getOAuthToken();

    var response = UrlFetchApp.fetch(url, {

      headers: { 'Authorization': 'Bearer ' + token },

      muteHttpExceptions: true

    });

    

    var responseCode = response.getResponseCode();

    Logger.log('Export response code: ' + responseCode);

    

    if (responseCode === 200) {

      var pdfName = (fileName || 'export') + '.pdf';

      var blob = response.getBlob().setName(pdfName);

      Logger.log('✓ PDF blob created: ' + blob.getBytes().length + ' bytes');

      return blob;

    } else {

      var errorMsg = response.getContentText();

      Logger.log('❌ Export failed: ' + errorMsg);

      throw new Error('Failed to export PDF (HTTP ' + responseCode + '): ' + errorMsg);

    }

  } catch (e) {

    Logger.log('❌ Error in generateSnapshotBlob: ' + e.toString());

    throw e;

  }

}



/**

 * Store config reference in image metadata

 */

function storeConfigReference(imageOrPresentationId, slideId, objectId, configId) {

  try {

    var image = imageOrPresentationId;
    var resolvedObjectId = objectId;

    if (!image || typeof image.getDescription !== 'function') {

      if (arguments.length >= 4) {

        resolvedObjectId = objectId;
        configId = arguments[3];

      } else {

        resolvedObjectId = slideId;
        configId = objectId;

      }

      var context = findImageContextByObjectId(resolvedObjectId);

      image = context ? context.image : null;

    }

    if (!image) {

      throw new Error('Image not found while storing config reference');

    }

    var existingDesc = image.getDescription() || '';

    if (!configId) {

      image.setDescription(stripTaggedPayload(existingDesc, CONFIG_LINK_TAG));

      return { success: true, message: 'Config link removed successfully' };

    }

    var payload = JSON.stringify({
      configId: String(configId),
      linkedAt: new Date().toISOString()
    });

    image.setDescription(upsertTaggedPayload(existingDesc, CONFIG_LINK_TAG, payload));

    return { success: true, message: 'Config linked successfully' };

  } catch (e) {

    return { success: false, message: 'Error: ' + e.toString() };

  }

}



/**

 * Get config reference from image metadata

 */

function getConfigReference(imageOrPresentationId, slideId, objectId) {

  try {

    var image = imageOrPresentationId;
    var resolvedSlideId = slideId;
    var resolvedObjectId = objectId;
    var presentationId = null;

    if (!image || typeof image.getDescription !== 'function') {

      presentationId = imageOrPresentationId;
      var context = findImageContextByObjectId(objectId);

      if (!context) {

        return null;

      }

      image = context.image;
      resolvedSlideId = context.slideId;
      resolvedObjectId = context.objectId;
      presentationId = presentationId || context.presentation.getId();

    }

    var taggedConfigId = getConfigIdFromTagPayload(extractTaggedPayload(image, CONFIG_LINK_TAG));

    if (taggedConfigId) {

      return taggedConfigId;

    }

    var syncMetadata = getSyncMetadata(image);

    if (syncMetadata && syncMetadata.configId) {

      storeConfigReference(image, syncMetadata.configId);

      return String(syncMetadata.configId);

    }

    if (!presentationId) {

      presentationId = SlidesApp.getActivePresentation().getId();

    }

    var legacyConfigId = getLegacyConfigReference(presentationId, resolvedSlideId, resolvedObjectId);

    if (legacyConfigId) {

      storeConfigReference(image, legacyConfigId);
      clearLegacyConfigReference(presentationId, resolvedSlideId, resolvedObjectId);
      return String(legacyConfigId);

    }

    return null;

  } catch (e) {

    Logger.log('Error getting config reference: ' + e.toString());

    return null;

  }

}



/**
 * Get Google Identity Token for Cloud Run via Service Account Impersonation
 * This uses the IAM Credentials API to generate an OIDC token keylessly.
 */
function getCloudRunIdentityToken(targetAudience) {
  throw new Error('Cloud Run identity token flow is not configured in this public repo. Configure a deployment-specific service account before using this helper.');

  try {
    Logger.log("Starting Service Account Impersonation flow...");
    
    // 1. Get the Active User's OAuth Token
    const oauthToken = ScriptApp.getOAuthToken();
    
    // 2. Prepare the IAM Credentials API request
    const iamUrl = `https://iamcredentials.googleapis.com/v1/projects/-/serviceAccounts/${SERVICE_ACCOUNT_EMAIL}:generateIdToken`;
    
    const payload = {
      audience: targetAudience, // Cloud Run requires the audience to match its URL
      includeEmail: true
    };
    
    const options = {
      method: "post",
      contentType: "application/json",
      headers: {
        "Authorization": `Bearer ${oauthToken}`
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    
    // 3. Request the OIDC token
    const response = UrlFetchApp.fetch(iamUrl, options);
    const responseCode = response.getResponseCode();
    
    if (responseCode === 200) {
      const result = JSON.parse(response.getContentText());
      Logger.log("✅ Successfully generated OIDC token via impersonation.");
      return result.token; // The IAM API returns the OIDC token inside the "token" property
    } else {
      Logger.log("❌ IAM API Error: " + response.getContentText());
      throw new Error("Failed to generate token. HTTP " + responseCode + ": " + response.getContentText());
    }
    
  } catch (e) {
    Logger.log("❌ Error in getCloudRunIdentityToken: " + e.toString());
    throw new Error("Failed to get authentication token: " + e.toString());
  }
}


/**

 * Resolves named range to actual cell coordinates (DEBUG VERSION)

 * Returns null if not a named range, or the A1 notation if it is

 */

function resolveNamedRange(ss, rangeName) {

  console.log('🔍 resolveNamedRange called with: "' + rangeName + '"');

  

  try {

    // First check if it looks like a cell range (contains ':')

    if (rangeName.indexOf(':') > -1) {

      console.log('  → Contains ":", treating as regular cell range');

      return null;

    }

    

    console.log('  → Attempting to get named range...');

    const namedRange = ss.getRangeByName(rangeName);

    

    if (!namedRange) {

      console.log('  → Not a named range, returning null');

      return null; // Not a named range, treat as regular cell reference

    }

    

    // Return the A1 notation (e.g., "D10:AD62")

    const a1Notation = namedRange.getA1Notation();

    console.log('  ✅ Named range found! Resolved to: ' + a1Notation);

    return a1Notation;

    

  } catch (e) {

    console.log('  ⚠️ Error resolving named range "' + rangeName + '": ' + e.toString());

    return null;

  }

}


function convertPdfToPngCloudConvert(pdfBlob, configName, apiKey, pngSettings) {
  var baseUrl = 'https://sync.api.cloudconvert.com/v2/jobs';
  var safeName = String(configName || 'sheet_export').replace(/[^a-zA-Z0-9_-]/g, '_');
  var fileName = safeName + '.pdf';
  var outputName = safeName + '.png';
  var pixelDensity = parseInt(pngSettings.dpi, 10) || 300;

  var payload = {
    tasks: {
      'import-pdf': {
        operation: 'import/base64',
        file: Utilities.base64Encode(pdfBlob.getBytes()),
        filename: fileName
      },
      'convert-pdf': {
        operation: 'convert',
        input: 'import-pdf',
        input_format: 'pdf',
        output_format: 'png',
        filename: outputName,
        pixel_density: pixelDensity,
        pages: '1'
      },
      'export-png': {
        operation: 'export/url',
        input: 'convert-pdf'
      }
    }
  };

  var response = UrlFetchApp.fetch(baseUrl, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    headers: {
      Authorization: 'Bearer ' + apiKey
    },
    muteHttpExceptions: true
  });

  if (response.getResponseCode() >= 300) {
    throw new Error('CloudConvert job failed (' + response.getResponseCode() + '): ' + response.getContentText());
  }

  var result = JSON.parse(response.getContentText());
  var tasks = result && result.data && result.data.tasks ? result.data.tasks : [];
  var exportTask = null;
  for (var i = 0; i < tasks.length; i++) {
    if (tasks[i].name === 'export-png') {
      exportTask = tasks[i];
      break;
    }
  }

  if (!exportTask || exportTask.status !== 'finished') {
    throw new Error('CloudConvert export task did not finish successfully.');
  }

  var files = exportTask.result && exportTask.result.files ? exportTask.result.files : [];
  if (!files.length || !files[0].url) {
    throw new Error('CloudConvert did not return a downloadable PNG URL.');
  }

  var downloadResponse = UrlFetchApp.fetch(files[0].url, {
    muteHttpExceptions: true
  });

  if (downloadResponse.getResponseCode() >= 300) {
    throw new Error('CloudConvert download failed (' + downloadResponse.getResponseCode() + '): ' + downloadResponse.getContentText());
  }

  return downloadResponse.getBlob().setName(outputName).setContentType('image/png');
}


// ============================================
// GCS HELPER FUNCTIONS (OAuth token – no JSON key)
// ============================================

/**
 * Upload a PDF blob to the GCS input bucket using the user's OAuth token.
 * Conversion parameters are stored as GCS object metadata so Cloud Run
 * can read them from the Pub/Sub event payload.
 *
 * @param {string} bucket         GCS bucket name
 * @param {string} objectName     destination object name (e.g. "config_123_1234567890.pdf")
 * @param {Blob}   blob           PDF blob to upload
 * @param {Object} conversionMeta key/value pairs – all values must be strings
 */
function uploadPdfToGcs(bucket, objectName, blob, conversionMeta) {
  var boundary      = 'gcs_boundary_' + new Date().getTime();
  var metaJson      = JSON.stringify({ name: objectName, metadata: conversionMeta });
  var delimiter     = '\r\n--' + boundary + '\r\n';
  var closeDelimiter = '\r\n--' + boundary + '--';

  var headerPart = delimiter
    + 'Content-Type: application/json; charset=UTF-8\r\n\r\n'
    + metaJson
    + delimiter
    + 'Content-Type: application/pdf\r\n\r\n';

  var bodyBytes = Utilities.newBlob(headerPart).getBytes()
    .concat(blob.getBytes())
    .concat(Utilities.newBlob(closeDelimiter).getBytes());

  var url = 'https://storage.googleapis.com/upload/storage/v1/b/'
    + encodeURIComponent(bucket)
    + '/o?uploadType=multipart';

  var response = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'multipart/related; boundary=' + boundary,
    payload: bodyBytes,
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  });

  if (response.getResponseCode() >= 300) {
    throw new Error(
      'GCS upload failed (' + response.getResponseCode() + '): '
      + response.getContentText()
    );
  }
  Logger.log('✓ PDF uploaded to gs://' + bucket + '/' + objectName);
}


/**
 * Poll the GCS output bucket every 3 seconds until the PNG appears or timeout.
 *
 * @param {string} bucket       GCS bucket to poll
 * @param {string} objectName   expected PNG object name
 * @param {number} timeoutSecs  max seconds to wait (default 60)
 */
function pollForPngInGcs(bucket, objectName, timeoutSecs) {
  var maxMs      = (timeoutSecs || 60) * 1000;
  var intervalMs = 3000;
  var elapsed    = 0;

  var url = 'https://storage.googleapis.com/storage/v1/b/'
    + encodeURIComponent(bucket)
    + '/o/'
    + encodeURIComponent(objectName);

  while (elapsed < maxMs) {
    Utilities.sleep(intervalMs);
    elapsed += intervalMs;

    var response = UrlFetchApp.fetch(url, {
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    });

    if (response.getResponseCode() === 200) {
      Logger.log('✓ PNG ready after ' + (elapsed / 1000) + 's');
      return;
    }
    Logger.log('  PNG not ready yet (' + (elapsed / 1000) + 's elapsed)...');
  }

  throw new Error('Timeout: PNG not found in GCS after ' + (timeoutSecs || 60) + 's');
}


/**
 * Download a PNG blob from the GCS output bucket using the user's OAuth token.
 *
 * @param  {string} bucket      GCS bucket name
 * @param  {string} objectName  object to download
 * @return {Blob}               PNG blob ready for use in Slides
 */
function downloadPngFromGcs(bucket, objectName) {
  var url = 'https://storage.googleapis.com/storage/v1/b/'
    + encodeURIComponent(bucket)
    + '/o/'
    + encodeURIComponent(objectName)
    + '?alt=media';

  var response = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  });

  if (response.getResponseCode() >= 300) {
    throw new Error(
      'GCS download failed (' + response.getResponseCode() + '): '
      + response.getContentText()
    );
  }
  Logger.log('✓ PNG downloaded from gs://' + bucket + '/' + objectName);
  return response.getBlob().setName(objectName);
}


/**
 * Full GCS conversion pipeline:
 *   1. Upload PDF to GCS input bucket with conversion params as object metadata
 *   2. Poll GCS output bucket until Cloud Run writes the PNG (up to 60s)
 *   3. Download and return the PNG blob
 *
 * To use this path, set pngConversionMethod = 'gcs' in a config's settings JSON.
 *
 * @param  {Blob}   pdfBlob     PDF blob from generateSnapshotBlob()
 * @param  {string} configName  used to build unique GCS object names
 * @param  {Object} pngSettings same settings object used by other converters
 * @return {Blob}               converted PNG blob
 */
function convertPdfToPngViaGcs(pdfBlob, configName, pngSettings, inputBucket, outputBucket) {
  var timestamp  = new Date().getTime();
  // Random 8-char suffix ensures two parallel calls for the same config never
  // share a GCS filename, even if they start within the same millisecond.
  var uniqueId   = Math.random().toString(36).substr(2, 8);
  var safeName   = configName.replace(/[^a-zA-Z0-9_-]/g, '_');
  var inputName  = safeName + '_' + timestamp + '_' + uniqueId + '.pdf';
  var outputName = safeName + '_' + timestamp + '_' + uniqueId + '.png';

  // GCS object metadata values must all be strings
  var meta = {
    dpi:         String(pngSettings.dpi        || 300),
    quality:     String(pngSettings.quality    || 95),
    background:  String(pngSettings.background || 'white'),
    auto_crop:   String(pngSettings.autoCrop   || false),
    crop_left:   String(pngSettings.cropLeft   || 0),
    crop_top:    String(pngSettings.cropTop    || 0),
    crop_right:  String(pngSettings.cropRight  || 0),
    crop_bottom: String(pngSettings.cropBottom || 0)
  };

  Logger.log('GCS pipeline: uploading ' + inputName + ' to gs://' + inputBucket);
  uploadPdfToGcs(inputBucket, inputName, pdfBlob, meta);

  // 120s timeout (instead of 60s) to handle parallel batch syncs where Cloud Run
  // may process multiple requests sequentially or need a cold-start (~20-30s).
  // With 3 images in parallel: cold-start (~25s) + 3×processing (~15s each) = ~70s,
  // safely within the 120s window.
  Logger.log('GCS pipeline: waiting for gs://' + outputBucket + '/' + outputName);
  pollForPngInGcs(outputBucket, outputName, 120);

  Logger.log('GCS pipeline: downloading PNG');
  var pngBlob = downloadPngFromGcs(outputBucket, outputName);

  // Clean up both temporary GCS files immediately after successful download.
  // A 1-day lifecycle rule on the buckets acts as a safety net for any files
  // left behind if this cleanup fails (e.g. AppScript timeout).
  Logger.log('GCS pipeline: cleaning up temporary files');
  deleteGcsObject(inputBucket, inputName);
  deleteGcsObject(outputBucket, outputName);

  return pngBlob;
}


/**
 * Delete a single object from a GCS bucket using the user's OAuth token.
 * Failures are logged but not thrown so cleanup never blocks the sync result.
 *
 * @param {string} bucket      GCS bucket name
 * @param {string} objectName  object to delete
 */
function deleteGcsObject(bucket, objectName) {
  try {
    var url = 'https://storage.googleapis.com/storage/v1/b/'
      + encodeURIComponent(bucket)
      + '/o/'
      + encodeURIComponent(objectName);

    var response = UrlFetchApp.fetch(url, {
      method: 'delete',
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    });

    var code = response.getResponseCode();
    if (code === 204 || code === 404) {
      // 204 = deleted successfully, 404 = already gone — both are fine
      Logger.log('\u2713 Deleted gs://' + bucket + '/' + objectName);
    } else {
      Logger.log('\u26a0\ufe0f GCS delete returned ' + code + ' for gs://' + bucket + '/' + objectName);
    }
  } catch (e) {
    Logger.log('\u26a0\ufe0f deleteGcsObject failed (non-fatal): ' + e.toString());
  }
}
