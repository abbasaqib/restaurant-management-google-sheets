function test(){
  Logger.log(callGeminiAPI('What is abc'))
}
// ==============================================
// MAIN CONFIGURATION
// ==============================================
const CONFIG = {
  COMPANY_NAME: 'Zoma Culinary',
  CHEF_NAME: 'Chef Martha Garcia',
  EMAIL: 'info@zomaculinary.com',
  PHONE: '(123) 456-7890',
  WEBSITE: 'www.zomaculinary.com',
  TAX_RATE: 0.0895, // 8.95%
  FOLDER_ID: '1kTHyN9ZLTHPBFbMa66Kk1a-fyKcC0lAR', // Replace with your folder ID
  LOGO_FILE_ID: '1u34EHuMY8q-_ZOtOu43CfENIrdhxPuaR', // Logo file ID
  SIGNATURE_IMAGE_ID: 'YOUR_SIGNATURE_IMAGE_ID', // Replace with your signature image ID
  SPREADSHEET_ID: '1TrB-ACjCGKSee8QWwZ68EHa96WDFfI7q8vjG6gNCjsE', // Replace with your Google Sheet ID
  GEMINI_API_KEY: 'AIzaSyCDqwZwZK0zA24-YDypDQ8qbhvl6kTQRYg', // Add your Gemini API key here
  EDITORS: [
    'martha@zomaculinary.com',
    'jasonhurley@zomaculinary.com',
    'auraofblue@gmail.com'
  ],
  // Contract-specific config
  COMPANY_ADDRESS: 'Zoma Culinary, Austin, TX',
  CHEF_TITLE: 'Executive Chef & Owner',
  SIGNATURE_NAME: 'Martha Garcia',
  DEFAULT_DEPOSIT_PERCENT: 50,
  DEFAULT_CANCELLATION_DAYS: 5,
  MENU_TEXT:""
};

// ==============================================
// MAIN FORM HANDLER
// ==============================================
function saveForm(formData) {
  try {
    console.log('Form submission received');
    console.log('Create quote document:', formData.createQuote);
    console.log('Create contract document:', formData.contractData?.createContract);
    console.log('Create reheating instructions:', formData.createReheatingInstructions);
    console.log('Create shopping list:', formData.createShoppingList);
    
    // Validate form data based on document type
    const validation = validateFormData(formData);
    if (!validation.isValid) {
      throw new Error(validation.message);
    }
    
    // Generate a unique submission ID if not present
    if (!formData.submissionId) {
      formData.submissionId = Utilities.getUuid();
    }
    
    // Save to Google Sheet
    const sheetRow = saveToSpreadsheet(formData);
    
    // Create subfolder for this submission
    const submissionFolder = createSubmissionFolder(formData);
    
    // Create documents based on checkbox selection
    let quoteDoc = null;
    let labelsDocs = []; // Changed to array to hold multiple label documents
    let scheduleDoc = null;
    let contractDoc = null;
    let quotePdf = null;
    let labelsPdfs = []; // Changed to array to hold multiple PDFs
    let schedulePdf = null;
    let contractPdf = null;
    
    if (formData.createQuote) {
      // Create quote, labels, and schedule
      quoteDoc = createQuoteDocument(formData, submissionFolder);
      labelsDocs = createMenuLabelsDocument(formData, submissionFolder); // Returns array
      scheduleDoc = createMenuScheduleDocument(formData, submissionFolder);
      
      // Generate PDF versions
      quotePdf = createPDF(quoteDoc);
      labelsPdfs = labelsDocs.map(doc => createPDF(doc)).filter(pdf => pdf); // Generate PDFs for all label docs
      schedulePdf = createPDF(scheduleDoc);
    } else {
      // Create only menu schedule and labels
      labelsDocs = createMenuLabelsDocument(formData, submissionFolder); // Returns array
      scheduleDoc = createMenuScheduleDocument(formData, submissionFolder);
      
      // Generate PDF versions
      labelsPdfs = labelsDocs.map(doc => createPDF(doc)).filter(pdf => pdf);
      schedulePdf = createPDF(scheduleDoc);
    }
    
    // Create contract if requested
    if (formData.contractData?.createContract) {
      contractDoc = createContractDocument(formData, submissionFolder);
      contractPdf = createPDF(contractDoc);
    }
    
    // Update spreadsheet with document URLs (store first label doc URL for compatibility)
    updateDocumentUrls(
      formData.submissionId, 
      quoteDoc ? quoteDoc.getUrl() : '',
      labelsDocs.length > 0 ? labelsDocs[0].getUrl() : '', // Store first label doc URL
      scheduleDoc.getUrl(), 
      contractDoc ? contractDoc.getUrl() : ''
    );
    
    // Store all label document IDs for later use in email
    if (labelsDocs.length > 0) {
      const docIds = labelsDocs.map(doc => doc.getId());
      PropertiesService.getScriptProperties().setProperty('labelsDocs_' + formData.submissionId, JSON.stringify(docIds));
    }
    
    // Create additional documents if requested
    let reheatingInstructionsSent = false;
    let shoppingListSent = false;
    
    if (formData.createReheatingInstructions) {
      reheatingInstructionsSent = sendReheatingInstructionsEmail(formData, submissionFolder);
    }
    
    if (formData.createShoppingList) {
      shoppingListSent = sendShoppingListEmail(formData, submissionFolder);
    }
    
    // Send confirmation emails
    sendConfirmationEmail(
      formData, 
      quoteDoc, 
      labelsDocs, // Pass array of label docs
      scheduleDoc, 
      contractDoc,
      quotePdf, 
      labelsPdfs, // Pass array of label PDFs
      schedulePdf, 
      contractPdf,
      formData.createQuote,
      formData.contractData?.createContract || false
    );
    
    // Return success response
    return {
      success: true,
      message: getSuccessMessage(formData),
      submissionId: formData.submissionId,
      sheetRow: sheetRow,
      quoteUrl: quoteDoc ? quoteDoc.getUrl() : '',
      quoteName: quoteDoc ? quoteDoc.getName() : '',
      labelsUrl: labelsDocs.length > 0 ? labelsDocs[0].getUrl() : '',
      labelsName: labelsDocs.length > 0 ? labelsDocs[0].getName() : '',
      totalLabelDocuments: labelsDocs.length,
      scheduleUrl: scheduleDoc.getUrl(),
      scheduleName: scheduleDoc.getName(),
      contractUrl: contractDoc ? contractDoc.getUrl() : '',
      contractName: contractDoc ? contractDoc.getName() : '',
      reheatingInstructionsSent: reheatingInstructionsSent,
      shoppingListSent: shoppingListSent
    };
    
  } catch (error) {
    console.error('Error in saveForm:', error);
    return {
      success: false,
      message: error.toString()
    };
  }
}

function getSuccessMessage(formData) {
  let message = '';
  if (formData.createQuote && formData.contractData?.createContract) {
    message = 'Quote and contract documents created successfully!';
  } else if (formData.createQuote) {
    message = 'Quote document created successfully!';
  } else if (formData.contractData?.createContract) {
    message = 'Contract document created successfully!';
  } else {
    message = 'Menu schedule and labels created successfully!';
  }
  return message;
}

// ==============================================
// SPREADSHEET FUNCTIONS
// ==============================================
function saveToSpreadsheet(formData) {
  try {
    const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    let sheet = spreadsheet.getSheetByName('Responses');
    
    // Create sheet if it doesn't exist
    if (!sheet) {
      sheet = spreadsheet.insertSheet('Responses');
      // Create headers
      const headers = [
        'Submission ID',
        'Timestamp',
        'Client Name',
        'Email',
        'Employee Emails',
        'Client Phone',
        'Service Type',
        'Date Requested',
        'Event Address',
        'Event Type',
        'Adults',
        'Kids',
        'Total Guests',
        'Event Title',
        'Event Name',
        'Adult Price',
        'Kid Price',
        'Menu Design',
        'Shopping',
        'Delivery',
        'Container Removal',
        'Service Style',
        'Service Start Time',
        'Service End Time',
        'Setup Time',
        'Kitchen Reset',
        'Menu Final Date',
        'Min Guest Count',
        'Min Food Spend',
        'Chef Labor Hours',
        'Chef Hourly Rate',
        'Num Servers',
        'Server Hours',
        'Server Hourly Rate',
        'Assistant',
        'Assistant Hours',
        'Assistant Rate',
        'Overtime Rate',
        'Delivery Fee',
        'Travel Fee',
        'Equipment Fee',
        'Rentals',
        'Rental Cost',
        'Disposables Cost',
        'Chafers Provided',
        'Linen Provided',
        'Water Pitchers Provided',
        'Glassware Provided',
        'Flowers Provided',
        'Other Rental Notes',
        'Deposit Percent',
        'Deposit Due Date',
        'Balance Due Date',
        'Payment Methods',
        'Included Services',
        'Other Services Notes',
        'Dishes JSON',
        'Contract Dishes JSON',
        'Create Quote',
        'Create Contract',
        'Create Reheating Instructions',
        'Create Shopping List',
        'Documents Folder URL',
        'Quote Document URL',
        'Labels Document URL',
        'Schedule Document URL',
        'Contract Document URL'
      ];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    }
    
    // Find existing row by submission ID
    let row = 2;
    const submissionId = formData.submissionId;
    const rcount = sheet.getLastRow() - 1;
    let ids = [];
    if (rcount > 0) {
      ids = sheet.getRange(2, 1, rcount, 1).getValues();
    }
    
    for (let i = 0; i < ids.length; i++) {
      if (ids[i][0] === submissionId) {
        row = i + 2;
        break;
      }
    }
    
    // If new submission, add to the end
    if (row === 2 && ids.length > 0 && ids[0][0] !== submissionId) {
      row = sheet.getLastRow() + 1;
    }
    
    // Serialize dishes arrays to JSON
    const dishesJson = JSON.stringify(formData.dishes || []);
    
    // For contract dishes, use the same dishes data if contract is being created
    let contractDishesJson = '[]';
    if (formData.contractData?.createContract) {
      contractDishesJson = JSON.stringify(formData.dishes || []);
    } else {
      contractDishesJson = JSON.stringify(formData.contractData?.contractDishes || []);
    }
    
    // Calculate total guests
    const adults = parseInt(formData.adults) || 0;
    const kids = parseInt(formData.kids) || 0;
    const totalGuests = adults + kids;
    
    // Collect payment methods and included services
    const paymentMethods = formData.contractData?.paymentMethods ? formData.contractData.paymentMethods.join(', ') : '';
    const includedServices = formData.contractData?.includedServices ? formData.contractData.includedServices.join(', ') : '';
    
    // Prepare data row
    const rowData = [
      submissionId,
      new Date(),
      formData.clientName,
      formData.email,
      formData.employeeEmails || '',
      formData.contractData?.clientPhone || '',
      formData.serviceType,
      formData.dateRequested,
      formData.contractData?.eventAddress || '',
      formData.contractData?.eventType || '',
      formData.adults,
      formData.kids || 0,
      totalGuests,
      formData.eventTitle || '',
      formData.eventName || '',
      formData.adultPrice,
      formData.kidPrice || 0,
      formData.options?.menuDesign || false,
      formData.options?.shopping || false,
      formData.options?.delivery || false,
      formData.options?.containerRemoval || false,
      formData.contractData?.serviceStyle || '',
      formData.contractData?.serviceStartTime || '',
      formData.contractData?.serviceEndTime || '',
      formData.contractData?.setupTime || '',
      formData.contractData?.kitchenReset || 'No',
      formData.contractData?.menuFinalDate || '',
      formData.contractData?.minGuestCount || '0',
      formData.contractData?.minFoodSpend || '0',
      formData.contractData?.chefLaborHours || '0',
      formData.contractData?.chefHourlyRate || '0',
      formData.contractData?.numServers || '0',
      formData.contractData?.serverHours || '0',
      formData.contractData?.serverHourlyRate || '0',
      formData.contractData?.assistant || 'No',
      formData.contractData?.assistantHours || '0',
      formData.contractData?.assistantRate || '0',
      formData.contractData?.overtimeRate || '0',
      formData.contractData?.deliveryFee || '0',
      formData.contractData?.travelFee || '0',
      formData.contractData?.equipmentFee || '0',
      formData.contractData?.rentals || 'No',
      formData.contractData?.rentalCost || '0',
      formData.contractData?.disposablesCost || '0',
      formData.contractData?.chafersProvided || '',
      formData.contractData?.linenProvided || '',
      formData.contractData?.waterPitchersProvided || '',
      formData.contractData?.glasswareProvided || '',
      formData.contractData?.flowersProvided || '',
      formData.contractData?.otherRentalNotes || '',
      formData.contractData?.depositPercent || '0',
      formData.contractData?.depositDueDate || '',
      formData.contractData?.balanceDueDate || '',
      paymentMethods,
      includedServices,
      formData.contractData?.otherServicesNotes || '',
      dishesJson,
      contractDishesJson,
      formData.createQuote || false,
      formData.contractData?.createContract || false,
      formData.createReheatingInstructions || false,
      formData.createShoppingList || false,
      '', // Documents Folder URL - will be updated later
      '', // Quote Document URL - will be updated later
      '', // Labels Document URL - will be updated later
      '', // Schedule Document URL - will be updated later
      ''  // Contract Document URL - will be updated later
    ];
    
    // Write to sheet
    sheet.getRange(row, 1, 1, rowData.length).setValues([rowData]);
    
    console.log('Saved to spreadsheet row:', row);
    return row;
    
  } catch (error) {
    console.error('Error saving to spreadsheet:', error);
    throw error;
  }
}

function getPastSubmissions() {
  try {
    const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName('Responses');

    if (!sheet || sheet.getLastRow() <= 1) {
      return [];
    }

    // Read all columns
    const data = sheet
      .getRange(2, 1, sheet.getLastRow() - 1, 66)
      .getValues();

    const submissions = [];

    data.forEach(row => {
      if (!row[0]) return; // skip rows without submission ID

      const submission = {
        submissionId: String(row[0]),
        timestamp: row[1] instanceof Date ? row[1].toISOString() : String(row[1] || ''),
        clientName: String(row[2] || ''),
        email: String(row[3] || ''),
        employeeEmails: String(row[4] || ''),
        clientPhone: String(row[5] || ''),
        serviceType: String(row[6] || ''),
        dateRequested: row[7] instanceof Date ? row[7].toISOString().split('T')[0] : String(row[7] || ''),
        eventAddress: String(row[8] || ''),
        eventType: String(row[9] || ''),
        adults: Number(row[10] || 0),
        kids: Number(row[11] || 0),
        totalGuests: Number(row[12] || 0),
        eventTitle: String(row[13] || ''),
        eventName: String(row[14] || ''),
        adultPrice: Number(row[15] || 0),
        kidPrice: Number(row[16] || 0),
        options: {
          menuDesign: Boolean(row[17]),
          shopping: Boolean(row[18]),
          delivery: Boolean(row[19]),
          containerRemoval: Boolean(row[20])
        },
        serviceStyle: String(row[21] || ''),
        serviceStartTime: String(row[22] || ''),
        serviceEndTime: String(row[23] || ''),
        setupTime: Number(row[24] || 0),
        kitchenReset: String(row[25] || 'No'),
        menuFinalDate: row[26] instanceof Date ? row[26].toISOString().split('T')[0] : String(row[26] || ''),
        minGuestCount: Number(row[27] || 0),
        minFoodSpend: Number(row[28] || 0),
        chefLaborHours: Number(row[29] || 0),
        chefHourlyRate: Number(row[30] || 0),
        numServers: Number(row[31] || 0),
        serverHours: Number(row[32] || 0),
        serverHourlyRate: Number(row[33] || 0),
        assistant: String(row[34] || 'No'),
        assistantHours: Number(row[35] || 0),
        assistantRate: Number(row[36] || 0),
        overtimeRate: Number(row[37] || 0),
        deliveryFee: Number(row[38] || 0),
        travelFee: Number(row[39] || 0),
        equipmentFee: Number(row[40] || 0),
        rentals: String(row[41] || 'No'),
        rentalCost: Number(row[42] || 0),
        disposablesCost: Number(row[43] || 0),
        chafersProvided: String(row[44] || ''),
        linenProvided: String(row[45] || ''),
        waterPitchersProvided: String(row[46] || ''),
        glasswareProvided: String(row[47] || ''),
        flowersProvided: String(row[48] || ''),
        otherRentalNotes: String(row[49] || ''),
        depositPercent: Number(row[50] || 0),
        depositDueDate: row[51] instanceof Date ? row[51].toISOString().split('T')[0] : String(row[51] || ''),
        balanceDueDate: row[52] instanceof Date ? row[52].toISOString().split('T')[0] : String(row[52] || ''),
        paymentMethods: String(row[53] || '').split(', ').filter(m => m),
        includedServices: String(row[54] || '').split(', ').filter(s => s),
        otherServicesNotes: String(row[55] || ''),
        
        // Safe JSON parsing - REGULAR DISHES (always present)
        dishes: (() => {
          try {
            return JSON.parse(row[56] || '[]');
          } catch (e) {
            return [];
          }
        })(),
        
        // Safe JSON parsing - CONTRACT DISHES (may have extra fields)
        contractDishes: (() => {
          try {
            return JSON.parse(row[57] || '[]');
          } catch (e) {
            return [];
          }
        })(),
        
        createQuote: Boolean(row[58]),
        createContract: Boolean(row[59]),
        createReheatingInstructions: Boolean(row[60]),
        createShoppingList: Boolean(row[61])
      };

      submissions.push(submission);
    });

    Logger.log(`Loaded ${submissions.length} submissions`);
    return submissions;

  } catch (error) {
    console.error('Error getting submissions:', error);
    return [];
  }
}

// ==============================================
// LOAD SUBMISSION FUNCTION
// ==============================================
function getSubmissionById(submissionId) {
  try {
    const submissions = getPastSubmissions();
    const submission = submissions.find(sub => sub.submissionId === submissionId);
    
    if (!submission) return null;
    
    // For dishes, use the regular dishes array as the base
    const baseDishes = submission.dishes || [];
    
    // If contractDishes exist, merge their extended fields into the base dishes
    const contractDishes = submission.contractDishes || [];
    
    // Create a map of contract dishes by a unique key (date + time + name)
    const contractDishMap = {};
    contractDishes.forEach(cd => {
      const key = `${cd.date}_${cd.time}_${cd.name}`;
      contractDishMap[key] = cd;
    });
    
    // Merge the dish data, with contract dish fields taking precedence for extended fields
    const mergedDishes = baseDishes.map(dish => {
      const key = `${dish.date}_${dish.time}_${dish.name}`;
      const contractDish = contractDishMap[key] || {};
      
      return {
        ...dish,
        dietaryTags: contractDish.dietaryTags || dish.dietaryTags || [],
        portionFormat: contractDish.portionFormat || dish.portionFormat || '',
        clientNotes: contractDish.clientNotes || dish.clientNotes || '',
        internalNotes: contractDish.internalNotes || dish.internalNotes || ''
      };
    });
    
    // Add contract data structure for form
    submission.contractData = {
      createContract: submission.createContract,
      clientPhone: submission.clientPhone,
      eventAddress: submission.eventAddress,
      eventType: submission.eventType,
      guaranteedGuestDate: submission.guaranteedGuestDate || '',
      serviceStyle: submission.serviceStyle,
      serviceStartTime: submission.serviceStartTime,
      serviceEndTime: submission.serviceEndTime,
      setupTime: submission.setupTime,
      kitchenReset: submission.kitchenReset,
      menuFinalDate: submission.menuFinalDate,
      minGuestCount: submission.minGuestCount,
      minFoodSpend: submission.minFoodSpend,
      chefLaborHours: submission.chefLaborHours,
      chefHourlyRate: submission.chefHourlyRate,
      numServers: submission.numServers,
      serverHours: submission.serverHours,
      serverHourlyRate: submission.serverHourlyRate,
      assistant: submission.assistant,
      assistantHours: submission.assistantHours,
      assistantRate: submission.assistantRate,
      overtimeRate: submission.overtimeRate,
      deliveryFee: submission.deliveryFee,
      travelFee: submission.travelFee,
      equipmentFee: submission.equipmentFee,
      rentals: submission.rentals,
      rentalCost: submission.rentalCost,
      disposablesCost: submission.disposablesCost,
      chafersProvided: submission.chafersProvided,
      linenProvided: submission.linenProvided,
      waterPitchersProvided: submission.waterPitchersProvided,
      glasswareProvided: submission.glasswareProvided,
      flowersProvided: submission.flowersProvided,
      otherRentalNotes: submission.otherRentalNotes,
      depositPercent: submission.depositPercent,
      depositDueDate: submission.depositDueDate,
      balanceDueDate: submission.balanceDueDate,
      paymentMethods: submission.paymentMethods,
      includedServices: submission.includedServices,
      otherServicesNotes: submission.otherServicesNotes,
      contractDishes: mergedDishes
    };
    
    // Use the merged dishes for the main dishes array
    submission.dishes = mergedDishes;
    
    return submission;
    
  } catch (error) {
    console.error('Error getting submission by ID:', error);
    return null;
  }
}

function updateSubmissionFolderUrl(submissionId, folderUrl) {
  try {
    const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName('Responses');
    
    if (!sheet) return;
    
    // Find the submission row
    const ids = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
    for (let i = 0; i < ids.length; i++) {
      if (ids[i][0] === submissionId) {
        // Update folder URL (column 62)
        sheet.getRange(i + 2, 62).setValue(folderUrl);
        break;
      }
    }
  } catch (error) {
    console.error('Error updating folder URL:', error);
  }
}

function updateDocumentUrls(submissionId, quoteUrl, labelsUrl, scheduleUrl, contractUrl) {
  try {
    const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName('Responses');
    
    if (!sheet) return;
    
    // Find the submission row
    const ids = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
    for (let i = 0; i < ids.length; i++) {
      if (ids[i][0] === submissionId) {
        // Update document URLs at correct columns
        // Quote Document URL is column 63
        if (quoteUrl) sheet.getRange(i + 2, 63).setValue(quoteUrl);
        // Labels Document URL is column 64
        if (labelsUrl) sheet.getRange(i + 2, 64).setValue(labelsUrl);
        // Schedule Document URL is column 65
        if (scheduleUrl) sheet.getRange(i + 2, 65).setValue(scheduleUrl);
        // Contract Document URL is column 66
        if (contractUrl) sheet.getRange(i + 2, 66).setValue(contractUrl);
        break;
      }
    }
  } catch (error) {
    console.error('Error updating document URLs:', error);
  }
}

// ==============================================
// FOLDER MANAGEMENT
// ==============================================
function createSubmissionFolder(formData) {
  try {
    // Get the main folder
    const mainFolder = DriveApp.getFolderById(CONFIG.FOLDER_ID);
    
    // Clean client name for folder name
    const clientNameClean = formData.clientName.replace(/[^a-zA-Z0-9\s]/g, '').trim();
    const date = new Date(formData.dateRequested);
    const dateStr = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    
    // Create folder name
    let folderName;
    if (formData.eventTitle && formData.eventTitle.trim() !== '') {
      const eventTitleClean = formData.eventTitle.replace(/[^a-zA-Z0-9\s]/g, ' ').trim().substring(0, 30);
      folderName = `${dateStr} - ${eventTitleClean} - ${clientNameClean}`;
    } else {
      folderName = `${dateStr} - ${formData.serviceType} - ${clientNameClean}`;
    }
    
    // Check if folder already exists for this submission
    const existingFolders = mainFolder.getFoldersByName(folderName);
    if (existingFolders.hasNext()) {
      return existingFolders.next();
    }
    
    // Create new folder
    const submissionFolder = mainFolder.createFolder(folderName);
    
    // Update spreadsheet with folder URL
    updateSubmissionFolderUrl(formData.submissionId, submissionFolder.getUrl());
    
    console.log('Created submission folder:', folderName);
    return submissionFolder;
    
  } catch (error) {
    console.error('Error creating submission folder:', error);
    throw error;
  }
}

// ==============================================
// QUOTE DOCUMENT CREATION
// ==============================================
function createQuoteDocument(formData, submissionFolder) {
  // Generate document name from event title or client name
  let docName;
  const clientLastName = formData.clientName.split(' ').pop();
  const date = new Date(formData.dateRequested);
  const dateStr = Utilities.formatDate(date, Session.getScriptTimeZone(), 'MM.dd.yy');
  
  if (formData.eventTitle && formData.eventTitle.trim() !== '') {
    const eventTitleClean = formData.eventTitle.replace(/[^a-zA-Z0-9\s]/g, ' ').trim();
    docName = `${eventTitleClean}_Quote_${clientLastName}_${dateStr.replace(/\./g, '')}`;
  } else {
    docName = `${CONFIG.COMPANY_NAME.replace(/\s+/g, '_')}_Quote_${clientLastName}_${dateStr.replace(/\./g, '')}`;
  }
  
  console.log('Creating quote document:', docName);
  
  // Create the document
  const doc = DocumentApp.create(docName);
  
  // Add editors to the document
  addEditorsToDocument(doc.getId());
  
  const body = doc.getBody();
  
  // Clear default content
  body.clear();
  
  // Apply document styles
  body.setMarginTop(72);
  body.setMarginBottom(72);
  body.setMarginLeft(72);
  body.setMarginRight(72);
  body.setFontFamily('Georgia');
  body.setFontSize(11);
  
  // Add logo at the top-left
  try {
    if (CONFIG.LOGO_FILE_ID) {
      const logoBlob = DriveApp.getFileById(CONFIG.LOGO_FILE_ID).getBlob();
      
      const PT = 72;
      const logoWidth = 2 * PT;
      const logoHeight = 1.1 * PT;
      const tinyMargin = 0.02 * PT;
      
      const logoPara = body.insertParagraph(0, '');
      logoPara.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
      logoPara.setSpacingBefore(0);
      logoPara.setSpacingAfter(tinyMargin);
      
      const logoImage = logoPara.appendInlineImage(logoBlob);
      logoImage.setWidth(logoWidth).setHeight(logoHeight);
    }
  } catch (e) {
    console.log('Could not add logo to quote document:', e);
  }
  
  // Move document to submission folder
  moveFileToFolder(doc.getId(), submissionFolder.getId());
  
  // Build quote content
  buildQuoteContent(body, formData, doc);
  
  // Save and close
  doc.saveAndClose();
  
  return doc;
}

function addEditorsToDocument(docId) {
  try {
    const doc = DocumentApp.openById(docId);
    const file = DriveApp.getFileById(docId);
    
    // Add each editor
    CONFIG.EDITORS.forEach(email => {
      try {
        file.addEditor(email);
        console.log(`Added editor: ${email} to document ${docId}`);
      } catch (e) {
        console.warn(`Could not add editor ${email}:`, e.message);
      }
    });
  } catch (error) {
    console.error('Error adding editors to document:', error);
  }
}

function buildQuoteContent(body, formData, doc) {
  const eventDate = new Date(formData.dateRequested);
  const today = new Date();
  
  // ==============================================
  // COMPANY HEADER - MATCHING CONTRACT DOCUMENT
  // ==============================================
  
  // Company Name - Large, bold, centered
  const companyName = body.appendParagraph(CONFIG.COMPANY_NAME);
  companyName.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  companyName.setBold(true);
  companyName.setFontSize(18);
  companyName.setFontFamily('Georgia');
  companyName.setSpacingAfter(4);
  
  // Chef Name
  const chefName = body.appendParagraph(CONFIG.CHEF_NAME);
  chefName.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  chefName.setBold(false);
  chefName.setFontSize(12);
  chefName.setFontFamily('Georgia');
  chefName.setSpacingAfter(4);
  
  // Contact Information
  const contactInfo = body.appendParagraph(CONFIG.EMAIL + ' | ' + CONFIG.PHONE + ' | ' + CONFIG.WEBSITE);
  contactInfo.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  contactInfo.setBold(false);
  contactInfo.setFontSize(10);
  contactInfo.setFontFamily('Georgia');
  contactInfo.setSpacingAfter(4);
  
  // Address
  const address = body.appendParagraph(CONFIG.COMPANY_ADDRESS);
  address.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  address.setBold(false);
  address.setFontSize(10);
  address.setFontFamily('Georgia');
  address.setSpacingAfter(20);
  
  // Document Title
  const docTitle = body.appendParagraph('PROPOSAL');
  docTitle.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  docTitle.setBold(true);
  docTitle.setFontSize(14);
  docTitle.setFontFamily('Georgia');
  docTitle.setSpacingAfter(8);
  
  // Document Date
  const docDate = body.appendParagraph('Date: ' + formatDateLong(today));
  docDate.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  docDate.setBold(false);
  docDate.setFontSize(11);
  docDate.setFontFamily('Georgia');
  docDate.setSpacingAfter(24);
  
  // ==============================================
  // EVENT TITLE AND DATE
  // ==============================================
  
  // Add title - centered, bold
  const title = body.appendParagraph(formData.eventTitle || formData.clientName + ' Event');
  title.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  title.setBold(true);
  title.setFontSize(14);
  title.setFontFamily('Cormorant Upright');
  title.setSpacingAfter(8);
  
  // Add date in MM.dd.yy format - centered, bold
  const dateStr = Utilities.formatDate(eventDate, Session.getScriptTimeZone(), 'MM.dd.yy');
  const datePara = body.appendParagraph(dateStr);
  datePara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  datePara.setBold(true);
  datePara.setFontSize(14);
  datePara.setFontFamily('Cormorant Upright');
  datePara.setSpacingAfter(24);
  
  // ==============================================
  // CLIENT DETAILS
  // ==============================================
  
  // Client line
  const clientLinePara = body.appendParagraph('');
  const clientLabel = clientLinePara.appendText('Client: ');
  clientLabel.setBold(true);
  const clientName = clientLinePara.appendText(formData.clientName);
  clientName.setBold(false);
  clientLinePara.setFontSize(11);
  clientLinePara.setFontFamily('Georgia');
  clientLinePara.setSpacingAfter(8);
  
  // Event Date line
  const fullDateStr = Utilities.formatDate(eventDate, Session.getScriptTimeZone(), 'MMMM dd, yyyy');
  let eventText = fullDateStr;
  if (formData.eventTitle && formData.eventTitle.trim() !== '') {
    eventText = `${formData.eventTitle} ${fullDateStr}`;
  }
  
  const dateLinePara = body.appendParagraph('');
  const dateLabel = dateLinePara.appendText('Event Date: ');
  dateLabel.setBold(true);
  const dateValue = dateLinePara.appendText(eventText);
  dateValue.setBold(false);
  dateLinePara.setFontSize(11);
  dateLinePara.setFontFamily('Georgia');
  dateLinePara.setSpacingAfter(8);
  
  // Guests line
  const adults = parseInt(formData.adults) || 0;
  const kids = parseInt(formData.kids) || 0;
  let guestsText = '';
  if (adults > 0) guestsText += `${adults} Adult${adults > 1 ? 's' : ''}`;
  if (kids > 0) {
    if (guestsText) guestsText += ' + ';
    guestsText += `${kids} Kid${kids > 1 ? 's' : ''}`;
  }
  
  const guestsPara = body.appendParagraph('');
  const guestsLabel = guestsPara.appendText('Guests: ');
  guestsLabel.setBold(true);
  const guestsValue = guestsPara.appendText(guestsText);
  guestsValue.setBold(false);
  guestsPara.setFontSize(11);
  guestsPara.setFontFamily('Georgia');
  guestsPara.setSpacingAfter(30);
  
  // Menu Overview Header
  const menuHeader = body.appendParagraph('Menu Overview');
  menuHeader.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  menuHeader.setBold(true);
  menuHeader.setFontSize(12);
  menuHeader.setFontFamily('Georgia');
  menuHeader.setSpacingAfter(16);
  
  // Process dishes by meal type/category
  const dishesByCategory = organizeDishesByCategory(formData.dishes);
  
  // Add each category
  for (const [category, dishes] of Object.entries(dishesByCategory)) {
    if (category !== 'Uncategorized') {
      const categoryHeader = body.appendParagraph(category);
      categoryHeader.setBold(true);
      categoryHeader.setFontSize(11);
      categoryHeader.setFontFamily('Georgia');
      categoryHeader.setSpacingAfter(8);
    }
    
    dishes.forEach(dish => {
      let dishText = '';
      
      if (dish.name.includes('---')) {
        dishText = dish.name;
      } else if (dish.ingredients && dish.ingredients.trim() !== '') {
        dishText = `${dish.name} --- ${dish.ingredients}`;
      } else {
        dishText = dish.name;
      }
      
      if (dishText.includes('Choice') || dishText.includes('choice') || 
          dishText.includes('Choose') || dishText.includes('choose')) {
        const dishPara = body.appendParagraph(dishText);
        dishPara.setFontSize(11);
        dishPara.setFontFamily('Georgia');
        dishPara.setBold(false);
        dishPara.setSpacingAfter(8);
      } else {
        const dishPara = body.appendParagraph(`• ${dishText}`);
        dishPara.setFontSize(11);
        dishPara.setFontFamily('Georgia');
        dishPara.setBold(false);
        dishPara.setIndentStart(36);
        dishPara.setSpacingAfter(4);
      }
    });
    
    body.appendParagraph('');
  }
  
  // If no dishes were added, add a placeholder
  if (!formData.dishes || formData.dishes.length === 0) {
    const placeholder = body.appendParagraph('Menu to be customized based on client preferences.');
    placeholder.setFontSize(11);
    placeholder.setFontFamily('Georgia');
    placeholder.setBold(false);
    body.appendParagraph('');
  }
  
  // Pricing Section Header
  const pricingHeader = body.appendParagraph('Pricing');
  pricingHeader.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  pricingHeader.setBold(true);
  pricingHeader.setFontSize(12);
  pricingHeader.setFontFamily('Georgia');
  pricingHeader.setSpacingAfter(16);
  
  const adultPrice = parseFloat(formData.adultPrice) || 0;
  const kidPrice = parseFloat(formData.kidPrice) || 0;
  const adultTotal = adultPrice * adults;
  const kidTotal = kidPrice * kids;
  const subtotal = adultTotal + kidTotal;
  const tax = subtotal * CONFIG.TAX_RATE;
  const total = subtotal + tax;
  
  // Add pricing details
  if (adults > 0) {
    const adultLine = `Adults: ${adults} × $${adultPrice.toFixed(0)} = $${adultTotal.toFixed(0)}`;
    const adultPara = body.appendParagraph(adultLine);
    adultPara.setFontSize(11);
    adultPara.setFontFamily('Georgia');
    adultPara.setBold(false);
    adultPara.setSpacingAfter(4);
  }
  
  if (kids > 0) {
    const kidLine = `Kids: ${kids} × $${kidPrice.toFixed(0)} = $${kidTotal.toFixed(0)}`;
    const kidPara = body.appendParagraph(kidLine);
    kidPara.setFontSize(11);
    kidPara.setFontFamily('Georgia');
    kidPara.setBold(false);
    kidPara.setSpacingAfter(8);
  }
  
  body.appendParagraph('');
  
  // Total line
  const totalLine = body.appendParagraph('');
  const totalLabel = totalLine.appendText('Total Food & Chef Service: ');
  totalLabel.setBold(true);
  const totalValue = totalLine.appendText(`$${total.toFixed(0)} including tax`);
  totalValue.setBold(false);
  totalLine.setFontSize(11);
  totalLine.setFontFamily('Georgia');
  totalLine.setSpacingAfter(16);
  
  // Service inclusions
  const inclusions = body.appendParagraph('Includes menu planning, premium ingredient sourcing, preparation, and drop off and platting. No service staff and clean up included.');
  inclusions.setFontSize(11);
  inclusions.setFontFamily('Georgia');
  inclusions.setBold(false);
  inclusions.setItalic(true);
  inclusions.setSpacingAfter(8);
  
  // Additional notes about choices
  if (hasChoiceItems(formData.dishes)) {
    const choiceNote = body.appendParagraph('Client selects options as indicated above.');
    choiceNote.setFontSize(11);
    choiceNote.setFontFamily('Georgia');
    choiceNote.setBold(false);
    choiceNote.setItalic(true);
    choiceNote.setSpacingAfter(30);
  } else {
    body.appendParagraph('');
    body.appendParagraph('');
  }
  
  // Thank you message
  const thankYou = body.appendParagraph('Thank you for considering Zoma Culinary.');
  thankYou.setFontSize(11);
  thankYou.setFontFamily('Georgia');
  thankYou.setBold(false);
  thankYou.setSpacingAfter(16);
  
  // Signature section
  const warmRegards = body.appendParagraph('Warm Regards,');
  warmRegards.setFontSize(11);
  warmRegards.setFontFamily('Georgia');
  warmRegards.setBold(false);
  warmRegards.setSpacingAfter(8);
  
  // Add signature image if available
  try {
    if (CONFIG.SIGNATURE_IMAGE_ID && CONFIG.SIGNATURE_IMAGE_ID !== 'YOUR_SIGNATURE_IMAGE_ID') {
      const signatureBlob = DriveApp.getFileById(CONFIG.SIGNATURE_IMAGE_ID).getBlob();
      body.appendImage(signatureBlob);
      body.appendParagraph('');
    }
  } catch (e) {
    console.log('Could not add signature image:', e);
  }
  
  // Chef name
  const chefSignature = body.appendParagraph(CONFIG.CHEF_NAME);
  chefSignature.setFontSize(11);
  chefSignature.setFontFamily('Georgia');
  chefSignature.setBold(true);
}

function organizeDishesByCategory(dishes) {
  const categories = {
    'Starter': [],
    'Salad': [],
    'Main Courses': [],
    'Sides': [],
    'Dessert': [],
    'Kids Menu': [],
    'Uncategorized': []
  };
  
  if (!dishes) return categories;
  
  dishes.forEach(dish => {
    const time = dish.time || '';
    const name = dish.name || '';
    
    if (time.toLowerCase().includes('appetizer') || name.toLowerCase().includes('ceviche') || 
        name.toLowerCase().includes('bread') || name.toLowerCase().includes('starter')) {
      categories['Starter'].push(dish);
    } else if (time.toLowerCase().includes('salad') || name.toLowerCase().includes('salad')) {
      categories['Salad'].push(dish);
    } else if (time.toLowerCase().includes('main') || time.toLowerCase().includes('dinner') || 
               name.toLowerCase().includes('prime') || name.toLowerCase().includes('rib') ||
               name.toLowerCase().includes('pescado') || name.toLowerCase().includes('main')) {
      categories['Main Courses'].push(dish);
    } else if (time.toLowerCase().includes('side') || name.toLowerCase().includes('gnocchi') ||
               name.toLowerCase().includes('potato') || name.toLowerCase().includes('beans') ||
               name.toLowerCase().includes('brussels')) {
      categories['Sides'].push(dish);
    } else if (time.toLowerCase().includes('dessert') || name.toLowerCase().includes('tart') ||
               name.toLowerCase().includes('cake') || name.toLowerCase().includes('dessert')) {
      categories['Dessert'].push(dish);
    } else if (time.toLowerCase().includes('kids') || name.toLowerCase().includes('kids') ||
               name.toLowerCase().includes('child')) {
      categories['Kids Menu'].push(dish);
    } else {
      categories['Uncategorized'].push(dish);
    }
  });
  
  // Remove empty categories
  Object.keys(categories).forEach(category => {
    if (categories[category].length === 0) {
      delete categories[category];
    }
  });
  
  return categories;
}

function hasChoiceItems(dishes) {
  if (!dishes) return false;
  
  for (const dish of dishes) {
    const name = dish.name || '';
    if (name.includes('Choice') || name.includes('choice') || 
        name.includes('Choose') || name.includes('choose')) {
      return true;
    }
  }
  return false;
}

// ==============================================
// MENU LABELS DOCUMENT CREATION - UPDATED TO CREATE MULTIPLE DOCUMENTS
// ==============================================
function createMenuLabelsDocument(formData, submissionFolder) {
  const clientNameClean = formData.clientName.replace(/[^a-zA-Z0-9]/g, '_').substring(0, 30);
  const date = new Date(formData.dateRequested);
  const dateStr = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyyMMdd');
  
  console.log('Creating menu labels documents');
  
  const dishes = formData.dishes || [];
  if (dishes.length === 0) {
    return [];
  }
  
  const sortedDishes = sortDishesForDocuments(dishes);
  
  const allLabels = [];
  
  sortedDishes.forEach(dish => {
    allLabels.push({
      type: 'dish',
      dish: dish,
      formData: formData
    });
    
    if (dish.createIngredientLabels && dish.ingredients && dish.ingredients.trim() !== '') {
      const ingredients = dish.ingredients.split(',')
        .map(ing => ing.trim())
        .filter(ing => ing.length > 0);
      
      ingredients.forEach(ingredient => {
        allLabels.push({
          type: 'ingredient',
          dish: dish,
          ingredient: ingredient,
          formData: formData
        });
      });
    }
  });
  
  // If no labels to create, return empty array
  if (allLabels.length === 0) {
    return [];
  }
  
  const labelsPerDocument = 30;
  const totalDocuments = Math.ceil(allLabels.length / labelsPerDocument);
  
  console.log(`Total labels: ${allLabels.length}, will create ${totalDocuments} document(s)`);
  
  const createdDocs = [];
  
  for (let docIndex = 0; docIndex < totalDocuments; docIndex++) {
    // Create document name with suffix
    let docName;
    if (totalDocuments === 1) {
      docName = `${CONFIG.COMPANY_NAME}_Menu_Labels_${clientNameClean}_${dateStr}`;
    } else {
      docName = `${CONFIG.COMPANY_NAME}_Menu_Labels_${clientNameClean}_${dateStr}-${docIndex + 1}`;
    }
    
    console.log(`Creating labels document: ${docName}`);
    
    const doc = DocumentApp.create(docName);
    addEditorsToDocument(doc.getId());
    
    const body = doc.getBody();
    body.clear();
    
    body.setMarginTop(18);
    body.setMarginBottom(18);
    body.setMarginLeft(13.5);
    body.setMarginRight(13.5);
    body.setFontFamily('Cormorant Upright');
    
    moveFileToFolder(doc.getId(), submissionFolder.getId());
    
    // Calculate the range of labels for this document
    const startIndex = docIndex * labelsPerDocument;
    const endIndex = Math.min(startIndex + labelsPerDocument, allLabels.length);
    const labelsForThisDoc = allLabels.slice(startIndex, endIndex);
    
    // Build content for this document
    buildLabelsContentForDocument(body, labelsForThisDoc, docIndex + 1, totalDocuments);
    
    doc.saveAndClose();
    createdDocs.push(doc);
  }
  
  return createdDocs;
}

function buildLabelsContentForDocument(body, labels, docNumber, totalDocs) {
  if (labels.length === 0) {
    body.appendParagraph('No labels to display.');
    return;
  }
  
  const table = body.appendTable();
  table.setBorderWidth(0);
  
  let currentRow = null;
  let labelCount = 0;
  const labelsPerRow = 3;
  
  
  for (let i = 0; i < labels.length; i++) {
    const labelItem = labels[i];
    
    if (labelCount % labelsPerRow === 0) {
      currentRow = table.appendTableRow();
      currentRow.setMinimumHeight(72);
    }
    
    const cell = currentRow.appendTableCell();
    cell.setWidth(204);
    cell.setVerticalAlignment(DocumentApp.VerticalAlignment.TOP);
    cell.setPaddingTop(2);
    cell.setPaddingBottom(2);
    cell.setPaddingLeft(2);
    cell.setPaddingRight(2);
    
    const labelContent = cell.appendParagraph('');
    labelContent.setLineSpacing(1);
    labelContent.setSpacingBefore(0);
    labelContent.setSpacingAfter(0);
    
    if (labelItem.type === 'dish') {
      const dish = labelItem.dish;
      
      const dishName = labelContent.appendText(dish.name || 'Dish Name');
      dishName.setBold(true);
      dishName.setFontSize(13);
      dishName.setFontFamily('Cormorant Upright');
      
      labelContent.appendText('\n');
      
      let secondLine = '';
      if (dish.time) {
        secondLine += dish.time;
      }
      
      const dateText = dish.date ? formatDateShort(dish.date) : formatDateShort(labelItem.formData.dateRequested);
      if (secondLine) {
        secondLine += ' • ' + dateText;
      } else {
        secondLine = dateText;
      }
      
      const infoText = labelContent.appendText(secondLine);
      infoText.setBold(true);
      infoText.setFontSize(10);
      infoText.setFontFamily('Cormorant Upright');
    } else {
      const dish = labelItem.dish;
      const ingredient = labelItem.ingredient;
      
      const ingredientName = labelContent.appendText(ingredient);
      ingredientName.setBold(true);
      ingredientName.setFontSize(12);
      ingredientName.setFontFamily('Cormorant Upright');
      
      labelContent.appendText('\n');
      
      let secondLine = '';
      if (dish.time) {
        secondLine += dish.time;
      }
      
      const dateText = dish.date ? formatDateShort(dish.date) : formatDateShort(labelItem.formData.dateRequested);
      if (secondLine) {
        secondLine += ' • ' + dateText;
      } else {
        secondLine = dateText;
      }
      
      const infoText = labelContent.appendText(secondLine);
      infoText.setBold(true);
      infoText.setFontSize(10);
      infoText.setFontFamily('Cormorant Upright');
    }
    
    labelContent.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    
    labelCount++;
  }
  
  // Fill any remaining cells in the last row with empty cells
  const labelsInLastRow = labelCount % labelsPerRow;
  if (labelsInLastRow !== 0) {
    const emptyCellsNeeded = labelsPerRow - labelsInLastRow;
    for (let i = 0; i < emptyCellsNeeded; i++) {
      const emptyCell = currentRow.appendTableCell();
      emptyCell.setWidth(189);
    }
  }
}

// Original buildLabelsContent kept for backward compatibility
function buildLabelsContent(body, formData) {
  const dishes = formData.dishes || [];
  
  if (dishes.length === 0) {
    body.appendParagraph('No dishes specified for labels.');
    return;
  }
  
  const sortedDishes = sortDishesForDocuments(dishes);
  
  const allLabels = [];
  
  sortedDishes.forEach(dish => {
    allLabels.push({
      type: 'dish',
      dish: dish,
      formData: formData
    });
    
    if (dish.createIngredientLabels && dish.ingredients && dish.ingredients.trim() !== '') {
      const ingredients = dish.ingredients.split(',')
        .map(ing => ing.trim())
        .filter(ing => ing.length > 0);
      
      ingredients.forEach(ingredient => {
        allLabels.push({
          type: 'ingredient',
          dish: dish,
          ingredient: ingredient,
          formData: formData
        });
      });
    }
  });
  
  // Call the new function with the full label set
  buildLabelsContentForDocument(body, allLabels, 1, 1);
}

function formatDateShort(dateString) {
  try {
    const date = new Date(dateString);
    return Utilities.formatDate(date, Session.getScriptTimeZone(), 'MM.dd');
  } catch (e) {
    const match = dateString.match(/(\d{4})-(\d{2})-(\d{2})/);
    if (match) {
      return `${match[2]}.${match[3]}`;
    }
    return dateString;
  }
}

// ==============================================
// MENU SCHEDULE DOCUMENT CREATION
// ==============================================
function createMenuScheduleDocument(formData, submissionFolder) {
  const clientNameClean = formData.clientName.replace(/[^a-zA-Z0-9]/g, '_').substring(0, 30);
  const date = new Date(formData.dateRequested);
  const dateStr = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyyMMdd');
  const docName = `${CONFIG.COMPANY_NAME}_Menu_Schedule_${clientNameClean}_${dateStr}`;
  
  console.log('Creating menu schedule document:', docName);
  
  const doc = DocumentApp.create(docName);
  
  addEditorsToDocument(doc.getId());
  
  const body = doc.getBody();
  
  body.clear();
  
  try {
    if (CONFIG.LOGO_FILE_ID) {
      const logoBlob = DriveApp.getFileById(CONFIG.LOGO_FILE_ID).getBlob();
      
      const PT = 72;
      const logoWidth = 2 * PT;
      const logoHeight = 1.1 * PT;
      const tinyMargin = 0.02 * PT;
      
      const logoPara = body.insertParagraph(0, '');
      logoPara.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
      logoPara.setSpacingBefore(0);
      logoPara.setSpacingAfter(tinyMargin);
      
      const logoImage = logoPara.appendInlineImage(logoBlob);
      logoImage.setWidth(logoWidth).setHeight(logoHeight);
    }
  } catch (e) {
    console.log('Could not add logo:', e);
  }
  
  body.setMarginTop(72);
  body.setMarginBottom(72);
  body.setMarginLeft(72);
  body.setMarginRight(72);
  body.setFontFamily('Cormorant Upright');
  body.setFontSize(11);
  
  moveFileToFolder(doc.getId(), submissionFolder.getId());
  
  buildMenuScheduleContent(body, formData);
  
  doc.saveAndClose();
  
  return doc;
}

function buildMenuScheduleContent(body, formData) {
  const dishes = formData.dishes || [];
  
  if (dishes.length === 0) {
    body.appendParagraph('No dishes specified for menu schedule.');
    return;
  }
  
  const sortedDishes = sortDishesForDocuments(dishes);
  
  const dishesByDate = groupDishesByDateForSchedule(sortedDishes);
  
  let title = '';
  if (formData.eventTitle && formData.eventTitle.trim() !== '') {
    title = formData.eventTitle;
  } else {
    title = `${formData.clientName} ${formData.eventTitle}`;
  }
  
  const eventDate = new Date(formData.dateRequested);
  const dateStr = Utilities.formatDate(eventDate, Session.getScriptTimeZone(), 'MM.dd.yy');
  title += ` ${dateStr}`;
  
  const titlePara = body.appendParagraph(title);
  titlePara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  titlePara.setBold(true);
  titlePara.setFontSize(14);
  titlePara.setFontFamily('Cormorant Upright');
  titlePara.setSpacingAfter(24);
  
  const sortedDates = Object.keys(dishesByDate).sort();
  
  sortedDates.forEach(date => {
    const meals = dishesByDate[date];
    const dateShort = formatDateMenuSchedule(date);
    
    const sortedMealTypes = Object.keys(meals).sort((a, b) => {
      const orderA = mealTypeOrderForSorting[a] || 99;
      const orderB = mealTypeOrderForSorting[b] || 99;
      return orderA - orderB;
    });
    
    sortedMealTypes.forEach(mealType => {
      const mealDishes = meals[mealType];
      
      if (mealType !== 'Other') {
        const mealHeader = body.appendParagraph(mealType);
        mealHeader.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
        mealHeader.setBold(true);
        mealHeader.setFontSize(12);
        mealHeader.setFontFamily('Cormorant Upright');
        mealHeader.setSpacingBefore(16);
        mealHeader.setSpacingAfter(8);
      }
      
      mealDishes.forEach((dish, index) => {
        const dishPara = body.appendParagraph('');
        dishPara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
        
        let dishName = dish.name || 'Dish Name';
        dishName = dishName.replace(/\s*\d{1,2}\.\d{1,2}\.*\d{0,4}/g, '').trim();
        
        const dishNameText = dishPara.appendText(`${dishName} ${dateShort}`);
        dishNameText.setBold(true);
        dishNameText.setFontSize(12);
        
        if (dish.ingredients && dish.ingredients.trim() !== '') {
          dishPara.appendText('\n');
          const ingredientsText = dishPara.appendText(dish.ingredients);
          ingredientsText.setBold(false);
          ingredientsText.setFontSize(11);
        }
        
        dishPara.setSpacingBefore(mealType === 'Other' && index === 0 ? 16 : 8);
        dishPara.setSpacingAfter(8);
      });
    });
  });
  CONFIG.MENU_TEXT=body.getText();
}

// ==============================================
// CONTRACT DOCUMENT CREATION
// ==============================================
function createContractDocument(formData, submissionFolder) {
  const clientNameClean = formData.clientName.replace(/[^a-zA-Z0-9]/g, '_').substring(0, 30);
  const date = new Date(formData.dateRequested);
  const dateStr = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyyMMdd');
  const docName = `${CONFIG.COMPANY_NAME}_Contract_${clientNameClean}_${dateStr}`;
  
  console.log('Creating contract document:', docName);
  
  const doc = DocumentApp.create(docName);
  
  addEditorsToDocument(doc.getId());
  
  const body = doc.getBody();
  
  body.clear();
  
  body.setMarginTop(72);
  body.setMarginBottom(72);
  body.setMarginLeft(72);
  body.setMarginRight(72);
  body.setFontFamily('Cormorant Upright');
  body.setFontSize(11);
  
  try {
    if (CONFIG.LOGO_FILE_ID) {
      const logoBlob = DriveApp.getFileById(CONFIG.LOGO_FILE_ID).getBlob();
      
      const PT = 72;
      const logoWidth = 2 * PT;
      const logoHeight = 1.1 * PT;
      const tinyMargin = 0.02 * PT;
      
      const logoPara = body.insertParagraph(0, '');
      logoPara.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
      logoPara.setSpacingBefore(0);
      logoPara.setSpacingAfter(tinyMargin);
      
      const logoImage = logoPara.appendInlineImage(logoBlob);
      logoImage.setWidth(logoWidth).setHeight(logoHeight);
    }
  } catch (e) {
    console.log('Could not add logo:', e);
  }
  
  moveFileToFolder(doc.getId(), submissionFolder.getId());
  
  buildContractContent(body, formData);
  
  doc.saveAndClose();
  
  return doc;
}

function buildContractContent(body, formData) {
  const contractData = formData.contractData;
  const eventDate = new Date(formData.dateRequested);
  const today = new Date();
  
  // ==============================================
  // COMPANY HEADER
  // ==============================================
  
  const companyName = body.appendParagraph(CONFIG.COMPANY_NAME);
  companyName.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  companyName.setBold(true);
  companyName.setFontSize(18);
  companyName.setFontFamily('Georgia');
  companyName.setSpacingAfter(4);
  
  const chefName = body.appendParagraph(CONFIG.CHEF_NAME);
  chefName.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  chefName.setBold(false);
  chefName.setFontSize(12);
  chefName.setFontFamily('Georgia');
  chefName.setSpacingAfter(4);
  
  const contactInfo = body.appendParagraph(CONFIG.EMAIL + ' | ' + CONFIG.PHONE + ' | ' + CONFIG.WEBSITE);
  contactInfo.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  contactInfo.setBold(false);
  contactInfo.setFontSize(10);
  contactInfo.setFontFamily('Georgia');
  contactInfo.setSpacingAfter(4);
  
  const address = body.appendParagraph(CONFIG.COMPANY_ADDRESS);
  address.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  address.setBold(false);
  address.setFontSize(10);
  address.setFontFamily('Georgia');
  address.setSpacingAfter(20);
  
  const contractTitle = body.appendParagraph('CATERING SERVICES AGREEMENT');
  contractTitle.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  contractTitle.setBold(true);
  contractTitle.setFontSize(14);
  contractTitle.setFontFamily('Georgia');
  contractTitle.setSpacingAfter(8);
  
  const contractDate = body.appendParagraph('Date: ' + formatDateLong(today));
  contractDate.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  contractDate.setBold(false);
  contractDate.setFontSize(11);
  contractDate.setFontFamily('Georgia');
  contractDate.setSpacingAfter(24);
  
  // ==============================================
  // EVENT TITLE AND DATE
  // ==============================================
  
  const title = body.appendParagraph(formData.eventTitle || formData.clientName + ' Event');
  title.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  title.setBold(true);
  title.setFontSize(14);
  title.setFontFamily('Cormorant Upright');
  title.setSpacingAfter(8);
  
  const dateStr = Utilities.formatDate(eventDate, Session.getScriptTimeZone(), 'MM.dd.yy');
  const datePara = body.appendParagraph(dateStr);
  datePara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  datePara.setBold(true);
  datePara.setFontSize(14);
  datePara.setFontFamily('Cormorant Upright');
  datePara.setSpacingAfter(24);
  
  // ==============================================
  // MEAL EVENT OVERVIEW
  // ==============================================
  
  body.setFontFamily('Georgia');
  
  const overviewHeader = body.appendParagraph('MEAL EVENT OVERVIEW:');
  overviewHeader.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
  overviewHeader.setBold(true);
  overviewHeader.setFontSize(11);
  overviewHeader.setFontFamily('Georgia');
  overviewHeader.setSpacingAfter(16);
  
  const eventDetails = [
    ['Date:', formatDateLong(eventDate)],
    ['Location:', contractData.eventAddress || 'Event Location'],
    ['Client:', formData.clientName],
    ['Guest Count:', ((parseInt(formData.adults) || 0) + (parseInt(formData.kids) || 0)).toString()],
    ['Suggested Start time:', formatTime(contractData.serviceStartTime) || '6:00 PM'],
    ['Vibe:', contractData.eventType || 'Fun'],
    ['Timeline:', 'See above'],
    ['Chef and Team Arrival Time:', formatTime(contractData.teamArrivalTime) || '4:30 PM'],
    ['Food Start Time:', formatTime(contractData.serviceStartTime) || '6:30 PM'],
    ['Food Service End Time:', formatTime(contractData.serviceEndTime) || '10:00 PM'],
    ['Overall Length of Event:', calculateEventLength(contractData) || '6 hours']
  ];
  
  eventDetails.forEach(([label, value]) => {
    const line = body.appendParagraph(label + ' ' + value);
    line.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
    line.setBold(false);
    line.setFontSize(11);
    line.setFontFamily('Georgia');
    line.setSpacingAfter(4);
  });
  
  body.appendParagraph('');
  
  const equipmentHeader = body.appendParagraph('Equipment Provided:');
  equipmentHeader.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
  equipmentHeader.setBold(true);
  equipmentHeader.setFontSize(11);
  equipmentHeader.setFontFamily('Georgia');
  equipmentHeader.setSpacingAfter(8);
  
  const equipmentItems = [
    ['Plateware:', 'Client'],
    ['Flatware:', 'Client'],
    ['Linen:', 'Client'],
    ['Glassware:', 'Client'],
    ['Barwares:', 'Client']
  ];
  
  equipmentItems.forEach(([item, provider]) => {
    const line = body.appendParagraph(item + ' ' + provider);
    line.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
    line.setBold(false);
    line.setFontSize(11);
    line.setFontFamily('Georgia');
    line.setSpacingAfter(4);
  });
  
  body.appendParagraph('');
  
  const serviceHeader = body.appendParagraph(CONFIG.CHEF_NAME + ' and her team');
  serviceHeader.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
  serviceHeader.setBold(true);
  serviceHeader.setFontSize(11);
  serviceHeader.setFontFamily('Georgia');
  serviceHeader.setSpacingAfter(8);
  
  const serviceDesc = body.appendParagraph(
    `will handle the catering for the ${formData.eventTitle || 'Event'}, serving passed hors d'oeuvres and a late-night snack for ${((parseInt(formData.adults) || 0) + (parseInt(formData.kids) || 0))} guests.`
  );
  serviceDesc.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
  serviceDesc.setBold(false);
  serviceDesc.setFontSize(11);
  serviceDesc.setFontFamily('Georgia');
  serviceDesc.setSpacingAfter(8);
  
  const teamDesc = body.appendParagraph(
    'The team, including ' + CONFIG.CHEF_NAME + ', two kitchen assistants, and two servers, will manage tasks such as serving food, clearing tables, and assisting with the dessert table and the cheese and charcuterie station. To ensure a smooth event, we kindly request the provision of a small hot box and two chafing dishes for the late-night snack.'
  );
  teamDesc.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
  teamDesc.setBold(false);
  teamDesc.setFontSize(11);
  teamDesc.setFontFamily('Georgia');
  teamDesc.setSpacingAfter(8);
  
  const expectations = body.appendParagraph(
    'Our service expectations include catering to approximately ' + ((parseInt(formData.adults) || 0) + (parseInt(formData.kids) || 0)) + ' attendees, setting up and tidying the charcuterie and dessert stations, food service at the kitchen counter, and presenting passed hors d\'oeuvres when possible. The team will handle front-of-house (FOH) service, such as table clearing and tidying, as well as back-of-house (BOH) tidying and cleanup.'
  );
  expectations.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
  expectations.setBold(false);
  expectations.setFontSize(11);
  expectations.setFontFamily('Georgia');
  expectations.setSpacingAfter(8);
  
  const exclusions = body.appendParagraph(
    'Please note that bartending and alcohol service are excluded from our services. We look forward to contributing to the success of the event and ensuring a delightful experience for all attendees.'
  );
  exclusions.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
  exclusions.setBold(false);
  exclusions.setFontSize(11);
  exclusions.setFontFamily('Georgia');
  exclusions.setSpacingAfter(16);
  
  const cleanUpHeader = body.appendParagraph('Clean Up:');
  cleanUpHeader.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
  cleanUpHeader.setBold(true);
  cleanUpHeader.setFontSize(11);
  cleanUpHeader.setFontFamily('Georgia');
  cleanUpHeader.setSpacingAfter(8);
  
  const cleanUpDesc = body.appendParagraph(
    'The chef and team will handle cleanup responsibilities, including clearing, cleaning, and washing dinner plate ware, flatware, cookware, and glassware as necessary. They will also wipe down kitchen equipment, countertops, and sweep the floor before the agreed-upon "Departure Time." Additionally, the team will clear and clean the setup space, such as the garage. It\'s important to note that the team won\'t be responsible for clearing or cleaning any glassware, plateware, or flatware used after the "Departure Time."'
  );
  cleanUpDesc.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
  cleanUpDesc.setBold(false);
  cleanUpDesc.setFontSize(11);
  cleanUpDesc.setFontFamily('Georgia');
  cleanUpDesc.setSpacingAfter(16);
  
  const pricingHeader = body.appendParagraph('Pricing:');
  pricingHeader.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
  pricingHeader.setBold(true);
  pricingHeader.setFontSize(11);
  pricingHeader.setFontFamily('Georgia');
  pricingHeader.setSpacingAfter(8);
  
  const adultPrice = parseFloat(formData.adultPrice) || 0;
  const kidPrice = parseFloat(formData.kidPrice) || 0;
  const adults = parseInt(formData.adults) || 0;
  const kids = parseInt(formData.kids) || 0;
  const adultTotal = adultPrice * adults;
  const kidTotal = kidPrice * kids;
  const subtotal = adultTotal + kidTotal;
  const tax = subtotal * CONFIG.TAX_RATE;
  const total = subtotal + tax;
  
  const priceLine = body.appendParagraph('$' + total.toFixed(2) + ' (Inclusive of all taxes, fees, and gratuities)');
  priceLine.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
  priceLine.setBold(false);
  priceLine.setFontSize(11);
  priceLine.setFontFamily('Georgia');
  priceLine.setSpacingAfter(8);
  
  const paymentDue = body.appendParagraph('Payment due by: ' + formatDateLong(new Date(contractData.paymentDueDate || eventDate)));
  paymentDue.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
  paymentDue.setBold(false);
  paymentDue.setFontSize(11);
  paymentDue.setFontFamily('Georgia');
  paymentDue.setSpacingAfter(8);
  
  const paymentMethods = body.appendParagraph('Accepted Payment Methods: ' + 
    (contractData.paymentMethods ? contractData.paymentMethods.join(', ') : 'Check, Bank Transfer, Venmo, or Credit Card'));
  paymentMethods.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
  paymentMethods.setBold(false);
  paymentMethods.setFontSize(11);
  paymentMethods.setFontFamily('Georgia');
  paymentMethods.setSpacingAfter(8);
  
  const additionalCosts = body.appendParagraph('Additional Costs: None');
  additionalCosts.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
  additionalCosts.setBold(false);
  additionalCosts.setFontSize(11);
  additionalCosts.setFontFamily('Georgia');
  additionalCosts.setSpacingAfter(16);
  
  const agreementHeader = body.appendParagraph('Agreement');
  agreementHeader.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
  agreementHeader.setBold(true);
  agreementHeader.setFontSize(11);
  agreementHeader.setFontFamily('Georgia');
  agreementHeader.setSpacingAfter(8);
  
  const agreementText = body.appendParagraph(
    `This contract is entered into between ${CONFIG.COMPANY_NAME} / ${CONFIG.CHEF_NAME} and ${formData.clientName} on ${formatDateLong(today)}. The Client, ${formData.clientName}, is hiring ${CONFIG.CHEF_NAME} to provide food and related services for the ${formData.eventTitle || 'Event'}.`
  );
  agreementText.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
  agreementText.setBold(false);
  agreementText.setFontSize(11);
  agreementText.setFontFamily('Georgia');
  agreementText.setSpacingAfter(8);
  
  const locationLine = body.appendParagraph('Location: ' + (contractData.eventAddress || 'Event Location'));
  locationLine.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
  locationLine.setBold(false);
  locationLine.setFontSize(11);
  locationLine.setFontFamily('Georgia');
  locationLine.setSpacingAfter(8);
  
  const guestsLine = body.appendParagraph('Anticipated Guests: ' + ((parseInt(formData.adults) || 0) + (parseInt(formData.kids) || 0)));
  guestsLine.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
  guestsLine.setBold(false);
  guestsLine.setFontSize(11);
  guestsLine.setFontFamily('Georgia');
  guestsLine.setSpacingAfter(16);
  
  // ==============================================
  // MENU TO BE SERVED SECTION
  // ==============================================
  
  const menuServedHeader = body.appendParagraph('Menu to Be Served');
  menuServedHeader.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
  menuServedHeader.setBold(true);
  menuServedHeader.setFontSize(11);
  menuServedHeader.setFontFamily('Georgia');
  menuServedHeader.setSpacingAfter(8);
  
  const menuServedText = body.appendParagraph(
    'Both parties have agreed upon the menu, with any revisions to be completed by Monday, ' + 
    formatDateLong(new Date(contractData.menuFinalDate || eventDate)) + '. ' + 
    CONFIG.CHEF_NAME.split(' ')[1] + ' reserves the right to make minor menu adjustments due to ingredient availability.'
  );
  menuServedText.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
  menuServedText.setBold(false);
  menuServedText.setFontSize(11);
  menuServedText.setFontFamily('Georgia');
  menuServedText.setSpacingAfter(16);
  
  body.setFontFamily('Cormorant Upright');
  
  if (formData.dishes && formData.dishes.length > 0) {
    const sortedDishes = sortDishesForDocuments(formData.dishes);
    
    const dishesByDate = groupDishesByDateForSchedule(sortedDishes);
    
    const sortedDates = Object.keys(dishesByDate).sort();
    
    sortedDates.forEach(date => {
      const meals = dishesByDate[date];
      const dateShort = formatDateMenuSchedule(date);
      
      const sortedMealTypes = Object.keys(meals).sort((a, b) => {
        const orderA = mealTypeOrderForSorting[a] || 99;
        const orderB = mealTypeOrderForSorting[b] || 99;
        return orderA - orderB;
      });
      
      sortedMealTypes.forEach(mealType => {
        const mealDishes = meals[mealType];
        
        if (mealType !== 'Other') {
          const mealHeader = body.appendParagraph(mealType);
          mealHeader.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
          mealHeader.setBold(true);
          mealHeader.setFontSize(12);
          mealHeader.setFontFamily('Cormorant Upright');
          mealHeader.setSpacingBefore(16);
          mealHeader.setSpacingAfter(8);
        }
        
        mealDishes.forEach((dish, index) => {
          const dishPara = body.appendParagraph('');
          dishPara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
          
          let dishName = dish.name || 'Dish Name';
          dishName = dishName.replace(/\s*\d{1,2}\.\d{1,2}\.*\d{0,4}/g, '').trim();
          
          const dishNameText = dishPara.appendText(`${dishName} ${dateShort}`);
          dishNameText.setBold(true);
          dishNameText.setFontSize(12);
          
          if (dish.ingredients && dish.ingredients.trim() !== '') {
            dishPara.appendText('\n');
            const ingredientsText = dishPara.appendText(dish.ingredients);
            ingredientsText.setBold(false);
            ingredientsText.setFontSize(11);
          }
          
          dishPara.setSpacingBefore(mealType === 'Other' && index === 0 ? 16 : 8);
          dishPara.setSpacingAfter(8);
        });
      });
    });
  } else {
    const placeholder = body.appendParagraph('Menu to be finalized');
    placeholder.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    placeholder.setBold(false);
    placeholder.setFontSize(11);
    placeholder.setFontFamily('Cormorant Upright');
    placeholder.setSpacingAfter(24);
  }
  
  body.setFontFamily('Georgia');
  
  // ==============================================
  // STANDARD TERMS
  // ==============================================
  
  const standardTerms = [
    {
      title: 'Cancellation',
      content: 'The Client understands that upon entering this agreement, ' + CONFIG.CHEF_NAME.split(' ')[1] + ' is committing time and resources to this event. Therefore, cancellation would result in lost income and missed business opportunities. If the Client requests cancellation of this agreement within ' + (CONFIG.DEFAULT_CANCELLATION_DAYS || 5) + ' days of the event, ' + CONFIG.CHEF_NAME + ' shall be entitled to 50% of the estimated total cost.'
    },
    {
      title: 'Guest Count and Final Headcount',
      content: '- Client must provide an estimated guest count at booking and a final guaranteed headcount no later than [X] days before the event. Billing will be based on the guaranteed headcount or actual attendance, whichever is greater.'
    },
    {
      title: 'Food Safety, Allergies and Dietary Requirements',
      content: '- Caterer follows standard health and safety practices. Client must notify Caterer in writing of known allergies or dietary restrictions prior to the final menu confirmation.\n- Caterer is not responsible for adverse reactions where Client or guests fail to disclose allergies or for cross-contamination in on-site venues that limit safe preparation.'
    },
    {
      title: 'Liability and Insurance',
      content: '- Caterer carries general liability insurance and will provide proof upon request.\n- To the maximum extent permitted by law, neither party will be liable for indirect, incidental, or consequential damages.\n- Client is responsible for property damage caused by event guests, except for damage resulting from Caterer\'s gross negligence or willful misconduct.'
    },
    {
      title: 'Indemnification',
      content: '- Each party indemnifies and holds the other harmless from claims arising from its negligence or willful misconduct related to the event.'
    },
    {
      title: 'Force Majeure',
      content: '- Neither party is liable for delays or failures caused by circumstances beyond reasonable control (e.g., severe weather, acts of God, strikes). The parties will attempt to reschedule; if rescheduling is impossible, cancellation provisions apply.'
    },
    {
      title: 'Termination for Cause',
      content: '- Either party may terminate this agreement for material breach if the breaching party fails to cure within [X] days of written notice. Termination does not relieve Client of payment obligations for services already performed or irrevocable commitments made on Client\'s behalf.'
    },
    {
      title: 'Alcohol Service',
      content: '- If alcohol is provided, Caterer (or licensed bartending subcontractor) will comply with applicable laws and venue policies. Client is responsible for any permits and for ensuring minors are not served. Client assumes liability for overconsumption by guests to the extent permitted by law.'
    },
    {
      title: 'Governing Law and Dispute Resolution',
      content: '- This Agreement is governed by the laws of the State of Texas.\n- Parties will attempt to resolve disputes in good faith; unresolved disputes shall be submitted to binding arbitration in [City, State] (or specify court jurisdiction).'
    },
    {
      title: 'Entire Agreement; Amendments',
      content: '- This Agreement, including Exhibits, constitutes the entire agreement and supersedes prior discussions. Amendments must be in writing and signed by both parties.'
    },
    {
      title: 'Severability',
      content: '- If any provision is found invalid, the remainder of the Agreement remains enforceable.'
    }
  ];
  
  standardTerms.forEach((term, index) => {
    const termNumber = index + 1;
    const termHeader = body.appendParagraph(termNumber + '. ' + term.title);
    termHeader.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
    termHeader.setBold(true);
    termHeader.setFontSize(11);
    termHeader.setFontFamily('Georgia');
    termHeader.setSpacingAfter(8);
    
    const termContent = body.appendParagraph(term.content);
    termContent.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
    termContent.setBold(false);
    termContent.setFontSize(11);
    termContent.setFontFamily('Georgia');
    termContent.setSpacingAfter(16);
  });
  
  const signaturesHeader = body.appendParagraph('20. Signatures');
  signaturesHeader.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
  signaturesHeader.setBold(true);
  signaturesHeader.setFontSize(11);
  signaturesHeader.setFontFamily('Georgia');
  signaturesHeader.setSpacingAfter(16);
  
  const clientSignature = body.appendParagraph('- Client: _________________________________________________________ Date: ______________');
  clientSignature.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
  clientSignature.setBold(false);
  clientSignature.setFontSize(11);
  clientSignature.setFontFamily('Georgia');
  clientSignature.setSpacingAfter(4);
  
  const clientName = body.appendParagraph('Name: [' + formData.clientName + '], Title (if applicable): [Title]');
  clientName.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
  clientName.setBold(false);
  clientName.setFontSize(11);
  clientName.setFontFamily('Georgia');
  clientName.setSpacingAfter(16);
  
  const catererSignature = body.appendParagraph('- Caterer: _______________________________________________________ Date: ______________');
  catererSignature.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
  catererSignature.setBold(false);
  catererSignature.setFontSize(11);
  catererSignature.setFontFamily('Georgia');
  catererSignature.setSpacingAfter(4);
  
  const catererName = body.appendParagraph('Name: [' + CONFIG.SIGNATURE_NAME + '], Title: [Owner/Manager]');
  catererName.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
  catererName.setBold(false);
  catererName.setFontSize(11);
  catererName.setFontFamily('Georgia');
  catererName.setSpacingAfter(4);
}

function calculateEventLength(contractData) {
  if (!contractData.serviceStartTime || !contractData.serviceEndTime) {
    return '6 hours';
  }
  
  try {
    const startTime = new Date('1970-01-01T' + contractData.serviceStartTime + ':00');
    const endTime = new Date('1970-01-01T' + contractData.serviceEndTime + ':00');
    
    const diffMs = endTime - startTime;
    const diffHours = Math.floor(diffMs / (1000 * 60 * 60));
    
    return diffHours + ' hour' + (diffHours !== 1 ? 's' : '');
  } catch (e) {
    return '6 hours';
  }
}

function formatTime(timeString) {
  if (!timeString) return '';
  
  try {
    const [hours, minutes] = timeString.split(':');
    const hourNum = parseInt(hours);
    
    if (hourNum >= 12) {
      const displayHour = hourNum > 12 ? hourNum - 12 : hourNum;
      return displayHour + ':' + (minutes || '00') + ' PM';
    } else {
      const displayHour = hourNum === 0 ? 12 : hourNum;
      return displayHour + ':' + (minutes || '00') + ' AM';
    }
  } catch (e) {
    return timeString;
  }
}

function formatDateLong(date) {
  if (!date) return '';
  
  try {
    const monthNames = ['January', 'February', 'March', 'April', 'May', 'June',
                       'July', 'August', 'September', 'October', 'November', 'December'];
    
    const d = new Date(date);
    const month = monthNames[d.getMonth()];
    const day = d.getDate();
    const year = d.getFullYear();
    
    return month + ' ' + day + ', ' + year;
  } catch (e) {
    return date.toString();
  }
}

// ==============================================
// SORTING FUNCTIONS
// ==============================================

const mealTypeOrderForSorting = {
  'Appetizer': 1,
  'Breakfast': 2,
  'Lunch': 3,
  'Dinner': 4,
  'Snack': 5,
  'Dessert': 6
};

function sortDishesForDocuments(dishes) {
  return dishes.sort((a, b) => {
    const dateA = new Date(a.date || '1970-01-01');
    const dateB = new Date(b.date || '1970-01-01');
    if (dateA.getTime() !== dateB.getTime()) {
      return dateA.getTime() - dateB.getTime();
    }
    
    const orderA = mealTypeOrderForSorting[a.time] || 99;
    const orderB = mealTypeOrderForSorting[b.time] || 99;
    return orderA - orderB;
  });
}

function formatDateMenuSchedule(dateString) {
  try {
    const date = new Date(dateString);
    return Utilities.formatDate(date, Session.getScriptTimeZone(), 'MM.dd');
  } catch (e) {
    const match = dateString.match(/(\d{4})-(\d{2})-(\d{2})/);
    if (match) {
      return `${match[2]}.${match[3]}`;
    }
    return dateString;
  }
}

function groupDishesByDateForSchedule(dishes) {
  const dishesByDate = {};
  
  dishes.forEach(dish => {
    const date = dish.date || '';
    if (!dishesByDate[date]) {
      dishesByDate[date] = {};
    }
    
    const mealType = dish.time || 'Other';
    if (!dishesByDate[date][mealType]) {
      dishesByDate[date][mealType] = [];
    }
    
    dishesByDate[date][mealType].push(dish);
  });
  
  return dishesByDate;
}

// ==============================================
// PREVIEW FUNCTIONALITY
// ==============================================

function generateMenuPreview(dishes) {
  const sortedDishes = sortDishesForDocuments(dishes);
  
  const dishesByDate = groupDishesByDateForSchedule(sortedDishes);
  
  let html = '<div style="font-family: \'Cormorant Upright\', serif; font-size: 14px; line-height: 1.5; max-width: 600px; margin: 0 auto; padding: 20px;">';
  
  const eventDate = dishes.length > 0 && dishes[0].date ? 
    formatDateForPreview(dishes[0].date) : 
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MM.dd.yy');
  
  html += `<div style="text-align: center; font-weight: bold; font-size: 16px; margin-bottom: 24px;">
    Menu Schedule ${eventDate}
  </div>`;
  
  const sortedDates = Object.keys(dishesByDate).sort();
  
  sortedDates.forEach(date => {
    const meals = dishesByDate[date];
    const dateShort = formatDateMenuSchedule(date);
    
    const sortedMealTypes = Object.keys(meals).sort((a, b) => {
      const orderA = mealTypeOrderForSorting[a] || 99;
      const orderB = mealTypeOrderForSorting[b] || 99;
      return orderA - orderB;
    });
    
    sortedMealTypes.forEach(mealType => {
      const mealDishes = meals[mealType];
      
      if (mealType !== 'Other') {
        html += `<div style="text-align: center; font-weight: bold; font-size: 12px; 
                 margin-top: 16px; margin-bottom: 8px;">
          ${mealType}
        </div>`;
      }
      
      mealDishes.forEach((dish, index) => {
        let dishName = dish.name || 'Dish Name';
        dishName = dishName.replace(/\s*\d{1,2}\.\d{1,2}\.*\d{0,4}/g, '').trim();
        
        html += `<div style="text-align: center; 
                 margin-top: ${mealType === 'Other' && index === 0 ? '16px' : '8px'}; 
                 margin-bottom: 8px;">`;
        
        html += `<div style="font-weight: bold; font-size: 14px;">
          ${dishName} ${dateShort}
        </div>`;
        
        if (dish.ingredients && dish.ingredients.trim() !== '') {
          html += `<div style="font-size: 14px; margin-top: 4px; color: #555;">
            ${dish.ingredients}
          </div>`;
        }
        
        if (dish.createIngredientLabels) {
          html += `<div style="font-size: 12px; font-style: italic; color: #888; margin-top: 2px;">
            <i class="fas fa-tag"></i> Ingredient labels will be created
          </div>`;
        }
        
        html += `</div>`;
      });
    });
  });
  
  const totalDishes = sortedDishes.length;
  const totalDates = sortedDates.length;
  
  html += `<div style="margin-top: 30px; padding: 15px; background: #f8f9fa; 
           border-left: 4px solid #4361ee; border-radius: 4px;">
    <div style="font-weight: bold; color: #4361ee; margin-bottom: 8px;">
      <i class="fas fa-chart-bar"></i> Menu Summary
    </div>
    <div style="font-size: 14px;">
      • ${totalDishes} dish${totalDishes !== 1 ? 'es' : ''} across ${totalDates} date${totalDates !== 1 ? 's' : ''}<br>
      • Sorted by date, then by meal type: Appetizer → Breakfast → Lunch → Dinner → Snack → Dessert
    </div>
  </div>`;
  
  return html;
}

function formatDateForPreview(dateString) {
  try {
    const date = new Date(dateString);
    return Utilities.formatDate(date, Session.getScriptTimeZone(), 'MM.dd.yy');
  } catch (e) {
    const match = dateString.match(/(\d{4})-(\d{2})-(\d{2})/);
    if (match) {
      return `${match[2]}.${match[3]}.${match[1].substring(2)}`;
    }
    return dateString;
  }
}

function getMenuPreview(dishesData) {
  try {
    const dishes = JSON.parse(dishesData);
    
    if (!dishes || dishes.length === 0) {
      return {
        success: false,
        html: '<div style="text-align: center; padding: 40px; color: #666; font-family: \'Cormorant Upright\', serif;">No dishes to preview</div>'
      };
    }
    
    const previewHtml = generateMenuPreview(dishes);
    
    return {
      success: true,
      html: previewHtml
    };
    
  } catch (error) {
    console.error('Error generating preview:', error);
    return {
      success: false,
      html: `<div style="text-align: center; padding: 40px; color: #dc3545; font-family: \'Cormorant Upright\', serif;">
        Error generating preview: ${error.message}
      </div>`
    };
  }
}

// ==============================================
// HELPER FUNCTIONS
// ==============================================
function validateFormData(formData) {
  if (!formData.clientName || formData.clientName.toString().trim() === '') {
    return { isValid: false, message: 'Client Name is required' };
  }
  
  if (!formData.email || formData.email.toString().trim() === '') {
    return { isValid: false, message: 'Client Email is required' };
  }
  
  const clientEmails = formData.email.split(',')
    .map(email => email.trim())
    .filter(email => email && isValidEmail(email));
  
  if (clientEmails.length === 0) {
    return { isValid: false, message: 'At least one valid client email is required' };
  }
  
  if (formData.employeeEmails && formData.employeeEmails.trim() !== '') {
    const employeeEmails = formData.employeeEmails.split(',')
      .map(email => email.trim())
      .filter(email => email);
    
    const invalidEmails = employeeEmails.filter(email => !isValidEmail(email));
    if (invalidEmails.length > 0) {
      return { isValid: false, message: `Invalid employee email format: ${invalidEmails.join(', ')}` };
    }
  }
  
  if (!formData.dateRequested || formData.dateRequested.toString().trim() === '') {
    return { isValid: false, message: 'Date Requested is required' };
  }
  
  if (formData.contractData?.createContract) {
    if (!formData.contractData.eventAddress || formData.contractData.eventAddress.trim() === '') {
      return { isValid: false, message: 'Event Address is required for contract documents' };
    }
    
    if (!formData.contractData.eventType || formData.contractData.eventType.trim() === '') {
      return { isValid: false, message: 'Event Type is required for contract documents' };
    }
    
    if (!formData.contractData.depositDueDate || formData.contractData.depositDueDate.trim() === '') {
      return { isValid: false, message: 'Deposit Due Date is required for contract documents' };
    }
  }
  
  if (formData.createQuote) {
    if (!formData.serviceType || formData.serviceType.toString().trim() === '') {
      return { isValid: false, message: 'Service Type is required for quote documents' };
    }
    
    if (!formData.adultPrice || formData.adultPrice.toString().trim() === '') {
      return { isValid: false, message: 'Price per Adult is required for quote documents' };
    }
    
    const adults = parseInt(formData.adults) || 0;
    if (adults === 0) {
      return { isValid: false, message: 'At least one adult is required for quote documents' };
    }
  }
  
  return { isValid: true, message: 'OK' };
}

function isValidEmail(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

function formatDate(dateString) {
  try {
    const date = new Date(dateString);
    return Utilities.formatDate(date, Session.getScriptTimeZone(), 'MMMM dd, yyyy');
  } catch (e) {
    return dateString;
  }
}

function createPDF(doc) {
  try {
    const file = DriveApp.getFileById(doc.getId());
    const pdf = file.getAs('application/pdf');
    pdf.setName(doc.getName() + '.pdf');
    return pdf;
  } catch (error) {
    console.error('Error creating PDF:', error);
    return null;
  }
}

function moveFileToFolder(fileId, folderId) {
  try {
    if (!folderId) {
      console.log('No folder ID specified, document will remain in root');
      return;
    }
    
    const file = DriveApp.getFileById(fileId);
    const folder = DriveApp.getFolderById(folderId);
    
    file.getParents().next().removeFile(file);
    folder.addFile(file);
    
    console.log('Moved file to folder:', folder.getName());
  } catch (error) {
    console.error('Error moving file to folder:', error);
  }
}

// ==============================================
// GEMINI API CALL FUNCTION - Updated to accept file attachments
// ==============================================
function callGeminiAPIWithFile(prompt, fileId) {
  try {
    const apiKey = CONFIG.GEMINI_API_KEY;
    
    if (!apiKey || apiKey === 'YOUR_GEMINI_API_KEY') {
      throw new Error('Gemini API key not configured');
    }
    
    // If a file is provided, upload it first
    if (fileId) {
      const file = DriveApp.getFileById(fileId);
      const fileBlob = file.getBlob();
      const fileData = Utilities.base64Encode(fileBlob.getBytes());
      const mimeType = fileBlob.getContentType();
      
      const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-3-flash-preview:generateContent?key=${apiKey}`;
      
      const payload = {
        contents: [
          {
            parts: [
              {
                text: prompt
              },
              {
                inline_data: {
                  mime_type: mimeType,
                  data: fileData
                }
              }
            ]
          }
        ],
        generationConfig: {
          temperature: 0.7,
          maxOutputTokens: 2048,
          topP: 0.95,
          topK: 40
        }
      };
      
      const options = {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      };
      
      const response = UrlFetchApp.fetch(url, options);
      const responseCode = response.getResponseCode();
      const responseText = response.getContentText();
      
      if (responseCode !== 200) {
        console.error('Gemini API error:', responseCode, responseText);
        throw new Error(`Gemini API returned status ${responseCode}`);
      }
      
      const result = JSON.parse(responseText);
      
      // Extract the generated text from the response
      if (result.candidates && result.candidates.length > 0 && 
          result.candidates[0].content && 
          result.candidates[0].content.parts && 
          result.candidates[0].content.parts.length > 0) {
        return result.candidates[0].content.parts[0].text;
      } else {
        throw new Error('Unexpected Gemini API response format');
      }
    } else {
      // Fall back to text-only if no file provided
      const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=${apiKey}`;
      
      const payload = {
        contents: [
          {
            parts: [
              {
                text: prompt
              }
            ]
          }
        ],
        generationConfig: {
          temperature: 0.7,
          maxOutputTokens: 2048,
          topP: 0.95,
          topK: 40
        }
      };
      
      const options = {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      };
      
      const response = UrlFetchApp.fetch(url, options);
      const responseCode = response.getResponseCode();
      const responseText = response.getContentText();
      
      if (responseCode !== 200) {
        console.error('Gemini API error:', responseCode, responseText);
        throw new Error(`Gemini API returned status ${responseCode}`);
      }
      
      const result = JSON.parse(responseText);
      
      // Extract the generated text from the response
      if (result.candidates && result.candidates.length > 0 && 
          result.candidates[0].content && 
          result.candidates[0].content.parts && 
          result.candidates[0].content.parts.length > 0) {
        return result.candidates[0].content.parts[0].text;
      } else {
        throw new Error('Unexpected Gemini API response format');
      }
    }
    
  } catch (error) {
    console.error('Error calling Gemini API:', error);
    throw error;
  }
}

// ==============================================
// REHEATING INSTRUCTIONS EMAIL FUNCTION - WITH GEMINI API USING PDF
// ==============================================
function sendReheatingInstructionsEmail(formData, submissionFolder) {
  try {
    const dishes = formData.dishes || [];
    
    if (dishes.length === 0) {
      console.log('No dishes to create reheating instructions for');
      return false;
    }
    
    // Get the menu schedule document ID from the submission
    let menuScheduleDoc = null;
    
    // Check if we already have a menu schedule document for this submission
    const folderFiles = submissionFolder.getFiles();
    while (folderFiles.hasNext()) {
      const file = folderFiles.next();
      if (file.getName().includes('Menu_Schedule') && file.getMimeType() === MimeType.GOOGLE_DOCS) {
        menuScheduleDoc = DocumentApp.openById(file.getId());
        break;
      }
    }
    
    // If we don't have a menu schedule document yet, create it
    if (!menuScheduleDoc) {
      menuScheduleDoc = createMenuScheduleDocument(formData, submissionFolder);
    }
    
    // Create the prompt for Gemini
    const prompt = `Zoma Culinary Reheating Instructions Generator

Create simple reheating instructions for the menu provided in the attached document.

Requirements:

Write one short paragraph per dish.

Include microwave and oven instructions when applicable.

Keep instructions clear, simple, and client-friendly.

Avoid overly technical language.

Mention if any fresh ingredients should be added after reheating (sauces, avocado, cucumber, herbs, etc.).

Format the response like this:

Dish Name
Short reheating paragraph including microwave and oven instructions.

The tone should be clean, professional, and concise, suitable for a private chef meal delivery client.

Please analyze the attached menu document and generate instructions for all dishes shown.`;

    // Call Gemini API with the menu schedule document attached
    const instructions = callGeminiAPIWithFile(prompt, menuScheduleDoc.getId());
    
    // Create a document with the instructions
    const clientNameClean = formData.clientName.replace(/[^a-zA-Z0-9]/g, '_').substring(0, 30);
    const date = new Date();
    const dateStr = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyyMMdd');
    const docName = `${CONFIG.COMPANY_NAME}_Reheating_Instructions_${clientNameClean}_${dateStr}`;
    
    const doc = DocumentApp.create(docName);
    addEditorsToDocument(doc.getId());
    
    const body = doc.getBody();
    body.clear();
    
    // Add header
    body.appendParagraph('REHEATING INSTRUCTIONS').setHeading(DocumentApp.ParagraphHeading.HEADING1);
    body.appendParagraph(`Menu: ${formData.eventTitle || formData.clientName} - ${formatDateLong(formData.dateRequested)}`);
    body.appendParagraph('');
    
    // Add instructions from Gemini
    body.appendParagraph(instructions);
    
    doc.saveAndClose();
    
    // Move to submission folder
    moveFileToFolder(doc.getId(), submissionFolder.getId());
    
    // Create PDF
    const pdf = createPDF(doc);
    
    // Send email to culinary team
    const recipients = ['auraofblue@gmail.com', 'martha@zomaculinary.com', 'jasonhurley@zomaculinary.com'];
    const subject = `Reheating Instructions: ${formData.eventTitle || formData.clientName} - Zoma Automation Bot`;
    
    let emailBody = `Hello Team,\n\n`;
    emailBody += `Please find the reheating instructions for the upcoming event attached.\n\n`;
    emailBody += `Event: ${formData.eventTitle || formData.clientName}\n`;
    emailBody += `Date: ${formatDateLong(formData.dateRequested)}\n\n`;
    emailBody += `The instructions were generated by AI based on the menu schedule document.\n\n`;
    emailBody += `Best,\n`;
    emailBody += `Zoma Automation Bot`;
    
    MailApp.sendEmail({
      to: recipients.join(','),
      subject: subject,
      body: emailBody,
      attachments: [pdf],
      name: CONFIG.COMPANY_NAME
    });
    
    console.log('Reheating instructions email sent to culinary team');
    return true;
    
  } catch (error) {
    console.error('Error sending reheating instructions email:', error);
    return false;
  }
}

// ==============================================
// SHOPPING LIST EMAIL FUNCTION - WITH GEMINI API USING PDF
// ==============================================
function sendShoppingListEmail(formData, submissionFolder) {
  try {
    const clientNameClean = formData.clientName.replace(/[^a-zA-Z0-9]/g, '_').substring(0, 30);
    const dishes = formData.dishes || [];
    
    if (dishes.length === 0) {
      console.log('No dishes to create shopping list for');
      return false;
    }
    
    // Get the menu schedule document ID from the submission
    let menuScheduleDoc = null;
    
    // Check if we already have a menu schedule document for this submission
    const folderFiles = submissionFolder.getFiles();
    while (folderFiles.hasNext()) {
      const file = folderFiles.next();
      if (file.getName().includes('Menu_Schedule') && file.getMimeType() === MimeType.GOOGLE_DOCS) {
        menuScheduleDoc = DocumentApp.openById(file.getId());
        break;
      }
    }
    
    // If we don't have a menu schedule document yet, create it
    if (!menuScheduleDoc) {
      menuScheduleDoc = createMenuScheduleDocument(formData, submissionFolder);
    }
    
    // Create the prompt for Gemini
    const prompt = `ZOMA CULINARY – SHOPPING LIST GENERATOR

Create a grocery shopping list from the menu in the attached document.

Follow these instructions EXACTLY.

Organize the shopping list into these four sections ONLY:

Protein
Fruits & Vegetables
Dairy
Dry Goods

Use this exact structure:

ZOMA CULINARY
Shopping List - ${clientNameClean} Menu & Delivery [Week]

Client: ${clientNameClean}
Meals: 1 Client

Protein
• ingredient + amount

Fruits & Vegetables
• ingredient + amount

Dairy
• ingredient + amount

Dry Goods
• ingredient + amount

Every ingredient MUST include a realistic amount or quantity appropriate for 1 client.

Examples:
• ½ lb ground turkey
• 2 carrots
• 1 cup spinach
• 1 can chickpeas
• 2 tbsp tahini

Combine duplicate ingredients across meals.

Example:
If carrots appear in multiple meals, list one total quantity.

Include supporting ingredients needed to cook the dishes such as:
olive oil, vinegar, spices, tortillas, tahini, etc.

Keep the list clean and optimized for grocery shopping and chef prep.

Please analyze the attached menu document and generate a comprehensive shopping list based on all dishes shown.`;

    // Call Gemini API with the menu schedule document attached
    const shoppingListEmail = callGeminiAPIWithFile(prompt, menuScheduleDoc.getId());
    
    // Create a document with the shopping list
    const date = new Date();
    const dateStr = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyyMMdd');
    const docName = `${clientNameClean}_Shopping_List`;
    
    const doc = DocumentApp.create(docName);
    addEditorsToDocument(doc.getId());
    
    const body = doc.getBody();
    body.clear();
    
    // Add header
    body.appendParagraph('SHOPPING LIST').setHeading(DocumentApp.ParagraphHeading.HEADING1);
    body.appendParagraph(`Menu: ${formData.eventTitle || formData.clientName} - ${formatDateLong(formData.dateRequested)}`);
    body.appendParagraph('');
    
    // Add shopping list from Gemini
    body.appendParagraph(shoppingListEmail);
    
    doc.saveAndClose();
    
    // Move to submission folder
    moveFileToFolder(doc.getId(), submissionFolder.getId());
    
    // Create PDF
    const pdf = createPDF(doc);
    
    // Extract the email body parts (subject and body)
    const emailParts = parseShoppingListEmail(shoppingListEmail, formData);
    
    // Send email to client
    const empEmails = formData.employeeEmails.split(',').map(e => e.trim()).filter(e => isValidEmail(e));
    
    if (empEmails.length > 0) {
      MailApp.sendEmail({
        to: empEmails[0],
        cc: empEmails.slice(1).join(','),
        subject: emailParts.subject,
        body: emailParts.body,
        attachments: [pdf],
        name: CONFIG.COMPANY_NAME
      });
      
      console.log('Shopping list email sent to client');
    }
    
    return true;
    
  } catch (error) {
    console.error('Error sending shopping list email:', error);
    return false;
  }
}

// Helper function to parse the Gemini response into subject and body
function parseShoppingListEmail(geminiResponse, formData) {
  try {
    // Default subject
    let subject = `Shopping List: Zoma Culinary - ${formData.eventTitle || formData.clientName} ${formatDateLong(formData.dateRequested)}`;
    let body = geminiResponse;
    
    // Try to extract subject if it's in the format "Subject: ..."
    const subjectMatch = geminiResponse.match(/Subject:\s*(.+?)(?:\n|$)/i);
    if (subjectMatch) {
      subject = subjectMatch[1].trim();
      // Remove the subject line from the body
      body = geminiResponse.replace(/Subject:\s*.+?\n/i, '').trim();
    }
    
    // Replace [Name] with actual client first name
    const firstName = formData.clientName.split(' ')[0];
    body = body.replace(/\[Name\]/g, firstName);
    body = body.replace(/Hi \[Name\]/g, `Hi ${firstName}`);
    
    // Replace [Insert Menu Name/Date] with actual values
    const menuInfo = `${formData.eventTitle || formData.clientName} ${formatDateLong(formData.dateRequested)}`;
    body = body.replace(/\[Insert Menu Name\/Date\]/g, menuInfo);
    
    return { subject, body };
  } catch (e) {
    console.error('Error parsing shopping list email:', e);
    return {
      subject: `Shopping List: Zoma Culinary - ${formData.eventTitle || formData.clientName} ${formatDateLong(formData.dateRequested)}`,
      body: geminiResponse
    };
  }
}

// Keep the original callGeminiAPI for backward compatibility if needed
function callGeminiAPI(prompt) {
  return callGeminiAPIWithFile(prompt, null);
}

// ==============================================
// UPDATED EMAIL FUNCTION WITH SWAPPED LABEL DISTRIBUTION
// ==============================================
function sendConfirmationEmail(formData, quoteDoc, labelsDocs, scheduleDoc, contractDoc, 
                               quotePdf, labelsPdfs, schedulePdf, contractPdf, 
                               includeQuote, includeContract) {
  try {
    let clientEmails = [];
    if (formData.email && formData.email.trim() !== '') {
      clientEmails = formData.email.split(',')
        .map(email => email.trim())
        .filter(email => email && isValidEmail(email));
    }
    
    let employeeEmails = [];
    if (formData.employeeEmails && formData.employeeEmails.trim() !== '') {
      employeeEmails = formData.employeeEmails.split(',')
        .map(email => email.trim())
        .filter(email => email && isValidEmail(email));
    }
    
    let allRecipients = [...clientEmails, ...employeeEmails];
    
    allRecipients = [...new Set(allRecipients)];
    
    if (allRecipients.length === 0) {
      console.log('No valid email recipients');
      return;
    }
    
    const subject = `${CONFIG.COMPANY_NAME} - Documents for ${formData.eventTitle || formData.clientName}`;
    
    // Get all label documents for this submission from stored property if not provided directly
    let allLabelDocs = labelsDocs || [];
    let allLabelPdfs = labelsPdfs || [];
    
    if ((!allLabelDocs || allLabelDocs.length === 0) && formData.submissionId) {
      try {
        const storedDocIds = PropertiesService.getScriptProperties().getProperty('labelsDocs_' + formData.submissionId);
        if (storedDocIds) {
          const docIds = JSON.parse(storedDocIds);
          allLabelDocs = docIds.map(id => {
            try {
              return DocumentApp.openById(id);
            } catch (e) {
              console.error(`Could not open document ${id}:`, e);
              return null;
            }
          }).filter(doc => doc);
          allLabelPdfs = allLabelDocs.map(doc => createPDF(doc)).filter(pdf => pdf);
          console.log(`Retrieved ${allLabelDocs.length} label documents from stored property`);
        }
      } catch (e) {
        console.error('Error retrieving label documents from stored property:', e);
      }
    }
    
    // Send email to each recipient individually
    allRecipients.forEach(recipient => {
      const isClient = clientEmails.includes(recipient);
      const isEmployee = employeeEmails.includes(recipient);
      
      // Build document links based on recipient type
      let documentLinks = [];
      
      if (includeContract && contractDoc) {
        documentLinks.push(`Contract Agreement\nhttps://docs.google.com/open?id=${contractDoc.getId()}`);
      }
      
      if (scheduleDoc) {
        documentLinks.push(`Menu Schedule\nhttps://docs.google.com/open?id=${scheduleDoc.getId()}`);
      }
      
      if (includeQuote && quoteDoc) {
        documentLinks.push(`Proposal\nhttps://docs.google.com/open?id=${quoteDoc.getId()}`);
      }
      
      // ADD LABELS LINKS FOR EMPLOYEES (multiple documents)
      if (isEmployee && allLabelDocs.length > 0) {
        if (allLabelDocs.length === 1) {
          documentLinks.push(`Menu Labels\nhttps://docs.google.com/open?id=${allLabelDocs[0].getId()}`);
        } else {
          allLabelDocs.forEach((doc, index) => {
            documentLinks.push(`Menu Labels (Part ${index + 1})\nhttps://docs.google.com/open?id=${doc.getId()}`);
          });
        }
      }
      
      // Build email body
      let body = `Dear ${formData.clientName},\n\n`;
      body += `Thank you for inviting Zoma Culinary to be part of your event. It is truly an honor to be part of such special occasion.\n\n`;
      body += `Please find the following documents prepared for your review:\n\n`;
      body += documentLinks.join('\n\n');
      body += `\n\nCopies are attached for your review.\n\n`;
      body += `When convenient, please look over the proposal and confirm that everything aligns. Upon your approval, I will issue the formal agreement for signature. To secure your date, the signed contract and deposit will be required.\n\n`;
      body += `Should you wish to adjust anything or discuss details further, I am always happy to connect.\n\n`;
      body += `Warmly,\n`;
      body += `${CONFIG.CHEF_NAME}\n`;
      body += `${CONFIG.COMPANY_NAME}\n`;
      body += `${CONFIG.EMAIL}\n`;
      body += `${CONFIG.PHONE}`;
      
      // Prepare attachments based on recipient type
      let attachments = [];
      
      // Common documents for everyone
      if (schedulePdf) attachments.push(schedulePdf);
      if (includeQuote && quotePdf) attachments.push(quotePdf);
      if (includeContract && contractPdf) attachments.push(contractPdf);
      
      // ADD ALL LABELS ATTACHMENTS FOR EMPLOYEES
      if (isEmployee && allLabelPdfs.length > 0) {
        allLabelPdfs.forEach(pdf => {
          if (pdf) attachments.push(pdf);
        });
      }
      
      try {
        MailApp.sendEmail({
          to: recipient,
          subject: subject,
          body: body,
          attachments: attachments,
          name: CONFIG.COMPANY_NAME
        });
        
        console.log(`Email sent to: ${recipient}${isClient ? ' (client)' : ' (employee)'}`);
        if (isEmployee && allLabelDocs.length > 0) {
          console.log(`  - ${allLabelDocs.length} label document(s) included for employee`);
        }
      } catch (emailError) {
        console.error(`Error sending email to ${recipient}:`, emailError);
      }
    });
    
    // Clean up stored document IDs after sending
    if (formData.submissionId) {
      PropertiesService.getScriptProperties().deleteProperty('labelsDocs_' + formData.submissionId);
    }
    
  } catch (error) {
    console.error('Error sending emails:', error);
  }
}

// ==============================================
// WEB APP FUNCTIONS
// ==============================================
function loadSubmission(submissionId) {
  return getSubmissionById(submissionId);
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Zoma Culinary Service Request')
    .setWidth(1200)
    .setHeight(800);
}

// ==============================================
// SETUP AND INITIALIZATION
// ==============================================
function setupSpreadsheet() {
  try {
    const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID)
    
    const responsesSheet = spreadsheet.insertSheet('Responses');
    
    const headers = [
      'Submission ID',
      'Timestamp',
      'Client Name',
      'Email',
      'Employee Emails',
      'Client Phone',
      'Service Type',
      'Date Requested',
      'Event Address',
      'Event Type',
      'Adults',
      'Kids',
      'Total Guests',
      'Event Title',
      'Event Name',
      'Adult Price',
      'Kid Price',
      'Menu Design',
      'Shopping',
      'Delivery',
      'Container Removal',
      'Service Style',
      'Service Start Time',
      'Service End Time',
      'Setup Time',
      'Kitchen Reset',
      'Menu Final Date',
      'Min Guest Count',
      'Min Food Spend',
      'Chef Labor Hours',
      'Chef Hourly Rate',
      'Num Servers',
      'Server Hours',
      'Server Hourly Rate',
      'Assistant',
      'Assistant Hours',
      'Assistant Rate',
      'Overtime Rate',
      'Delivery Fee',
      'Travel Fee',
      'Equipment Fee',
      'Rentals',
      'Rental Cost',
      'Disposables Cost',
      'Chafers Provided',
      'Linen Provided',
      'Water Pitchers Provided',
      'Glassware Provided',
      'Flowers Provided',
      'Other Rental Notes',
      'Deposit Percent',
      'Deposit Due Date',
      'Balance Due Date',
      'Payment Methods',
      'Included Services',
      'Other Services Notes',
      'Dishes JSON',
      'Contract Dishes JSON',
      'Create Quote',
      'Create Contract',
      'Create Reheating Instructions',
      'Create Shopping List',
      'Documents Folder URL',
      'Quote Document URL',
      'Labels Document URL',
      'Schedule Document URL',
      'Contract Document URL'
    ];
    
    responsesSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    responsesSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    
    console.log('Created new spreadsheet:', spreadsheet.getId());
    console.log('Please update CONFIG.SPREADSHEET_ID with:', spreadsheet.getId());
    
    return spreadsheet.getId();
  } catch (error) {
    console.error('Error setting up spreadsheet:', error);
    throw error;
  }
}

// ==============================================
// TEST FUNCTION
// ==============================================
function testDocumentCreation() {
  const testData = {
    clientName: 'Lawson Family',
    email: 'client@example.com, client2@example.com',
    employeeEmails: 'employee1@company.com, employee2@company.com',
    serviceType: 'Weekly Chef Services',
    dateRequested: '2025-11-26',
    adults: '4',
    kids: '0',
    eventTitle: 'Lawson Family Dinner',
    eventName: 'Weekly Meal Service',
    adultPrice: '85',
    kidPrice: '0',
    createReheatingInstructions: true,
    createShoppingList: true,
    options: {
      menuDesign: true,
      shopping: true,
      delivery: true,
      containerRemoval: false
    },
    createQuote: true,
    dishes: [
      { date: '2025-11-26', time: 'Dinner', name: 'Chicken a la Vodka', ingredients: 'Roasted Potatoes, Butternut Squash', createIngredientLabels: true, dietaryTags: ['GF'], portionFormat: 'Per Person', clientNotes: 'Gluten-free option available' },
      { date: '2025-11-27', time: 'Dinner', name: 'Chile Rellenos', ingredients: 'Tomato sauce, Cilantro Rice, Black Beans', createIngredientLabels: true },
      { date: '2025-11-28', time: 'Dinner', name: 'Pork Meatballs', ingredients: 'Rigatoni Pasta, Asparagus', createIngredientLabels: false },
      { date: '2025-11-29', time: 'Dinner', name: 'Grilled New York Strip', ingredients: 'Potato Pave, Roasted Broccoli, Herb Butter', createIngredientLabels: true }
    ],
    contractData: {
      createContract: true,
      clientPhone: '(123) 456-7890',
      eventAddress: '123 Main St, Austin, TX 78701',
      eventType: 'Private Party',
      serviceStyle: 'Plated',
      serviceStartTime: '18:00',
      serviceEndTime: '22:00',
      setupTime: '2',
      kitchenReset: 'Yes',
      menuFinalDate: '2025-11-19',
      minGuestCount: '4',
      chefLaborHours: '8',
      chefHourlyRate: '75',
      numServers: '2',
      serverHours: '6',
      serverHourlyRate: '35',
      assistant: 'Yes',
      assistantHours: '6',
      assistantRate: '25',
      overtimeRate: '100',
      deliveryFee: '0',
      travelFee: '0',
      equipmentFee: '0',
      rentals: 'No',
      depositPercent: '50',
      depositDueDate: '2025-11-19',
      balanceDueDate: '2025-11-26',
      paymentMethods: ['Check', 'Bank Transfer'],
      includedServices: ['MenuPlanning', 'OnsiteExecution', 'SetupService', 'KitchenReset']
    }
  };
  
  const result = saveForm(testData);
  Logger.log('Test result:', result);
  return result;
}
