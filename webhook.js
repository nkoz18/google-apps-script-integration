const _ = LodashGS.load();

/**
 * Recieves the webhook POST request from Active Campaign
 *
 * @param {object} e POST request event object
 * @return {object} message recieved object
 */
function doPost(e) {
  Logger.log(`Deal update event triggered!`);
  Logger.log(JSON.stringify(e.parameters));

  const stageTitle = _.get(e, `parameter["deal[stage_title]"]`, "");

  // Check the updated fields to see if the deal was just moved into the discovery stage and does not already have a job number
  let jobDiscoveryFlag = false;
  let dealOrgAssigned = false;
  let stageUpdated = false;
  let jobNumber;

  console.log("Parameter: ", JSON.stringify(e.parameter));

  for (const [key, value] of Object.entries(e.parameter)) {
    // console.log(`${key}: ${value}`);

    if (key.includes("deal[orgname]") && value) {
      Logger.log("The deal has an organization assigned.");
      dealOrgAssigned = true;
    }
    if (key.includes("updated_fields") && value === "stage") {
      Logger.log("One of the updated fields is the stage.");
      stageUpdated = true;
    }

    // Check to see if the job number exists and filter out keys of recently updated fields
    if (value === "Job Number" && !key.includes("updated_fields")) {
      const jobNumberValueKey = key.replace("[key]", "[value]");
      jobNumber = e.parameter[jobNumberValueKey];
      Logger.log("Job Number is: " + jobNumber);
    }
  }

  // remove this later
  if (!dealOrgAssigned) {
    Logger.log("Deal has no account set, stopping automation.");
  }

  if (
    stageTitle === "Discovery" &&
    !jobNumber &&
    stageUpdated &&
    dealOrgAssigned
  ) {
    jobDiscoveryFlag = true;
    Logger.log(
      'A deal without a job number was just updated to the "Discovery" stage.'
    );
  }

  const dealId = _.get(e, `parameters["deal[id]"][0]`, null);
  Logger.log(`Deal ID: ` + JSON.stringify(dealId));

  const contactId = _.get(e, `parameters["deal[contactid]"][0]`, null);
  Logger.log(`Contact ID: ` + JSON.stringify(contactId));

  if (dealId && jobDiscoveryFlag) {
    const dealObj = _.get(getDeal(dealId), `deal`, null);
    const contactObj = _.get(getContact(contactId), `contact`, null);

    if (dealObj) {
      const newJobNumber = getLatestJobNumber() + 1;
      const dealId = dealObj.id;
      const dealTitle = dealObj.title;
      const dealUserId = dealObj.owner;
      let newJobFolder;
      let organizationObj;
      let organizationCustomFields = {};

      // Update Active Campaign deal with new job number
      updateDealWithJobNumber(dealId, dealTitle, newJobNumber);

      const organizationId = _.get(dealObj, `organization`);

      // if there is an orgnization add it to the job title in Google Drive
      if (organizationId) {
        organizationObj = getDealOrganization(organizationId);

        newJobFolder = createNewJobFolder(
          newJobNumber,
          dealTitle,
          organizationObj.account.name
        );

        // Get the organization custom fields
        organizationCustomFields = getOrganizationCustomFields(organizationId);
      } else {
        newJobFolder = createNewJobFolder(newJobNumber, dealTitle);
      }

      // Get the contact custom fields
      const contactCustomFields = getContactCustomFields(contactId);

      // Get the user that is the owner of the deal
      const dealUserObj = getDealUser(dealUserId);

      // Get the deal custom fields
      const dealCustomFields = getDealCustomFields(dealId);

      // Get the deal notes
      const dealNotes = getDealNotes(dealId);

      let jobObj = {
        jobNumber: newJobNumber,
        jobFolderUrl: newJobFolder.getUrl(),
        jobFolderId: newJobFolder.getId(),
        contact: {
          ...contactCustomFields,
          ...contactObj,
        },
        deal: {
          ...dealCustomFields,
          ...dealNotes,
          ...dealObj,
        },
        organization: {
          ...organizationCustomFields,
          ...organizationObj,
        },
        user: dealUserObj.user,
      };

      jobObj = appendCustomFieldsToJobObject(jobObj);

      Logger.log("New Job Object");
      Logger.log(JSON.stringify(jobObj));

      // Update the Customer Inquiry Tracking Sheet
      updateCustomerInquiryTrackingSheet(jobObj);

      // Create Discovery Job Folders
      createDiscoveryJobFolders(newJobFolder, jobObj);

      // Update Active Campaign with Google Drive Resources
      updateDealGoogleFolderIds(jobObj);
    }
  }

  return ContentService.createTextOutput(
    JSON.stringify({
      message: "Request received.",
    })
  );
}

/**
 * Searches the "Customer Inquiry Tracking Sheet" for the last row
 * with data in the current year and find the latest job number
 *
 * @return {number} the job number of the last job
 */
function getLatestJobNumber() {
  let customerInquiryTrackingSheet = SpreadsheetApp.openById(
    CREDENTIALS.customerInquiryTrackingSheet
  );
  let currYear = new Date().getFullYear().toString();
  let currYearSheet = customerInquiryTrackingSheet.getSheetByName(currYear);
  let lastRow = currYearSheet.getLastRow();
  let latestJobRange = currYearSheet.getRange(lastRow, 2);
  let latestJobNumber = latestJobRange.getValue();

  Logger.log("Latest job number is: " + latestJobNumber.toString());

  return latestJobNumber;
}

/**
 * Accepts the job object and updates the
 * "Customer Inquiry Tracking Sheet" with the new job
 *
 * @param {object} jobObj an enourmous object of job data
 */
function updateCustomerInquiryTrackingSheet(jobObj) {
  let customerInquiryTrackingSheet = SpreadsheetApp.openById(
    CREDENTIALS.customerInquiryTrackingSheet
  );

  let currYear = new Date().getFullYear().toString();
  let currYearSheet = customerInquiryTrackingSheet.getSheetByName(currYear);

  currYearSheet.appendRow([
    // Captain / Owner
    jobObj.user.firstName.concat(" ", jobObj.user.lastName),
    // Job Number
    jobObj.jobNumber,
    // Inquiry Date / Date Created
    new Date().toLocaleDateString("en-US"),
    // OG CRM Status - Always "Awaiting Info" at this stage
    "Awaiting Info",
    // Client Company Name
    jobObj.organization?.account?.name,
    // Contact Name
    jobObj.contact["firstName"] + " " + jobObj.contact["lastName"],
    // Job Title - Not to be confused with the deal title
    // TODO: Should this be searching to see if the contact is working for the deal organization and choose that job title?
    _.get(jobObj, "contact.accountContacts[0].jobTitle", ""),
    //Description
    _.get(jobObj, "deal.description"),
    // End User
    jobObj.deal.endUser,
    // Full Name
    jobObj.fullJobFolderName,
    // NDA
    jobObj.deal.nda,
    // Date Estimate Sent
    jobObj.deal.dateEstimateSent
      ? new Date(jobObj.deal.dateEstimateSent).toLocaleDateString("en-US")
      : "",
    // Deal Value / Budget
    jobObj.deal.formattedValue,
    // Channel - Deprecated Field
    "",
    // Type - Deprecated Field
    "",
    // Industry
    jobObj.organization.industry,
    //Lead Type
    jobObj.contact.leadType,
    // Why Dead - Deprecated Field
    "",
    // Client Email
    jobObj.contact["email"],
    // Notes
    _.get(jobObj, "deal.notes[0].note"),
  ]);
}

/**
 * Process the objects returned by the Active Campaign API
 * to extract the relevant data for this job
 *
 * @param {object} jobObj an enourmous object of job data
 * @return {object} an enourmous object of job data
 */
function appendCustomFieldsToJobObject(jobObj) {
  jobObj.fullJobFolderName = jobObj.jobNumber
    .toString()
    .concat(" ", jobObj.deal.title);

  if (!isEmpty(jobObj.organization)) {
    jobObj.fullJobFolderName = jobObj.fullJobFolderName.concat(
      " ",
      jobObj.organization.account.name
    );
  }

  // Get Date Estimate was sent
  const dateEstimateSentObj = jobObj.deal.dealCustomFieldData.find(
    (customField) =>
      customField.customFieldId === CUSTOM_FIELD_ID_MAP.DEAL.DATE_ESTIMATE_SENT
  );

  jobObj.deal.dateEstimateSent = _.get(dateEstimateSentObj, "fieldValue", "");

  // Get End User
  const endUserObj = jobObj.deal.dealCustomFieldData.find(
    (customField) =>
      customField.customFieldId === CUSTOM_FIELD_ID_MAP.DEAL.END_USER
  );

  jobObj.deal.endUser = _.get(endUserObj, "fieldValue", "");

  // Get NDA
  const ndaObj = jobObj.deal.dealCustomFieldData.find(
    (customField) => customField.customFieldId === CUSTOM_FIELD_ID_MAP.DEAL.NDA
  );

  jobObj.deal.nda = _.get(ndaObj, "fieldValue[0]", "");

  // Get Industry

  let industryObj;

  const organizationCustomFieldData = _.get(
    jobObj.organization.customerAccountCustomFieldData,
    null
  );

  if (organizationCustomFieldData) {
    industryObj = organizationCustomFieldData.find(
      (customField) =>
        customField.custom_field_id ===
        CUSTOM_FIELD_ID_MAP.ORGANIZATION.INDUSTRY.toString() // These IDs are stored as strings!
    );
  }

  jobObj.organization.industry = _.get(
    industryObj,
    "custom_field_text_value",
    ""
  );

  // Get Lead Type
  const leadTypeObj = jobObj.contact.fieldValues.find(
    (customField) =>
      customField.field === CUSTOM_FIELD_ID_MAP.CONTACT.LEAD_SOURCE.toString() // These IDs are stored as strings!
  );

  jobObj.contact.leadType = _.get(leadTypeObj, "value", "");

  // Get Delivery Location
  const deliveryLocationObj = jobObj.deal.dealCustomFieldData.find(
    (customField) =>
      customField.customFieldId === CUSTOM_FIELD_ID_MAP.DEAL.DELIVERY_LOCATION
  );

  jobObj.deal.deliveryLocation = _.get(deliveryLocationObj, "fieldValue", "");

  // Get Delivery Location
  const deliveryDateObj = jobObj.deal.dealCustomFieldData.find(
    (customField) =>
      customField.customFieldId === CUSTOM_FIELD_ID_MAP.DEAL.DELIVERY_DATE
  );

  jobObj.deal.deliveryDate = _.get(deliveryDateObj, "fieldValue", "");

  // Get Install Date
  const installDateObj = jobObj.deal.dealCustomFieldData.find(
    (customField) =>
      customField.customFieldId === CUSTOM_FIELD_ID_MAP.DEAL.INSTALL_DATE
  );

  jobObj.deal.installDate = _.get(installDateObj, "fieldValue", "");

  // Get Install Date
  const strikeDateObj = jobObj.deal.dealCustomFieldData.find(
    (customField) =>
      customField.customFieldId === CUSTOM_FIELD_ID_MAP.DEAL.STRIKE_DATE
  );

  jobObj.deal.strikeDate = _.get(strikeDateObj, "fieldValue", "");

  // Get Company Phone
  const organizationPhoneObj =
    jobObj.organization.customerAccountCustomFieldData.find(
      (customField) =>
        customField.custom_field_id ===
        CUSTOM_FIELD_ID_MAP.ORGANIZATION.PHONE.toString() // These IDs are stored as strings!
    );

  // Format the value of the deal
  // Active Campaign returns a string with the cents numbers included and no decimal ?

  jobObj.deal.formattedValue = _.get(jobObj.deal, "value", "");

  if (jobObj.deal.formattedValue) {
    jobObj.deal.formattedValue =
      jobObj.deal.formattedValue.substring(
        0,
        jobObj.deal.formattedValue.length - 2
      ) +
      "." +
      jobObj.deal.formattedValue.substring(
        jobObj.deal.formattedValue.length - 2
      );
  }

  jobObj.organization.account.phone = _.get(
    organizationPhoneObj,
    "custom_field_text_value",
    ""
  );

  return jobObj;
}

/**
 * Fetches the organization from the Active Campaign API
 *
 * @param {number} organizationId the organization ID in Active Campaign
 * @return {object} the result of the fetch from the Active Campaign API
 */
function getDealOrganization(organizationId) {
  const result = fetchFromActiveCampaign(
    "get",
    CREDENTIALS.activeCampaignBaseUrl + "accounts/" + organizationId
  );

  Logger.log("Organization Fetch Result: " + JSON.stringify(result));

  return result;
}

/**
 * Fetches the organization custom fields from the Active Campaign API
 *
 * @param {number} organizationId the organization ID in Active Campaign
 * @return {object} the result of the fetch from the Active Campaign API
 */
function getOrganizationCustomFields(organizationId) {
  const result = fetchFromActiveCampaign(
    "get",
    CREDENTIALS.activeCampaignBaseUrl +
      "accounts/" +
      organizationId +
      "/accountCustomFieldData"
  );

  Logger.log(
    "Organization Custom Fields Fetch Result: " + JSON.stringify(result)
  );

  return result;
}

/**
 * Fetches a deal by ID from Active Campaign API
 *
 * @param {number} dealId the ID of the deal in Active Campaign
 * @return {object} the result of the fetch from the Active Campaign API
 */
function getDeal(dealId) {
  const result = fetchFromActiveCampaign(
    "get",
    CREDENTIALS.activeCampaignBaseUrl + "deals/" + dealId
  );

  Logger.log("Deal Fetch Result: " + JSON.stringify(result));

  return result;
}

/**
 * Fetches a contact by ID from Active Campaign API
 *
 * @param {number} contactId the ID of the contact in Active Campaign
 * @return {object} the result of the fetch from the Active Campaign API
 */
function getContact(contactId) {
  const result = fetchFromActiveCampaign(
    "get",
    CREDENTIALS.activeCampaignBaseUrl + "contacts/" + contactId
  );

  Logger.log("Contact Fetch Result: " + JSON.stringify(result));

  return result;
}

/**
 * Fetches the deals for a contact from the Active Campaign API
 *
 * @param {number} contactId the contact ID in Active Campaign
 * @return {object} the result of the fetch from the Active Campaign API
 */
function getContactDeals(contactId) {
  const result = fetchFromActiveCampaign(
    "get",
    CREDENTIALS.activeCampaignBaseUrl + "contacts/" + contactId + "/deals"
  );

  Logger.log("Deals Fetch Result: " + JSON.stringify(result));

  return result;
}

/**
 * Fetches the user from the Active Campaign API
 *
 * @param {number} userId the user ID in Active Campaign
 * @return {object} the result of the fetch from the Active Campaign API
 */
function getDealUser(userId) {
  const result = fetchFromActiveCampaign(
    "get",
    CREDENTIALS.activeCampaignBaseUrl + "users/" + userId
  );

  Logger.log("User Fetch Result: " + JSON.stringify(result));

  return result;
}

/**
 * Fetches the deal custom fields from the Active Campaign API
 *
 * @param {number} dealId the deal ID in Active Campaign
 * @return {object} the result of the fetch from the Active Campaign API
 */
function getDealCustomFields(dealId) {
  const result = fetchFromActiveCampaign(
    "get",
    CREDENTIALS.activeCampaignBaseUrl +
      "deals/" +
      dealId +
      "/dealCustomFieldData"
  );

  Logger.log("Deal Custom Fields Fetch Result: " + JSON.stringify(result));

  return result;
}

/**
 * Fetches the deal notes from the Active Campaign API
 *
 * @param {number} dealId the deal ID in Active Campaign
 * @return {object} the result of the fetch from the Active Campaign API
 */
function getDealNotes(dealId) {
  const result = fetchFromActiveCampaign(
    "get",
    CREDENTIALS.activeCampaignBaseUrl + "deals/" + dealId + "/notes"
  );

  Logger.log("Deal Notes Fetch Result: " + JSON.stringify(result));

  return result;
}

/**
 * Fetches the contact custom fields from the Active Campaign API
 *
 * @param {number} contactId the contact ID in Active Campaign
 * @return {object} the result of the fetch from the Active Campaign API
 */
function getContactCustomFields(contactId) {
  const result = fetchFromActiveCampaign(
    "get",
    CREDENTIALS.activeCampaignBaseUrl + "contacts/" + contactId
  );

  Logger.log("Contact Custom Fields Fetch Result: " + JSON.stringify(result));

  // Strip everyting except the custom fields and accountContacts
  return _.pick(result, ["accountContacts", "fieldValues"]);
}

/**
 * Fetches the contact custom fields from the Active Campaign API
 *
 * @param {number} contactId the contact ID in Active Campaign
 * @return {object} the result of the fetch from the Active Campaign API
 */
function fetchFromActiveCampaign(method, url) {
  const options = {
    method: method,
    contentType: "application/json",
    headers: {
      "Api-Token": CREDENTIALS.activeCampaignApiToken,
    },
  };

  const result = UrlFetchApp.fetch(url, options);

  return JSON.parse(result.getContentText());
}

/**
 * Runs a PUT request to the Active Campaign API to
 * update the deal with the newly created job folders and sheets IDs
 *
 * @param {object} jobObj a huge object containing job related data
 * @return {object} the result of the request from the Active Campaign API
 */
function updateDealGoogleFolderIds(jobObj) {
  let data = {
    deal: {
      fields: [
        {
          customFieldId: CUSTOM_FIELD_ID_MAP.DEAL.JOB_FOLDER_URL,
          fieldValue: jobObj.jobFolderUrl,
        },
        {
          customFieldId: CUSTOM_FIELD_ID_MAP.DEAL.ESTIMATE_SHEET_ID,
          fieldValue: jobObj.estimateSpreadsheetId,
        },
        {
          customFieldId: CUSTOM_FIELD_ID_MAP.DEAL.GOOGLE_DRIVE_FOLDER_ID,
          fieldValue: jobObj.jobFolderId,
        },
      ],
    },
  };

  const result = updateActiveCampaign(
    "put",
    CREDENTIALS.activeCampaignBaseUrl + "deals/" + jobObj.deal.id,
    data
  );

  Logger.log(
    "Deal (ID: " +
      jobObj.deal.id +
      ") updated with newly created Google Drive folders."
  );

  Logger.log("Result of updated deal :" + result);
}

/**
 * Runs a PUT request to the Active Campaign API to
 * update the deal the newly created job number
 *
 * @param {number} dealId the id of the deal in Active Campaign
 * @param {string} dealTitle the name of the deal / job
 * @param {number} currentJobNumber the current job number
 */
function updateDealWithJobNumber(dealId, dealTitle, currentJobNumber) {
  let data = {
    deal: {
      title: currentJobNumber.toString() + " " + dealTitle,
      fields: [
        {
          customFieldId: CUSTOM_FIELD_ID_MAP.DEAL.JOB_NUMBER,
          fieldValue: currentJobNumber,
        },
      ],
    },
  };

  const result = updateActiveCampaign(
    "put",
    CREDENTIALS.activeCampaignBaseUrl + "deals/" + dealId,
    data
  );

  Logger.log(
    "Deal (ID: " + dealId + ") updated with new job number " + currentJobNumber
  );
  Logger.log("Result of updated deal :" + result);
}

/**
 * A helper function that run an Active Campaign API request
 *
 * @param {string} method a string representing the HTTP request method
 * @param {string} url The URL to hit
 * @param {object} payload the JSON object payload to use in the request
 * @return {number} the text result of the request
 */
function updateActiveCampaign(method, url, payload) {
  const options = {
    method: method,
    contentType: "application/json",
    headers: {
      "Api-Token": CREDENTIALS.activeCampaignApiToken,
    },
    payload: JSON.stringify(payload),
  };

  const result = UrlFetchApp.fetch(url, options);

  return JSON.parse(result.getContentText());
}

/**
 * Creates a new job folder in Google Drive estimates jobs folder
 *
 * @param {number} jobNumber the job number
 * @param {string} jobTitle the title of the job
 * @param {string} organizationName the name of the client company
 * @return {object} an instance of the Google Apps Script Folder Class
 */
function createNewJobFolder(jobNumber, jobTitle, organizationName) {
  const estimatesJobFolder = DriveApp.getFolderById(
    CREDENTIALS.estimatesJobFolder
  );

  let fullJobFolderName = jobNumber.toString().concat(" ", jobTitle);

  if (organizationName) {
    fullJobFolderName = fullJobFolderName.concat(" ", organizationName);
  }

  return estimatesJobFolder.createFolder(fullJobFolderName);
}

/**
 * A function that accepts a target folder and job object
 * and creates all the folders and files that are defined in the
 * Google Drive Structure Object for the "discovery" stage of the job
 *
 * @param {object} jobFolder the target job folder Google drive Folder instance
 * @param {object} jobObj an enormous object containing job data.
 */
function createDiscoveryJobFolders(jobFolder, jobObj) {
  for (const [folderName, folderContents] of Object.entries(
    DISCOVERY_JOB_FOLDER_STRUCTURE
  )) {
    let currFolder = jobFolder.createFolder(
      `${jobObj.jobNumber} ${folderName}`
    );
    let currFile;
    let newFile;

    folderContents.fileTemplates.forEach((file) => {
      currFileName = Object.entries(file)[0][0];
      currFileTemplateId = Object.entries(file)[0][1];

      Logger.log("Processing file with name:" + currFileName);
      Logger.log("Processing file with template ID:" + currFileTemplateId);

      // Get the current file by its ID
      currFile = DriveApp.getFileById(currFileTemplateId);

      // Copy the file into the current folder with the correct name
      newFile = currFile.makeCopy(
        `${jobObj.jobNumber} ${jobObj.deal.title} ${currFileName}`,
        currFolder
      );

      // Send it off to fill
      fillDocument(newFile, currFileName, jobObj);
    });
  }
}

/**
 * A helper that directs the filling of document
 * data to the correct function based on a files name
 *
 * @param {object} file an instance of the Google Apps Script File class
 * @param {string} fileName the name of the file to be filled
 * @param {object} jobObj an enourmous object containing job data
 */
function fillDocument(file, fileName, jobObj) {
  switch (fileName) {
    case "ESTIMATE":
      fillEstimateSheet(file, jobObj);
      break;
    case "COSTING":
      // code block
      break;
    default:
    // code block
  }
}

/**
 * Fills the Estimate file range with job data
 *
 * @param {object} file an instance of the Google Apps Script File class
 * @param {number} jobObj an enourmous object containing job data
 */
function fillEstimateSheet(file, jobObj) {
  Logger.log("Filling Estimate Sheet...");
  jobObj.estimateSpreadsheetId = file.getId();
  const estimateSpreadsheet = SpreadsheetApp.open(file);
  const informationSheet = estimateSpreadsheet.getSheetByName("Information");

  var values = [
    // Received By
    [jobObj.user.firstName.concat(" ", jobObj.user.lastName)],
    // Job Number
    [jobObj.jobNumber],
    // Job Name
    [jobObj.deal.title],
    // Date of Inquiry
    [Utilities.formatDate(new Date(jobObj.deal.cdate), "GMT-7", "MM/dd/yyyy")],
    // Client Contact
    [jobObj.contact["firstName"].concat(" ", jobObj.contact["lastName"])],
    // Client Company
    [jobObj.organization?.account?.name],
    // Client Office Phone
    [jobObj.organization?.account?.phone],
    // Client Cell Phone
    [jobObj.contact["phone"]],
    // Client Email
    [jobObj.contact["email"]],
    // End User
    [jobObj.deal.endUser],
    // Designer ??
    ["Designer"],
    // Delivery Location
    [jobObj.deal.deliveryLocation],
    // Delivery Date
    [
      jobObj.deal.deliveryDate
        ? Utilities.formatDate(
            new Date(jobObj.deal.deliveryDate),
            "GMT-7",
            "MM/dd/yyyy"
          )
        : "",
    ],
    // Install Date
    [
      jobObj.deal.installDate
        ? Utilities.formatDate(
            new Date(jobObj.deal.installDate),
            "GMT-7",
            "MM/dd/yyyy"
          )
        : "",
    ],
    // Strike Date
    [
      jobObj.deal.strikeDate
        ? Utilities.formatDate(
            new Date(jobObj.deal.strikeDate),
            "GMT-7",
            "MM/dd/yyyy"
          )
        : "",
    ],
    // Budget / Value
    [jobObj.deal.formattedValue],
  ];

  const range = informationSheet.getRange("B3:B18");
  range.setValues(values);
}

/**
 * Checks to see if an object is empty
 *
 * @param {object} obj an object
 * @return {boolean}
 */
function isEmpty(obj) {
  return Object.keys(obj).length === 0;
}
