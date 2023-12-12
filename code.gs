var SCOPES = ['https://www.googleapis.com/auth/gmail.modify','https://www.google.com/m8/feeds'];

// function onOpen(e) {
//   var ui = SpreadsheetApp.getUi().createAddonMenu()
//       .addItem('Mail Merge', 'showSidebar')
//       .addToUi();
// }
function onOpen(e) {
  var ui = SpreadsheetApp.getUi().createMenu('Mail Merge')
      .addItem('Mail Merge', 'showSidebar')
      .addToUi();
      
}

function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar')
      .setTitle('Mail Merge');
  SpreadsheetApp.getUi().showSidebar(html);
}

function onInstall(e) {
  resetAuthorization(); // Reset authorization properties
  onOpen(e);
}

function authorizeAndShowSidebar() {
  var service = getOAuthService();

  // Check if the user has already authorized
  var userProperties = PropertiesService.getUserProperties();
  var hasAuthorized = userProperties.getProperty('hasAuthorized');

  if (service.hasAccess() || hasAuthorized === 'true') {
    // If already authorized, return the access token
    return service.getAccessToken();
  } else {
    // If not authorized, display the authorization link
    var authorizationUrl = service.getAuthorizationUrl();
    var html = '<a href="' + authorizationUrl + '" target="_blank">Authorize</a>. ' +
               'Reopen the sidebar when the authorization is complete.';

    var page = HtmlService.createHtmlOutput(html);

    // Show the sidebar with the authorization link
    SpreadsheetApp.getUi().showSidebar(page);

    // Return null to indicate that authorization is not yet complete
    //return null;
  }
}




// var AUTH_TEMPLATE = HtmlService.createTemplate(
//   '<a href="<?= authorizationUrl ?>" target="_blank">Authorize</a>. ' +
//   'Reopen the sidebar when the authorization is complete.'
// );

function authorize() {
  var service = getOAuthService();
  if (service.hasAccess()) {
    return service.getAccessToken();
  } else {
    AUTH_TEMPLATE.authorizationUrl = service.getAuthorizationUrl();
    Logger.log(service.getAuthorizationUrl())
    var page = AUTH_TEMPLATE.evaluate();
    SpreadsheetApp.getUi().showSidebar(page);
  }
}


function getOAuthService() {
  var scriptProps = PropertiesService.getScriptProperties();
  var clientId = '323888304851-h5rq2e64qfstafem88djtkf8coqe1pvu.apps.googleusercontent.com';
  var clientSecret = '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';
  var service = OAuth2.createService('Gmail')
      .setAuthorizationBaseUrl('https://accounts.google.com/o/oauth2/auth')
      .setTokenUrl('https://accounts.google.com/o/oauth2/token')
      .setClientId(clientId)
      .setClientSecret(clientSecret)
      .setScope(SCOPES)
      .setCallbackFunction('authCallback')
      .setPropertyStore(scriptProps)
      .setCache(CacheService.getUserCache())
      .setParam('login_hint', Session.getActiveUser().getEmail())
      .setParam('authuser', Session.getActiveUser().getEmail())
      .setParam('approval_prompt', 'auto')
      .setParam('access_type', 'offline')
      .setParam('response_type', 'code')
      .setParam('authMode', 'LIMITED') // Add this line to set the authMode parameter to LIMITED
      .setParam('hl', 'en');


  return service;
}



function authCallback(request) {
  var service = getOAuthService();
  var authorized = service.handleCallback(request);

  if (authorized) {
    // Set the flag to indicate that authorization is complete
    PropertiesService.getUserProperties().setProperty('hasAuthorized', 'true');

    return HtmlService.createHtmlOutput('Authorization successful. You can close this tab and return to the Mail Merge sidebar.');
  } else {
    return HtmlService.createHtmlOutput('Authorization failed. Please try again.');
  }
}



function sendEmailFromDraftWithProgress(draftId, statusColumn) {
  // Authorize and get the access token
  var accessToken = authorizeAndShowSidebar();
  if (!accessToken) {
    // Authorization failed, return immediately
    return;
  }

  var draft = GmailApp.getDrafts().find(function(draft) {
    return draft.getId() === draftId;
  });

  if (draft) {
    var sheet = SpreadsheetApp.getActiveSheet();
    var selectedRange = sheet.getActiveRange();
    var selectedValues = selectedRange.getValues();
  
    // Get the draft message and its properties
    var message = draft.getMessage();
    var cc = message.getCc();
    var bcc = message.getBcc();
    var subject = message.getSubject();
    var body = message.getBody();
    
    for (var i = 0; i < selectedValues.length; i++) {
      var recipientEmail = selectedValues[i][0];
      var recipientName = getRecipientName(recipientEmail); // fetch recipient name using email address

      // Replace recipient name in the draft body
      var bodyWithRecipientName = body.replace(/{{recipient_name}}/g, recipientName);

      try {
        // Construct the raw email message
        var messageBody = [          'To: ' + recipientEmail,          'Cc: ' + cc,          'Bcc: ' + bcc,          'Subject: ' + subject,          'Content-Type: text/html; charset=UTF-8',          '',          bodyWithRecipientName        ].join('\r\n');
        var encodedMessage = Utilities.base64EncodeWebSafe(messageBody);

        // Send the email
        var headers = {
          'Authorization': 'Bearer ' + accessToken,
          'Content-Type': 'application/json'
        };
        var payload = {
          'raw': encodedMessage
        };
        var options = {
          'method': 'post',
          'headers': headers,
          'payload': JSON.stringify(payload)
        };
        var response = UrlFetchApp.fetch('https://www.googleapis.com/gmail/v1/users/me/messages/send', options);
        Utilities.sleep(10000); // add a 10-second delay between sending each email

        // Update the spreadsheet with "Sent" status
        var row = selectedRange.getRowIndex() + i;
        var column = selectedRange.getLastColumn() + 1;
        sheet.getRange(row, statusColumn).setValue("Sent");
        sheet.getRange(row, statusColumn).setBackground('green');
      } catch (error) {
        // Log any errors that occur while sending the email
        // Update the spreadsheet with "Sent" status
        var row = selectedRange.getRowIndex() + i;
        sheet.getRange(row, statusColumn).setValue("Failed: " + error.message);
        sheet.getRange(row, statusColumn).setBackground('red');
        Logger.log('Failed to send email to ' + recipientEmail + ': ' + error.message);
      }
    }

  }
}

function getRecipientName(recipientEmail) {
  var service = getOAuthService();
  var url = "https://people.googleapis.com/v1/people/me/connections?personFields=names,emailAddresses&pageSize=1000&sortOrder=LAST_MODIFIED_DESCENDING";
  var headers = {
    'Authorization': 'Bearer ' + service.getAccessToken()
  };
  var options = {
    'headers': headers,
    'method': 'get',
    'muteHttpExceptions': true // add this option to mute HTTP exceptions
  };
  var response = UrlFetchApp.fetch(url, options);
  var result = JSON.parse(response.getContentText());
  Logger.log(result)
  if (result && result.connections && result.connections.length > 0) {
    var connections = result.connections;
    for (var i = 0; i < connections.length; i++) {
      var contact = connections[i];
       Logger.log(contact)
      if (contact.emailAddresses && contact.emailAddresses.length > 0) {
        var emailAddress = contact.emailAddresses[0].value;
        if (emailAddress === recipientEmail && contact.names && contact.names.length > 0) {
          var fullName = contact.names[0].displayName;
          if (fullName) {
            Logger.log(fullName)
            return fullName;
          }
        }
      }
    }
  }
  // if no contact found, return only the email address
  return recipientEmail;
}


function resetAuthorization() {
  PropertiesService.getUserProperties().deleteProperty('hasAuthorized');
  var service = getOAuthService();
  service.reset();
}

// function resetAuthorization() {
//   var service = getOAuthService();
//   service.reset();
// }


function getDrafts() {
  var drafts = GmailApp.getDrafts();
  var draftList = [];
  for (var i = 0; i < drafts.length; i++) {
    var draft = drafts[i];
    var message = draft.getMessage();
    var subject = message.getSubject();
    var id = draft.getId();
    draftList.push({subject: subject, id: id});
  }
  return draftList;
}




