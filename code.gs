version 13

var SCOPES = ['https://www.googleapis.com/auth/gmail.modify','https://www.google.com/m8/feeds'];



// var AUTH_TEMPLATE = HtmlService.createTemplate(
//   '<a href="<?= authorizationUrl ?>" target="_blank">Authorize</a>. ' +
//   'Reopen the sidebar when the authorization is complete.'
// );

// function authorize() {
//   var AUTH_TEMPLATE = HtmlService.createTemplate(
//     '<a href="<?= authorizationUrl ?>" target="_blank">Authorize</a>. ' +
//     'Reopen the sidebar when the authorization is complete.'
//   );

//   var service = getOAuthService();
//   if (service.hasAccess()) {
//     return service.getAccessToken();
//   } else {
//     AUTH_TEMPLATE.authorizationUrl = service.getAuthorizationUrl();
//     Logger.log(service.getAuthorizationUrl());
//     var page = AUTH_TEMPLATE.evaluate();
//     SpreadsheetApp.getUi().showSidebar(page);
//   }
// }



// function authorizeUser() {
//   var authorizationUrl = getAuthorizationUrl();
//   var htmlOutput = HtmlService.createHtmlOutput('<a href="' + authorizationUrl + '" target="_blank">Authorize</a>. Reopen the sidebar when the authorization is complete.');
//   SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Authorization');
// }


 function authorize() {
  var service = getOAuthService();
  var currentUserEmail = Session.getActiveUser().getEmail();
  // Ensure that the service has access and the access token is not empty
  // Check if the stored user email matches the current user
  if (service.hasAccess() ) {
    return service.getAccessToken();
  }  else {
    var authorizationUrl = getAuthorizationUrl();
    var htmlOutput = HtmlService.createHtmlOutput('<a href="' + authorizationUrl + '" target="_blank">Authorize</a>. Reopnnnnnnnnnnnnnn the sidebar when the authorization is complete.');
    SpreadsheetApp.getUi().showSidebar(htmlOutput);
  //  return null; // Return null to indicate that authorization is still required
  }
}




function getAuthorizationUrl() {
  var service = getOAuthService();
  return service.getAuthorizationUrl();
}




function getOAuthService() {
  var scriptProps = PropertiesService.getScriptProperties();
   
  var currentUserEmail = Session.getActiveUser().getEmail();
  
  var clientId = '323888304851-h5rq2e64qfstafem88djtkf8coqe1pvu.apps.googleusercontent.com';
  var clientSecret = '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';

  //scriptProps.setProperty('userEmail', currentUserEmail); 
  
  var service = OAuth2.createService('Gmail' + currentUserEmail)
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
      
  //scriptProps.setProperty('userEmail', currentUserEmail);

  return service;
}



function authCallback(request) {
  var service = getOAuthService();
  var authorized = service.handleCallback(request);
  if (authorized) {
    Logger.log('Authorization successful. Access token: ' + service.getAccessToken());
    return HtmlService.createHtmlOutput('Authorization successful. You can close this tab and return to the Mail Merge sidebar.');
  } else {
    Logger.log('Authorization failed. Error: ' + service.getLastError());
    return HtmlService.createHtmlOutput('Authorization failed. Please try again.');
  }
}


function displayAuthorizationLink() {
  console.log('Displaying authorization link...');
  var authorizationUrl = getAuthorizationUrl();
  console.log('Authorization URL:', authorizationUrl);
  var htmlOutput = HtmlService.createHtmlOutput('<a href="' + authorizationUrl + '" target="_blank">Authorize</a>. Reopen the sidebar when the authorization is complete.');
  SpreadsheetApp.getUi().showSidebar(htmlOutput);
}


function sendEmailFromDraftWithProgress(draftId, statusColumn) {
  // Authorize and get the access token
  var accessToken = authorize();
  Logger.log(accessToken)
  if (!accessToken) {
    var authorizationUrl = getAuthorizationUrl();
    var htmlOutput = HtmlService.createHtmlOutput('<a href="' + authorizationUrl + '" target="_blank">Authorize</a>. Reopnnnnnnnnnnnnnn the sidebar when the authorization is complete.');
    SpreadsheetApp.getUi().showSidebar(htmlOutput);
    // Authorization failed, return immediately
    throw new Error('Authorization required. Please authorize before sending emails.');

 
   // return;
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
    var row = selectedRange.getRowIndex() + i;
    var column = selectedRange.getLastColumn() + 1;

    if (error.message.indexOf('401') !== -1) {
      // Handle 401 Unauthorized error
      var authorizationUrl = getAuthorizationUrl();
      var htmlOutput = HtmlService.createHtmlOutput('<a href="' + authorizationUrl + '" target="_blank">Authorize</a>. Reopen the sidebar when the authorization is complete.');
      SpreadsheetApp.getUi().showSidebar(htmlOutput);

      sheet.getRange(row, statusColumn).setValue("Authorization required");
      sheet.getRange(row, statusColumn).setBackground('yellow');
      Logger.log('Authorization required. Please authorize before sending emails.');
    } else {
      // Handle other errors
      sheet.getRange(row, statusColumn).setValue("Failed: " + error.message);
      sheet.getRange(row, statusColumn).setBackground('red');
      Logger.log('Failed to send email to ' + recipientEmail + ': ' + error.message);
    }
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
  var service = getOAuthService();
  service.reset();
}


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


// function onOpen(e) {
//   var ui = SpreadsheetApp.getUi().createMenu('Mail Merge')
//       .addItem('Mail Merge', 'showSidebar')
//       .addToUi();
    
// }

function onOpen(e) {
  var ui = SpreadsheetApp.getUi().createMenu('Mail Merge')
    .addItem('Mail Merge', 'showSidebar')
    // .addItem('Authorize', 'authorizeUser')
    .addToUi();
}




// function showSidebar() {
//   authorize();  // Call authorize to display the authorization link
//   var html = HtmlService.createHtmlOutputFromFile('Sidebar').setTitle('Mail Merge');
//   SpreadsheetApp.getUi().showSidebar(html);
// }


function showSidebar() {
  var accessToken = authorize();

  if (accessToken) {
    // User is already authorized, proceed with your existing sidebar content
    var html = HtmlService.createHtmlOutputFromFile('Sidebar').setTitle('Mail Merge');
    SpreadsheetApp.getUi().showSidebar(html);
  } else {
    // User needs to authorize, show the authorization link
    var authorizationUrl = getAuthorizationUrl();
    var htmlOutput = HtmlService.createHtmlOutput('<a href="' + authorizationUrl + '" target="_blank">Authorize</a>. Authorization is required for this add-on.').setHeight(200).setWidth(300);
    SpreadsheetApp.getUi().showSidebar(htmlOutput);
  }
}

function onInstall(e) {
  resetAuthorization()
  onOpen(e);
}

