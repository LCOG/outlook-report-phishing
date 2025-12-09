var sendMail = function() {
            var request = '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">'+
            '  <soap:Body>'+
            '    <m:CreateItem MessageDisposition="SendAndSaveCopy">'+
            '      <m:SavedItemFolderId><t:DistinguishedFolderId Id="sentitems" /></m:SavedItemFolderId>'+
            '      <m:Items>'+
            '        <t:Message>'+
            '          <t:Subject>Hello, Outlook!</t:Subject>'+
            '          <t:Body BodyType="HTML">This message was sent from ' + Office.context.mailbox.diagnostics.hostName + ', version ' + Office.context.mailbox.diagnostics.hostVersion + '! <br/>Contents:' + document.getElementById('emailDetails').innerHTML + '</t:Body>'+
            '          <t:ToRecipients>'+
            '            <t:Mailbox><t:EmailAddress>' + Office.context.mailbox.userProfile.emailAddress + '</t:EmailAddress></t:Mailbox>'+
            '          </t:ToRecipients>'+
            '        </t:Message>'+
            '      </m:Items>'+
            '    </m:CreateItem>'+
            '  </soap:Body>'+
            '</soap:Envelope>';

            Office.context.mailbox.makeEwsRequestAsync(request, function (asyncResult) {
                console.log(request);
                if (asyncResult.status == "failed") {
                    console.log("Action failed with error: " + asyncResult.error.message);
                }
                else {
                    console.log(asyncResult);
                    console.log("Message sent!");
                }
            });
            test();
        }
        
        var test = function() {
            var req = '<?xml version="1.0" encoding="utf-8"?>'+
'<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"'+
'  xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">'+
'  <soap:Body>'+
'    <CreateItem MessageDisposition="SendAndSaveCopy" xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">'+
'      <SavedItemFolderId>'+
'        <t:DistinguishedFolderId Id="drafts" />'+
'      </SavedItemFolderId>'+
'      <Items>'+
'        <t:Message>'+
'          <t:ItemClass>IPM.Note</t:ItemClass>'+
'          <t:Subject>Project Action</t:Subject>'+
'          <t:Body BodyType="Text">Priority - Update specification</t:Body>'+
'          <t:ToRecipients>'+
'            <t:Mailbox>'+
'              <t:EmailAddress>tshireman@lcog-or.gov</t:EmailAddress>'+
'            </t:Mailbox>'+
'          </t:ToRecipients>'+
'          <t:IsRead>false</t:IsRead>'+
'        </t:Message>'+
'      </Items>'+
'    </CreateItem>'+
'  </soap:Body>'+
'</soap:Envelope>';
Office.context.mailbox.makeEwsRequestAsync(req, function (asyncResult) {
                console.log(request);
                if (asyncResult.status == "failed") {
                    console.log("Action failed with error: " + asyncResult.error.message);
                }
                else {
                    console.log(asyncResult);
                    console.log("Message sent!");
                }
            });
        }
Office.initialize = function (reason) {
  // Check if the mailbox item is available
  var item = Office.context.mailbox.item;

  // Get the subject of the email
  var subject = item.subject;
  document.getElementById('emailSubject').innerText = subject;
  document.getElementById('thesub').value = subject;

  // Get the sender's email address
  var senderEmail = item.from.emailAddress;
  document.getElementById('senderEmail').innerText = senderEmail;
  document.getElementById('thesent').value = senderEmail;

  // Get the recipient's email addresses
  var recipientEmails = item.to.map(function (recipient) {
    return recipient.emailAddress;
  }).join(', ');
  document.getElementById('recipientEmails').innerText = recipientEmails;

    var actualRecipient = Office.context.mailbox.userProfile.emailAddress;
    document.getElementById('actualRecipient').innerText = actualRecipient;

  document.getElementById('therec').value = actualRecipient;


  // Get the text body of the email
  item.body.getAsync(Office.CoercionType.Text, function (result) {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      document.getElementById('emailBodyText').innerText = result.value;
      document.getElementById('thetext').value = result.value;
    } else {
      console.error('Error retrieving email text body:', result.error.message);
    }
  });

  // Get the HTML body of the email
  item.body.getAsync(Office.CoercionType.Html, function (result) {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      document.getElementById('emailBodyHtml').innerHTML = result.value;
     // document.getElementById('thehtml').value =  result.value;

      function escapeHtmlUsingDOM(text) {
            var tempDiv = document.createElement('div');
            tempDiv.textContent = text;
            return tempDiv.innerHTML;
        }
        // Safely encode the HTML content
        var encodedHtml = escapeHtmlUsingDOM(result.value);
        // Set the encoded HTML as the value of the input field
        document.getElementById('thehtml').value = encodedHtml;
    } else {
      console.error('Error retrieving email HTML body:', result.error.message);
    }
  });

  // Get the names of the attachments
  var attachmentNames = item.attachments.map(function (attachment) {
    return attachment.name;
  }).join(', ');
  document.getElementById('attachments').innerText = attachmentNames;
  document.getElementById('theatt').value = attachmentNames;

  // Get the headers of the email
  item.getAllInternetHeadersAsync(function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      var headers = asyncResult.value;
      console.log(headers);
      // Do something with headers
      document.getElementById('emailHeaders').innerText = headers;
      document.getElementById('theheaders').value = headers;
    } else {
      console.error('Error retrieving email headers:', asyncResult.error.message);
    }
  });


};

