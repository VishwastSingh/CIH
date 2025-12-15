Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("forward-button").onclick = forwardEmail;
  }
});

function forwardEmail() {
  const messageArea = document.getElementById("message-area");
  const forwardButton = document.getElementById("forward-button");
  
  messageArea.textContent = "Forwarding email...";
  messageArea.style.color = "blue";
  forwardButton.disabled = true;
  
  const item = Office.context.mailbox.item;
  const forwardTo = "lilly_clinical_trials_test@lilly.com";
  
  // Try EWS first
  tryEWSForward(item, forwardTo, messageArea, forwardButton);
}

function tryEWSForward(item, forwardTo, messageArea, forwardButton) {
  const itemId = item.itemId;
  
  // Convert item ID to EWS format
  const ewsId = Office.context.mailbox.convertToEwsId(
    itemId,
    Office.MailboxEnums.RestVersion.v2_0
  );
  
  const ewsRequest = `<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
               xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"
               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
  <soap:Header>
    <t:RequestServerVersion Version="Exchange2013" />
  </soap:Header>
  <soap:Body>
    <m:CreateItem MessageDisposition="SendAndSaveCopy">
      <m:Items>
        <t:ForwardItem>
          <t:ToRecipients>
            <t:Mailbox>
              <t:EmailAddress>${forwardTo}</t:EmailAddress>
            </t:Mailbox>
          </t:ToRecipients>
          <t:ReferenceItemId Id="${ewsId}" />
          <t:NewBodyContent BodyType="Text">Forwarded via Clinical Inquiry Hub Forwarder</t:NewBodyContent>
        </t:ForwardItem>
      </m:Items>
    </m:CreateItem>
  </soap:Body>
</soap:Envelope>`;

  Office.context.mailbox.makeEwsRequestAsync(ewsRequest, function(result) {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      messageArea.textContent = "✓ Email forwarded automatically to " + forwardTo;
      messageArea.style.color = "green";
      forwardButton.disabled = false;
    } else {
      console.error('EWS failed:', result.error);
      console.error('Error code:', result.error.code);
      console.error('Error message:', result.error.message);
      
      // Fall back to compose window approach
      messageArea.textContent = "EWS not available. Opening compose window...";
      messageArea.style.color = "orange";
      
      setTimeout(() => {
        openComposeWindow(item, forwardTo, messageArea, forwardButton);
      }, 1000);
    }
  });
}

// Function to sanitize HTML and remove @mentions that trigger auto-recipient addition
function sanitizeHtmlForForwarding(html) {
  // Remove Outlook @mention spans/divs that trigger auto-recipient addition
  // These typically have data-* attributes or specific classes
  html = html.replace(/<a[^>]*data-auth[^>]*>(@[^<]*)<\/a>/gi, '$1');
  html = html.replace(/<span[^>]*data-mention[^>]*>([^<]*)<\/span>/gi, '$1');
  
  // Replace @mentions with [at]mentions to prevent auto-processing
  // This regex finds @word patterns that are likely mentions
  html = html.replace(/@(\w+)/g, '&#64;$1');
  
  return html;
}

function openComposeWindow(item, forwardTo, messageArea, forwardButton) {
  const subject = "FW: " + item.subject;
  
  // Get body as PLAIN TEXT to avoid @mention HTML tags
  item.body.getAsync(Office.CoercionType.Text, function(bodyResult) {
    if (bodyResult.status === Office.AsyncResultStatus.Succeeded) {
      let body = bodyResult.value;
      const from = item.from;
      const dateTimeCreated = item.dateTimeCreated;
      
      // Format as plain text (no HTML) to prevent @mention processing
      const forwardHeader = `\n\n---------- Forwarded message ---------\n` +
                          `From: ${from.displayName} <${from.emailAddress}>\n` +
                          `Date: ${dateTimeCreated}\n` +
                          `Subject: ${item.subject}\n\n`;
      
      const fullBody = forwardHeader + body;
      
      const forwardMessage = {
        toRecipients: [forwardTo],
        subject: subject,
        body: fullBody  // Using 'body' instead of 'htmlBody' sends as plain text
      };
      
      Office.context.mailbox.displayNewMessageForm(forwardMessage);
      
      messageArea.innerHTML = "✓ Forward ready. <strong>Please click Send</strong> to complete.<br><small>(Automatic send not available in this environment)</small>";
      messageArea.style.color = "green";
      forwardButton.disabled = false;
    } else {
      messageArea.textContent = "✗ Error preparing forward";
      messageArea.style.color = "red";
      forwardButton.disabled = false;
    }
  });
}