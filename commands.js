Office.onReady(() => {
  // Commands.js is ready
});

// This function is called when the "Perform an action" button is clicked
function action(event) {
  const item = Office.context.mailbox.item;
  const forwardTo = "lilly_clinical_trials_test@lilly.com";
  
  // Show a notification that forwarding is in progress
  Office.context.mailbox.item.notificationMessages.addAsync("forward-progress", {
    type: "informationalMessage",
    message: "Forwarding email to Clinical Inquiry Hub...",
    icon: "icon-16",
    persistent: false
  });
  
  // Try EWS forwarding first (automatic, no compose window)
  tryEWSForward(item, forwardTo, event);
}

function tryEWSForward(item, forwardTo, event) {
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
      // Success! Email forwarded automatically
      Office.context.mailbox.item.notificationMessages.replaceAsync("forward-progress", {
        type: "informationalMessage",
        message: "✓ Email forwarded to Clinical Inquiry Hub",
        icon: "icon-16",
        persistent: true
      });
      event.completed();
    } else {
      console.error('EWS failed:', result.error);
      
      // Fall back to compose window approach
      Office.context.mailbox.item.notificationMessages.replaceAsync("forward-progress", {
        type: "informationalMessage",
        message: "Opening compose window - please click Send to complete",
        icon: "icon-16",
        persistent: false
      });
      
      // Add delay before opening compose window and completing event
      setTimeout(() => {
        openComposeWindowDirect(item, forwardTo, event);
      }, 500);
    }
  });
}

// Function to sanitize HTML and remove @mentions - DEPRECATED, now using plain text
function sanitizeHtmlForForwarding(html) {
  html = html.replace(/<a[^>]*data-auth[^>]*>(@[^<]*)<\/a>/gi, '$1');
  html = html.replace(/<span[^>]*data-mention[^>]*>([^<]*)<\/span>/gi, '$1');
  html = html.replace(/@(\w+)/g, '&#64;$1');
  return html;
}

function openComposeWindowDirect(item, forwardTo, event) {
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
      
      try {
        Office.context.mailbox.displayNewMessageForm(forwardMessage);
        // Wait a bit to ensure the window opens before completing
        setTimeout(() => {
          event.completed();
        }, 1000);
      } catch (error) {
        console.error('Error opening compose window:', error);
        Office.context.mailbox.item.notificationMessages.addAsync("forward-error", {
          type: "errorMessage",
          message: "Could not open compose window. Please use 'Show Task Pane' instead.",
          icon: "icon-16",
          persistent: true
        });
        event.completed();
      }
    } else {
      Office.context.mailbox.item.notificationMessages.addAsync("forward-error", {
        type: "errorMessage",
        message: "Could not prepare forward. Please use 'Show Task Pane' instead.",
        icon: "icon-16",
        persistent: true
      });
      event.completed();
    }
  });
}

// Make sure the function is globally available for Office
if (typeof global !== "undefined") {
  global.action = action;
}