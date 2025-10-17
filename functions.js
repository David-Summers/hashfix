
function onMessageSendHandler(event) {
  console.log("HashFix: onMessageSendHandler triggered");

  const item = Office.context.mailbox.item;

  item.body.getAsync(Office.CoercionType.Text, function(result) {
    if (result.status === Office.AsyncResultStatus.Failed) {
      console.error("HashFix: Failed to get body", result.error);
      event.completed({ allowEvent: true });
      return;
    }

    const originalText = result.value;
    const fixedText = fixHashtagsSimple(originalText);

    if (fixedText !== originalText) {
      item.body.setAsync(fixedText, { coercionType: Office.CoercionType.Text }, function(setResult) {
        if (setResult.status === Office.AsyncResultStatus.Succeeded) {
          console.log("HashFix: Hashtags corrected");
        } else {
          console.error("HashFix: Failed to set body", setResult.error);
        }
        event.completed({ allowEvent: true });
      });
    } else {
      console.log("HashFix: No changes needed");
      event.completed({ allowEvent: true });
    }
  });
}

function fixHashtagsSimple(text) {
  return text.replace(/#\s+(\w+)/g, '#$1');
}

if (window.Office && Office.actions && Office.actions.associate) {
  Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
}

window.onMessageSendHandler = onMessageSendHandler;
