/* HashFix Outlook Add-in: On-send body transformer for new Outlook
   Rule: Replace "#3" with "<sup>#</sup>3" (idempotent). */

Office.initialize = () => {};

function getBodyHtml() {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.body.getAsync(Office.CoercionType.Html, result => {
      if (result.status === Office.AsyncResultStatus.Succeeded) resolve(result.value || "");
      else reject(result.error);
    });
  });
}

function setBodyHtml(html) {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.body.setAsync(html, { coercionType: Office.CoercionType.Html }, res => {
      if (res.status === Office.AsyncResultStatus.Succeeded) resolve();
      else reject(res.error);
    });
  });
}

function transform(html) {
  // keep idempotent
  const undone = html.split('<sup>#</sup>3').join('#3');
  return undone.split('#3').join('<sup>#</sup>3');
}

async function onMessageSend(event) {
  try {
    const html = await getBodyHtml();
    const fixed = transform(html);
    if (fixed !== html) await setBodyHtml(fixed);
    event.completed({ allowEvent: true }); // proceed with send
  } catch (e) {
    event.completed({ allowEvent: true }); // fail-open for POC
  }
}

window.onMessageSend = onMessageSend;
