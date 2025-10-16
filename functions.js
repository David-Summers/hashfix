/* HashFix Outlook Add-in: On-send body transformer (single-pass)
   Rule: Replace occurrences of "#3" with "<sup>#</sup>3" ONLY when not already superscripted.
   Implementation: protect existing "<sup>#</sup>3" with a placeholder, replace "#3", then restore. */

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

function transformOnce(html) {
  const TOKEN = "__HASHFIX_SUPHASH3__";
  // 1) Protect already-correct superscripts
  let protectedHtml = html.split("<sup>#</sup>3").join(TOKEN);
  // 2) Single-pass replacement for raw '#3'
  protectedHtml = protectedHtml.split("#3").join("<sup>#</sup>3");
  // 3) Restore protected ones
  return protectedHtml.split(TOKEN).join("<sup>#</sup>3");
}

async function onMessageSend(event) {
  try {
    const html = await getBodyHtml();
    const fixed = transformOnce(html);
    if (fixed !== html) {
      await setBodyHtml(fixed);
    }
    event.completed({ allowEvent: true }); // proceed with send
  } catch (e) {
    // Fail-open for POC
    event.completed({ allowEvent: true });
  }
}

window.onMessageSend = onMessageSend;
