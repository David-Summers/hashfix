(function () {
  // Associates the handler name with a function in classic Windows hosts
  if (window.Office && Office.actions && Office.actions.associate) {
    Office.actions.associate("onMessageSendHandler", window.onMessageSendHandler || function (event) {
      event.completed({ allowEvent: true });
    });
  }
})();