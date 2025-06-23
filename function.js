Office.initialize = function () {
  Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
};

function onMessageSendHandler(event) {
  console.log("Message send triggered.");
  event.completed({ allowEvent: true });
}
