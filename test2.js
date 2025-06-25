function onMessageSendHandler(event) {
  console.log("OnSend triggered");
  const recipients = Office.context.mailbox.item.to;
  const externalRecipients = recipients.filter(r => !r.emailAddress.endsWith("@contoso.com"));

  if (externalRecipients.length > 0) {
    Office.context.mailbox.item.subject.setAsync("[External] " + Office.context.mailbox.item.subject.getAsync, () => {
      event.completed({ allowEvent: true });
    });
  } else {
    event.completed({ allowEvent: true });
  }
}

Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
