Office.onReady();

const HOLD_MS = 10000; // 무조건 10초 지연

function allow(event) {
  event.completed({ allowEvent: true });
}
function allowAfterDelay(event, ms = HOLD_MS) {
  setTimeout(() => allow(event), ms);
}

function onMessageSendHandler(event) {
  // 제목 체크 없이 무조건 10초 지연 후 전송 허용
  allowAfterDelay(event);
}

function onAppointmentSendHandler(event) {
  // 제목 체크 없이 무조건 10초 지연 후 전송 허용
  allowAfterDelay(event);
}

function onMessageComposeHandler(event) { event.completed(); }
function onAppointmentComposeHandler(event) { event.completed(); }

Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
Office.actions.associate("onAppointmentSendHandler", onAppointmentSendHandler);
Office.actions.associate("onMessageComposeHandler", onMessageComposeHandler);
Office.actions.associate("onAppointmentComposeHandler", onAppointmentComposeHandler);
