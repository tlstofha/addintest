/* global Office */
Office.onReady();

function block(event, msg) {
  event.completed({ allowEvent: false, errorMessage: msg });
}
function allow(event) {
  event.completed({ allowEvent: true });
}
// 10초(10,000ms) 후 허용
function allowAfterDelay(event, ms) {
  setTimeout(function () { allow(event); }, ms);
}

function onMessageSendHandler(event) {
  Office.context.mailbox.item.subject.getAsync(function (r) {
    if (r.status !== Office.AsyncResultStatus.Succeeded) {
      return block(event, "제목 확인 중 오류가 발생했습니다.");
    }
    const subject = (r.value || "").trim();
    if (!subject) {
      return block(event, "제목이 비었습니다. 제목을 입력하세요.");
    }
    // ▶ OnSend 프로세스 10초 유지
    return allowAfterDelay(event, 10000);
  });
}

function onAppointmentSendHandler(event) {
  Office.context.mailbox.item.subject.getAsync(function (r) {
    if (r.status !== Office.AsyncResultStatus.Succeeded) {
      return block(event, "약속 제목 확인 중 오류가 발생했습니다.");
    }
    const subject = (r.value || "").trim();
    if (!subject) {
      return block(event, "약속 제목이 비었습니다. 제목을 입력하세요.");
    }
    // ▶ OnSend 프로세스 10초 유지
    return allowAfterDelay(event, 10000);
  });
}

function onMessageComposeHandler(event) { event.completed(); }
function onAppointmentComposeHandler(event) { event.completed(); }

Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
Office.actions.associate("onAppointmentSendHandler", onAppointmentSendHandler);
Office.actions.associate("onMessageComposeHandler", onMessageComposeHandler);
Office.actions.associate("onAppointmentComposeHandler", onAppointmentComposeHandler);
