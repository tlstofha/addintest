/* global Office */
Office.onReady();

// 공통 헬퍼
function block(event, msg) {
  event.completed({ allowEvent: false, errorMessage: msg });
}
function allow(event) {
  event.completed({ allowEvent: true });
}

// OnSend - 메일 전송 시 제목 체크
function onMessageSendHandler(event) {
  Office.context.mailbox.item.subject.getAsync(function (r) {
    if (r.status !== Office.AsyncResultStatus.Succeeded) {
      return block(event, "제목 확인 중 오류가 발생했습니다.");
    }
    const subject = (r.value || "").trim();
    if (!subject) {
      return block(event, "제목이 비었습니다. 제목을 입력하세요.");
    }
    allow(event);
  });
}

// OnSend - 약속 전송 시 제목 체크
function onAppointmentSendHandler(event) {
  Office.context.mailbox.item.subject.getAsync(function (r) {
    if (r.status !== Office.AsyncResultStatus.Succeeded) {
      return block(event, "약속 제목 확인 중 오류가 발생했습니다.");
    }
    const subject = (r.value || "").trim();
    if (!subject) {
      return block(event, "약속 제목이 비었습니다. 제목을 입력하세요.");
    }
    allow(event);
  });
}

// OnCompose - 새 메일 작성 시 실행 (필요 시 초기화 작업 등 추가 가능)
function onMessageComposeHandler(event) {
  console.log("새 메일 작성 시작");
  if (event && event.completed) event.completed();
}

// OnCompose - 새 약속 작성 시 실행
function onAppointmentComposeHandler(event) {
  console.log("새 약속 작성 시작");
  if (event && event.completed) event.completed();
}

// XML에 정의된 FunctionName과 매핑
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
Office.actions.associate("onAppointmentSendHandler", onAppointmentSendHandler);
Office.actions.associate("onMessageComposeHandler", onMessageComposeHandler);
Office.actions.associate("onAppointmentComposeHandler", onAppointmentComposeHandler);
