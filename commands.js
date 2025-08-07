/* global Office */
Office.onReady();

// 공통 헬퍼
function block(event, msg) {
  event.completed({ allowEvent: false, errorMessage: msg });
}
function allow(event) {
  event.completed({ allowEvent: true });
}

// 메일 전송 시 제목 확인
function onMessageSendHandler(event) {
	return block(event, "테스트 입니다");
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

// 약속 전송 시 제목 확인
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

// 새 메일/약속 작성 시 자동 로드용 dummy 핸들러
function onMessageComposeHandler(event) {
  event.completed();
}
function onAppointmentComposeHandler(event) {
  event.completed();
}

// 이벤트 연결
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
Office.actions.associate("onAppointmentSendHandler", onAppointmentSendHandler);
Office.actions.associate("onMessageComposeHandler", onMessageComposeHandler);
Office.actions.associate("onAppointmentComposeHandler", onAppointmentComposeHandler);