/* global Office */
Office.onReady();

// 공통 헬퍼
function block(event, msg) {
  event.completed({ allowEvent: false, errorMessage: msg });
}
function allow(event) {
  event.completed({ allowEvent: true });
}

// 메일 전송 시 제목 체크
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

// 약속 전송 시 제목 체크
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

/* 아래 두 함수는 리본 버튼(수동 체크) 쓰지 않으면 삭제해도 됩니다.
   XML에 버튼 정의가 없으면 실행되지 않습니다. */
function manualCheckSubject(event) {
  Office.context.mailbox.item.subject.getAsync(function (r) {
    const empty = (r.status === Office.AsyncResultStatus.Succeeded
      ? (r.value || "").trim() === ""
      : true);
    if (empty) {
      Office.context.ui.displayDialogAsync(
        "https://tlstofha.github.io/addintest/alert.html?msg=" +
          encodeURIComponent("제목이 비었습니다.")
      );
    }
    if (event && event.completed) event.completed();
  });
}

function manualCheckAppointment(event) {
  Office.context.mailbox.item.subject.getAsync(function (r) {
    const empty = (r.status === Office.AsyncResultStatus.Succeeded
      ? (r.value || "").trim() === ""
      : true);
    if (empty) {
      Office.context.ui.displayDialogAsync(
        "https://tlstofha.github.io/addintest/alert.html?msg=" +
          encodeURIComponent("약속 제목이 비었습니다.")
      );
    }
    if (event && event.completed) event.completed();
  });
}

// XML에 정의된 FunctionName과 매핑
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
Office.actions.associate("onAppointmentSendHandler", onAppointmentSendHandler);
Office.actions.associate("manualCheckSubject", manualCheckSubject);
Office.actions.associate("manualCheckAppointment", manualCheckAppointment);
