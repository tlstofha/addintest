Office.onReady();

// --- helpers ---
function block(event, msg) {
  event.completed({ allowEvent: false, errorMessage: msg });
}
function allow(event) {
  event.completed({ allowEvent: true });
}

// --- Appointment send only ---
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

// (Optional) manual check button
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

// --- associate ---
Office.actions.associate("onAppointmentSendHandler", onAppointmentSendHandler);
Office.actions.associate("manualCheckAppointment", manualCheckAppointment);