Office.onReady(() => {
  const item = Office.context.mailbox.item;

  item.subject.getAsync((result) => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      const subject = result.value;

      if (!subject || subject.trim() === "") {
        // 제목이 없을 경우 전송 취소
        Office.context.mailbox.item.notificationMessages.addAsync("nosubject", {
          type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
          message: "메일 제목이 비어 있습니다. 제목을 입력해 주세요."
        });

        Office.context.ui.closeContainer(); // 전송 취소
      } else {
        Office.context.ui.messageParent("pass"); // 통과
      }
    }
  });
});
