// 전역 변수로 현재 이메일 아이템을 저장할 객체
var mailboxItem;

Office.initialize = function (reason) {
  // Office 객체 초기화 시 현재 Item을 변수에 할당
  mailboxItem = Office.context.mailbox.item;
};

// Outlook 메일 전송(ItemSend) 이벤트 발생 시 호출되는 함수
// <param name="event">메일 전송 이벤트 컨텍스트 객체</param>
function validateBody(event) {
  // 1. 현재 아이템이 약속(회의 일정)인지 확인. 약속 아이템은 검사 대상에서 제외.
  if (mailboxItem.itemType === Office.MailboxEnums.ItemType.Appointment) {
    // 약속일 경우 그대로 전송 허용
    event.completed({ allowEvent: true });
    return;
  }

  // 2. 메일 본문을 텍스트 형태로 비동기 가져오기
  mailboxItem.body.getAsync("text", { asyncContext: event }, function(asyncResult) {
    var sendEvent = asyncResult.asyncContext;  // 이벤트 컨텍스트 저장
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      var bodyText = asyncResult.value;
      // 금지어 목록 정의 (소문자로 비교하여 대소문자 무시)
      var forbiddenWords = ["confidential", "password"];  // 금지어 배열 (예시)
      var bodyLower = bodyText.toLowerCase();
      var foundForbidden = false;
      // 금지어 목록에 있는 단어가 본문에 포함되어 있는지 확인
      for (var i = 0; i < forbiddenWords.length; i++) {
        if (bodyLower.indexOf(forbiddenWords[i]) !== -1) {
          foundForbidden = true;
          break;
        }
      }
      if (foundForbidden) {
        // 3. 금지어 발견 시: 사용자에게 알림 메시지를 표시하고 전송 차단
        mailboxItem.notificationMessages.addAsync('BlockSend', {
          type: 'errorMessage',
          message: '메일 본문에 금지어가 포함되어 있어 전송이 취소되었습니다.'
        });
        // 전송을 차단하도록 이벤트 완료 (allowEvent: false)
        sendEvent.completed({ allowEvent: false });
      } else {
        // 4. 금지어가 없을 경우: 전송 허용
        sendEvent.completed({ allowEvent: true });
      }
    } else {
      // 본문을 가져오는 데 실패한 경우 (예외 상황): 기본적으로 전송 허용
      sendEvent.completed({ allowEvent: true });
    }
  });
}
