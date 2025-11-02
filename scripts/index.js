Office.initialize = function () {
}

var item;

Office.onReady(function () {

    try {
        // Office is ready 
        item = Office.context.mailbox.item;

        $(document).ready(function () {
            // The document is ready
			
			if(item) {
				try {
					//메시지 제목 가져오기
					const subject = item.subject;
					console.log(subject);
					
					//메시지 ItemId 가져오기
					const itemId = encodeURIComponent(item.itemId);
					console.log(itemId);
					
					//페이지에 메시지 제목과 ItemId 출력
					$("#mSubject").text(subject);
					$("#mItemId").text(itemId);

				} catch (ex) {
					console.log(ex.message);
				}
			} else {
				console.log("메일 아이템을 가져오지 못했습니다.");
			}
        });

    } catch (ex) {
        console.log(ex.message);
    }

});