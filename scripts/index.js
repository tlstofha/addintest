/* index.js */

Office.initialize = function () { }

let item;

Office.onReady(function () {
  try {
    item = Office.context.mailbox.item;

    $(document).ready(function () {
      if (item) {
        try {
          const subject = item.subject;
          const itemId = encodeURIComponent(item.itemId);
          console.log("[Add-in] Subject:", subject);
          console.log("[Add-in] ItemId (EWS):", itemId);
          $("#mSubject").text(subject || "(no subject)");
          $("#mItemId").text(itemId || "(no id)");
        } catch (ex) {
          console.log("[Add-in] Read item error:", ex.message);
        }
      } else {
        console.log("[Add-in] 메일 아이템을 가져오지 못했습니다.");
      }

      // 페이지 로드시 바로 실행
      postOpenTypeExtensionExact();
    });
  } catch (ex) {
    console.log("[Add-in] onReady error:", ex.message);
  }
});

/**
 * POST https://graph.microsoft.com/v1.0/me/messages/{id}/extensions
 * Body:
 * {
 *   "@odata.type": "microsoft.graph.openTypeExtension",
 *   "extensionName": "Com.Innotek.Extension.Test",
 *   "addinTester": "TestName"
 * }
 */
async function postOpenTypeExtensionExact() {
console.log("request graph api");
  try {
    const accessToken = await Office.auth.getAccessToken({ allowSignInPrompt: true });
    if (!accessToken) {
      console.error("[Graph] Failed to acquire access token.");
      return;
    }

    const url =
      "https://graph.microsoft.com/v1.0/me/messages/" +
      "AAMkADY1ZDJjMmY0LTlkMWYtNDFlMy04OWI2LTFmNzczNzJhZjM1ZABGAAAAAAAZKl%2BQIyaYR4vSNp40i8AHBwAUrP2YSG0aTIVLeZwO3A6kAAAAAAEMAAAUrP2YSG0aTIVLeZwO3A6kAAAMQHLYAAA%3D" +
      "/extensions";

    const body = {
      "@odata.type": "microsoft.graph.openTypeExtension",
      "extensionName": "Com.Innotek.Extension.Test",
      "addinTester": "TestName"
    };

    const resp = await fetch(url, {
      method: "POST",
      headers: {
        "Authorization": Bearer ${accessToken},
        "Content-Type": "application/json"
      },
      body: JSON.stringify(body)
    });

    if (!resp.ok) {
      const text = await safeRead(resp);
      console.error([Graph] POST failed: ${resp.status} ${resp.statusText}, text);
      return;
    }

    const data = await resp.json();
    console.log("[Graph] POST success:", data);
  } catch (err) {
    console.error("[Graph] POST error:", err);
  }
}

async function safeRead(resp) {
  try {
    const ct = resp.headers.get("content-type") || "";
    if (ct.includes("application/json")) return await resp.json();
    return await resp.text();
  } catch {
    return "(no body)";
  }
}
