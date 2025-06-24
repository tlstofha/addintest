/*
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

Office.initialize = function (reason) {};

/**
 * Handles the OnMessageRecipientsChanged event.
 * @param {*} event The Office event object
 */
function tagExternal_onMessageRecipientsChangedHandler(event) {
  console.log("tagExternal_onMessageRecipientsChangedHandler method"); //debugging
  console.log("event: " + JSON.stringify(event)); //debugging
  if (event.changedRecipientFields.to) {
    checkForExternalTo();
  }
  if (event.changedRecipientFields.cc) {
    checkForExternalCc();
  }
  if (event.changedRecipientFields.bcc) {
    checkForExternalBcc();
  }
}

/**
 * Handles the ItemSend (OnSend) event.
 * Ensures tagging/disclaimer is consistent just before send.
 * @param {*} event The Office event object
 */
function tagExternal_onSendHandler(event) {
  try {
    // Re-evaluate external recipients and subject / disclaimer.
    _checkForExternal();               // reuse existing function
    event.completed({ allowEvent: true });
  } catch (e) {
    // Block send if something goes wrong
    event.completed({
      allowEvent: false,
      errorMessage: "Send blocked: " + (e.message || e.toString())
    });
  }
}

/* ---------- helper functions (unchanged) ---------- */

function checkForExternalTo() {
  console.log("checkForExternalTo method");
  Office.context.mailbox.item.to.getAsync(function (asyncResult) {
    if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
      console.error("Failed to get To recipients. " + JSON.stringify(asyncResult.error));
      return;
    }
    const toRecipients = JSON.stringify(asyncResult.value);
    const keyName = "tagExternalTo";
    if (toRecipients && toRecipients.includes(Office.MailboxEnums.RecipientType.ExternalUser)) {
      _setSessionData(keyName, true);
    } else {
      _setSessionData(keyName, false);
    }
  });
}

function checkForExternalCc() {
  console.log("checkForExternalCc method");
  Office.context.mailbox.item.cc.getAsync(function (asyncResult) {
    if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
      console.error("Failed to get Cc recipients. " + JSON.stringify(asyncResult.error));
      return;
    }
    const ccRecipients = JSON.stringify(asyncResult.value);
    const keyName = "tagExternalCc";
    if (ccRecipients && ccRecipients.includes(Office.MailboxEnums.RecipientType.ExternalUser)) {
      _setSessionData(keyName, true);
    } else {
      _setSessionData(keyName, false);
    }
  });
}

function checkForExternalBcc() {
  console.log("checkForExternalBcc method");
  Office.context.mailbox.item.bcc.getAsync(function (asyncResult) {
    if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
      console.error("Failed to get Bcc recipients. " + JSON.stringify(asyncResult.error));
      return;
    }
    const bccRecipients = JSON.stringify(asyncResult.value);
    const keyName = "tagExternalBcc";
    if (bccRecipients && bccRecipients.includes(Office.MailboxEnums.RecipientType.ExternalUser)) {
      _setSessionData(keyName, true);
    } else {
      _setSessionData(keyName, false);
    }
  });
}

function _setSessionData(key, value) {
  Office.context.mailbox.item.sessionData.setAsync(
    key,
    value.toString(),
    function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        console.log(`sessionData.setAsync(${key}) to ${value} succeeded`);
        if (value) {
          _tagExternal(value);
        } else {
          _checkForExternal();
        }
      } else {
        console.error(`Failed to set ${key} sessionData to ${value}. Error: ${JSON.stringify(asyncResult.error)}`);
      }
    });
}

function _checkForExternal() {
  console.log("_checkForExternal method");
  Office.context.mailbox.item.sessionData.getAllAsync(function (asyncResult) {
    if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
      console.error("Failed to get all sessionData. " + JSON.stringify(asyncResult.error));
      return;
    }
    const sessionData = JSON.stringify(asyncResult.value);
    if (sessionData && sessionData.includes("true")) {
      _tagExternal(true);
    } else {
      _tagExternal(false);
    }
  });
}

function _tagExternal(hasExternal) {
  console.log("_tagExternal method");
  const externalTag = "[External]";

  if (hasExternal) {
    // prepend subject tag
    Office.context.mailbox.item.subject.getAsync(function (asyncResult) {
      if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
        console.error("Failed to get subject. " + JSON.stringify(asyncResult.error));
        return;
      }
      let subject = asyncResult.value;
      if (!subject.includes(externalTag)) {
        subject = `${externalTag} ${subject}`;
        Office.context.mailbox.item.subject.setAsync(subject, function (a) {
          if (a.status !== Office.AsyncResultStatus.Succeeded) {
            console.error("Failed to set Subject. " + JSON.stringify(a.error));
          }
        });
      }
    });

    const disclaimer = '<p style="color:blue"><i>Caution: This email includes external recipients.</i></p>';
    Office.context.mailbox.item.body.appendOnSendAsync(disclaimer, { coercionType: Office.CoercionType.Html });
  } else {
    // remove tag
    Office.context.mailbox.item.subject.getAsync(function (asyncResult) {
      if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
        console.error("Failed to get subject. " + JSON.stringify(asyncResult.error));
        return;
      }
      const currentSubject = asyncResult.value;
      const externalTag = "[External]";
      if (currentSubject.startsWith(externalTag)) {
        const updatedSubject = currentSubject.replace(externalTag, "").trim();
        Office.context.mailbox.item.subject.setAsync(updatedSubject, function (a) {
          if (a.status !== Office.AsyncResultStatus.Succeeded) {
            console.error("Failed to set Subject. " + JSON.stringify(a.error));
          }
        });
      }
    });
    // clear disclaimer
    Office.context.mailbox.item.body.appendOnSendAsync(null);
  }
}

/* ---------- Office.actions association ---------- */
Office.actions.associate(
  "tagExternal_onMessageRecipientsChangedHandler",
  tagExternal_onMessageRecipientsChangedHandler
);
Office.actions.associate(
  "tagExternal_onSendHandler",
  tagExternal_onSendHandler          // ← new mapping
);
