/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

Office.onReady();

Office.onReady();

const HOLD_MS = 10000; // 무조건 10초 지연

function allow(event) {
  event.completed({ allowEvent: true });
}
function allowAfterDelay(event, ms = HOLD_MS) {
  setTimeout(() => allow(event), ms);
}

function onMessageSendHandler(event) {
  // 제목 체크 없이 무조건 10초 지연 후 전송 허용
  allowAfterDelay(event);
}

function onAppointmentSendHandler(event) {
  // 제목 체크 없이 무조건 10초 지연 후 전송 허용
  allowAfterDelay(event);
}

function onMessageComposeHandler(event) { event.completed(); }
function onAppointmentComposeHandler(event) { event.completed(); }

Office.actions.associate("onMessageComposeHandler", onItemComposeHandler);
Office.actions.associate("onAppointmentComposeHandler", onItemComposeHandler);
Office.actions.associate("onMessageSendHandler", onItemSendHandler);
Office.actions.associate("onAppointmentSendHandler", onItemSendHandler);