/* global Office */
Office.onReady();

// --- Helpers ---
function allow(event) {
  event.completed({ allowEvent: true });
}
function block(event, msg) {
  event.completed({ allowEvent: false, errorMessage: msg });
}
function allowAfterDelay(event, ms) {
  // Simple delay, then allow the send to proceed
  setTimeout(function () { allow(event); }, ms);
}

// --- Launch Event handlers ---

// No-op compose handlers (required by manifest bindings)
function onMessageComposeHandler(event) {
  if (event && event.completed) event.completed();
}
function onAppointmentComposeHandler(event) {
  if (event && event.completed) event.completed();
}

// OnSend: add a 5-second delay, then allow
function onMessageSendHandler(event) {
  try {
    // If you want to validate something here, do it before calling allowAfterDelay.
    allowAfterDelay(event, 5000);
  } catch (e) {
    block(event, "Unexpected error in onMessageSendHandler.");
  }
}

// OnSend (appointments): add a 5-second delay, then allow
function onAppointmentSendHandler(event) {
  try {
    allowAfterDelay(event, 5000);
  } catch (e) {
    block(event, "Unexpected error in onAppointmentSendHandler.");
  }
}

// Map JS functions to the FunctionName values in the manifest
Office.actions.associate("onMessageComposeHandler", onMessageComposeHandler);
Office.actions.associate("onAppointmentComposeHandler", onAppointmentComposeHandler);
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
Office.actions.associate("onAppointmentSendHandler", onAppointmentSendHandler);
