/* global Office */

function allow(event){ event.completed({ allowEvent: true }); }
function block(event,msg){ event.completed({ allowEvent:false, errorMessage: msg }); }
function hold(event, ms){ setTimeout(function(){ allow(event); }, ms); }

function onMessageComposeHandler(event){ if(event && event.completed) event.completed(); }
function onAppointmentComposeHandler(event){ if(event && event.completed) event.completed(); }

function onMessageSendHandler(event){
  try { hold(event, 10000); } // 10 seconds
  catch(e){ block(event, 'Unexpected error in onMessageSendHandler'); }
}

function onAppointmentSendHandler(event){
  try { hold(event, 10000); } // 10 seconds
  catch(e){ block(event, 'Unexpected error in onAppointmentSendHandler'); }
}

Office.actions.associate('onMessageComposeHandler', onMessageComposeHandler);
Office.actions.associate('onAppointmentComposeHandler', onAppointmentComposeHandler);
Office.actions.associate('onMessageSendHandler', onMessageSendHandler);
Office.actions.associate('onAppointmentSendHandler', onAppointmentSendHandler);
