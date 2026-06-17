// ✅ CONFIGURA AQUÍ LOS DOMINIOS INTERNOS DE TU EMPRESA
const INTERNAL_DOMAINS = [
    "wtwco.com",
    "willistowerswatson.com"
    // Agrega todos los dominios internos aquí
  ];
  
  Office.initialize = function () {};
  
  function onMessageSendHandler(event) {
    const item = Office.context.mailbox.item;
  
    Promise.all([
      getRecipients(item.to),
      getRecipients(item.cc),
      getRecipients(item.bcc)
    ]).then(([toRecipients, ccRecipients, bccRecipients]) => {
  
      const allRecipients = [...toRecipients, ...ccRecipients, ...bccRecipients];
      const externalRecipients = allRecipients.filter(r => isExternal(r.emailAddress));
  
      // Sin destinatarios externos → enviar sin confirmación
      if (externalRecipients.length === 0) {
        event.completed({ allowEvent: true });
        return;
      }
  
      // Con externos pero sin adjuntos → enviar sin confirmación
      const attachments = item.attachments || [];
      if (attachments.length === 0) {
        event.completed({ allowEvent: true });
        return;
      }
  
      // Externos + adjuntos → bloquear y mostrar confirmación
      const confirmData = {
        externalRecipients: externalRecipients,
        attachments: attachments.map(a => ({ name: a.name, size: a.size }))
      };
  
      Office.context.roamingSettings.set("confirmData", JSON.stringify(confirmData));
      Office.context.roamingSettings.saveAsync(() => {
  
        event.completed({ allowEvent: false });
  
        Office.context.ui.displayDialogAsync(
          "https://TU-DOMINIO.com/src/taskpane/taskpane.html",
          { height: 60, width: 40, displayInIframe: false },
          (asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
              const dialog = asyncResult.value;
  
              dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
                const message = JSON.parse(arg.message);
                dialog.close();
                if (message.confirmed === true) {
                  Office.context.roamingSettings.remove("confirmData");
                  Office.context.roamingSettings.saveAsync();
                }
              });
  
              dialog.addEventHandler(Office.EventType.DialogEventReceived, (arg) => {
                if (arg.error === 12006) dialog.close();
              });
            }
          }
        );
      });
    });
  }
  
  function isExternal(email) {
    if (!email) return false;
    const domain = email.split("@")[1]?.toLowerCase();
    return !INTERNAL_DOMAINS.some(d =>
      domain === d.toLowerCase() || domain?.endsWith("." + d.toLowerCase())
    );
  }
  
  function getRecipients(recipientField) {
    return new Promise((resolve) => {
      if (!recipientField) { resolve([]); return; }
      recipientField.getAsync((result) => {
        resolve(result.status === Office.AsyncResultStatus.Succeeded ? result.value : []);
      });
    });
  }
  
  Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
