Office.onReady(() => {
    loadConfirmationData();
  });
  
  function loadConfirmationData() {
    Office.context.roamingSettings.refreshAsync(() => {
      const raw = Office.context.roamingSettings.get("confirmData");
  
      if (!raw) {
        document.querySelector('.container').innerHTML =
          '<p style="color:red;padding:20px">Error al cargar datos. Cierra este panel e intenta de nuevo.</p>';
        return;
      }
  
      const data = JSON.parse(raw);
      renderRecipients(data.externalRecipients);
      renderAttachments(data.attachments);
    });
  }
  
  function renderRecipients(recipients) {
    const ul = document.getElementById("recipientsList");
    ul.innerHTML = "";
    recipients.forEach(r => {
      const li = document.createElement("li");
      li.textContent = r.displayName
        ? `${r.displayName} — ${r.emailAddress}`
        : r.emailAddress;
      ul.appendChild(li);
    });
  }
  
  function renderAttachments(attachments) {
    const ul = document.getElementById("attachmentsList");
    ul.innerHTML = "";
    attachments.forEach(a => {
      const li = document.createElement("li");
      const size = a.size ? ` (${(a.size / 1024).toFixed(1)} KB)` : "";
      li.textContent = `📄 ${a.name}${size}`;
      ul.appendChild(li);
    });
  }
  
  function toggleSendButton() {
    const checked = document.getElementById("confirmCheck").checked;
    document.getElementById("confirmBtn").disabled = !checked;
  }
  
  function confirmSend() {
    Office.context.ui.messageParent(JSON.stringify({ confirmed: true }));
  }
  
  function cancelSend() {
    Office.context.ui.messageParent(JSON.stringify({ confirmed: false }));
  }
