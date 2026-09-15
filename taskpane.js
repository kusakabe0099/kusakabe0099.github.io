console.log("START");

document.addEventListener("DOMContentLoaded", () => {

  const checks = document.querySelectorAll(".confirm");
  const button = document.getElementById("sendButton");
  
  checks.forEach(check => {
    check.addEventListener("change", () => {

      const allChecked =
        [...checks].every(c => c.checked);

      button.disabled = !allChecked;

      if (allChecked) {
        button.textContent = "送信OK";
      } else {
        button.textContent = "送信可能";
      }
    });
  });

});
/*
Office.onReady(() => {

    const supported =
        Office.context.requirements.isSetSupported(
            "Mailbox",
            "1.12"
        );

    document.body.innerHTML =
        "Mailbox 1.12 対応: " + supported;

});
*/
