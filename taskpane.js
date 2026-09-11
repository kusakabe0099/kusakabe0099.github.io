document.addEventListener("DOMContentLoaded", () => {

  const checks = document.querySelectorAll(".confirm");
  const button = document.getElementById("sendButton");

  Office.onReady(() => {
    console.log(
        Office.context.requirements.isSetSupported(
            "Mailbox",
            "1.12"
      )
    );
  });
  
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
