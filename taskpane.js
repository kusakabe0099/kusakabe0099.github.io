const checks = document.querySelectorAll(".confirm");
const button = document.getElementById("sendButton");

checks.forEach(c => {
    c.addEventListener("change", () => {

        const allChecked =
            [...checks].every(x => x.checked);

        button.disabled = !allChecked;
    });
});