Office.onReady(async () => {

    const item = Office.context.mailbox.item;
    
    loadSubject(item);
    loadTo(item);
    loadCc(item);
    loadBcc(item);
    loadAttachments(item);

});

function loadSubject(item) {

    item.subject.getAsync(res => {

        if (res.status === Office.AsyncResultStatus.Succeeded) {

            document.getElementById("subjectText")
                .innerText = res.value;

        }

    });

}

function loadRecipients(item) {

    item.to.getAsync(result => {

        renderAddressList(
            "toList",
            result.value
        );

    });

}

function loadAttachments(item) {

    item.getAttachmentsAsync(result => {

        if (
            result.status ===
            Office.AsyncResultStatus.Succeeded
        ) {

            renderAttachmentList(
                "attachList",
                result.value
            );

        }

    });

}

function renderAddressList(id, recipients) {

    const container =
        document.getElementById(id);

    recipients.forEach(r => {

        const row =
        document.createElement("div");

        row.className = "row";

        row.innerHTML = `
            <input type="checkbox"
                   class="confirm">

            <span>${r.displayName}</span>

            <span>${r.emailAddress}</span>
        `;

        container.appendChild(row);

    });

    bindCheckEvents();
}

function renderAttachmentList(id, files) {

    const container =
        document.getElementById(id);

    container.innerHTML = "";

    if (!files || files.length === 0) {

        container.innerHTML =
            "<div>添付ファイルなし</div>";

        return;
    }

    files.forEach(file => {

        const row =
            document.createElement("div");

        row.innerHTML = `
            <label>
                <input
                    type="checkbox"
                    class="confirm">
                ${file.name}
            </label>
        `;

        container.appendChild(row);

    });

    bindCheckEvents();
}

function bindCheckEvents() {

    document
        .querySelectorAll(".confirm")
        .forEach(cb => {

            cb.addEventListener(
                "change",
                updateProgress
            );

        });

}

function updateProgress() {

    const checks =
      document.querySelectorAll(".confirm");

    const completed =
      [...checks].filter(x => x.checked)
                 .length;

    const total =
      checks.length;

    document.getElementById("progress")
        .innerText =
        `確認済み: ${completed}/${total}`;

    document.getElementById("sendButton")
        .disabled =
        completed !== total;

}

function loadTo(item) {

    item.to.getAsync(result => {

        if (
            result.status ===
            Office.AsyncResultStatus.Succeeded
        ) {

            console.log("TO", result.value);

            renderRecipientList(
                "toList",
                result.value
            );

        }

    });

}

function loadCc(item) {

    item.cc.getAsync(result => {

        if (
            result.status ===
            Office.AsyncResultStatus.Succeeded
        ) {

            console.log("CC", result.value);

            renderRecipientList(
                "ccList",
                result.value
            );

        }

    });

}

function loadBcc(item) {

    item.bcc.getAsync(result => {

        if (
            result.status ===
            Office.AsyncResultStatus.Succeeded
        ) {

            console.log("BCC", result.value);

            renderRecipientList(
                "bccList",
                result.value
            );

        }

    });

}

function renderRecipientList(id, recipients) {

    const container =
        document.getElementById(id);

    container.innerHTML = "";

    if (!recipients || recipients.length === 0) {

        container.innerHTML = "<div>なし</div>";
        return;
    }

    recipients.forEach(r => {

        const isExternal =
            !r.emailAddress
                .toLowerCase()
                .endsWith("@company.co.jp");

        const row =
            document.createElement("div");

        row.className =
            isExternal
                ? "recipient-row external"
                : "recipient-row";

        row.innerHTML = `
            <label>
                <input
                    type="checkbox"
                    class="confirm">

                ${r.displayName}
                (${r.emailAddress})
            </label>
        `;

        container.appendChild(row);

    });

    bindCheckEvents();
}

document.addEventListener(
  "DOMContentLoaded",
  () => {

      const btn =
        document.getElementById(
          "sendButton"
        );

      console.log(btn);

      if (!btn) {
          return;
      }

      btn.addEventListener(
        "click",
        allowMailSend
      );

  }
);

