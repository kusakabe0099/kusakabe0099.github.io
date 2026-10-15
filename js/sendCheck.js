/* 送信確認ダイアログ（モーダル）。
   表示データは親（taskpane.js）から URL のハッシュで受け取る。
   ダイアログからは Outlook のメールを直接操作できないため、
   「確認した」「キャンセル」を親に通知するだけにして、保存と送信は親が行う。 */

Office.onReady(() => {

    document.getElementById("sendButton").addEventListener("click", onConfirm);
    document.getElementById("cancelButton").addEventListener("click", onCancel);

    const data = readData();

    if (!data) {
        setMessage("確認する内容を受け取れませんでした。閉じて、もう一度開いてください。", "error");
        return;
    }

    document.getElementById("subjectText").textContent = data.subject || "（件名なし）";

    ["To", "Cc", "Bcc"].forEach(kind => {
        renderRecipientList(
            kind.toLowerCase() + "List",
            data.recipients.filter(r => r.kind === kind)
        );
    });

    renderAttachmentList("attachList", data.files);

    updateProgress();
});

// 宛先などをサーバーへ送らないよう、クエリ（?）ではなくハッシュ（#）で受け取っている
function readData() {

    try {
        return JSON.parse(decodeURIComponent(location.hash.slice(1)));
    } catch (e) {
        return null;
    }
}

// innerHTML を使わず textContent で表示する（表示名に含まれる文字をHTMLとして解釈しない）
function renderRecipientList(id, recipients) {

    const container = document.getElementById(id);
    container.replaceChildren();

    if (!recipients || recipients.length === 0) {
        const none = document.createElement("div");
        none.textContent = "なし";
        container.appendChild(none);
        return;
    }

    recipients.forEach(r => {

        const isExternal = Okan.isExternal(Okan.domainOf(r.address));

        const row = document.createElement("div");
        row.className = isExternal ? "recipient-row external" : "recipient-row";

        const label = document.createElement("label");
        label.appendChild(createCheckbox());
        label.appendChild(document.createTextNode(
            r.name && r.name !== r.address
                ? ` ${r.name} (${r.address})`
                : ` ${r.address}`
        ));

        row.appendChild(label);
        container.appendChild(row);
    });
}

function renderAttachmentList(id, files) {

    const container = document.getElementById(id);
    container.replaceChildren();

    if (!files || files.length === 0) {
        const none = document.createElement("div");
        none.textContent = "添付ファイルなし";
        container.appendChild(none);
        return;
    }

    files.forEach(file => {

        const row = document.createElement("div");

        const label = document.createElement("label");
        label.appendChild(createCheckbox());
        label.appendChild(document.createTextNode(" " + file.name));

        row.appendChild(label);
        container.appendChild(row);
    });
}

function createCheckbox() {

    const cb = document.createElement("input");
    cb.type = "checkbox";
    cb.className = "confirm";
    cb.addEventListener("change", updateProgress);
    return cb;
}

function updateProgress() {

    const checks = [...document.querySelectorAll(".confirm")];
    const completed = checks.filter(x => x.checked).length;
    const total = checks.length;

    document.getElementById("progress").textContent =
        `確認済み: ${completed}/${total}`;

    // チェック項目が0件のときは有効にしない
    document.getElementById("sendButton").disabled =
        total === 0 || completed !== total;
}

function setMessage(text, cls) {

    const m = document.getElementById("message");
    m.textContent = text;
    m.className = cls || "";
}

function onConfirm() {

    // 二重送信を防ぐ
    document.getElementById("sendButton").disabled = true;
    document.getElementById("cancelButton").disabled = true;
    setMessage("送信しています…");

    Office.context.ui.messageParent(JSON.stringify({ type: "confirm" }));
}

function onCancel() {

    Office.context.ui.messageParent(JSON.stringify({ type: "cancel" }));
}
