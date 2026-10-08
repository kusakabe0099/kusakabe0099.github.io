/* 送信確認タスクペイン。
   全項目にチェックして「確認完了」を押すと、確認済みフラグ（宛先・添付の署名）を保存する。
   実際の送信可否は commands.js の OnMessageSend ハンドラが、このフラグを見て判定する。 */

let currentSignature = null;
let refreshSeq = 0;

Office.onReady(() => {
    const item = Office.context.mailbox.item;

    document.getElementById("sendButton").addEventListener("click", allowMailSend);

    // 宛先・添付が編集されたら再描画（チェックはリセットされる）
    item.addHandlerAsync(Office.EventType.RecipientsChanged, refresh);
    item.addHandlerAsync(Office.EventType.AttachmentsChanged, refresh);

    refresh();
});

// 件名・宛先・添付を1回の取得結果から描画する（取得と描画がずれないようにする）
async function refresh() {

    const seq = ++refreshSeq;
    const item = Office.context.mailbox.item;

    let state, subject;

    try {
        [state, subject] = await Promise.all([
            Okan.collect(item),
            Okan.call(cb => item.subject.getAsync(cb))
        ]);
    } catch (e) {
        setMessage("読み込めませんでした。画面を開き直してください。", "error");
        return;
    }

    // 後から呼ばれた refresh があれば、古い結果は捨てる
    if (seq !== refreshSeq) return;

    currentSignature = state.signature;

    document.getElementById("subjectText").textContent = subject;

    ["To", "Cc", "Bcc"].forEach(kind => {
        renderRecipientList(
            kind.toLowerCase() + "List",
            state.recipients.filter(r => r.kind === kind)
        );
    });

    renderAttachmentList("attachList", state.files);

    setMessage("");
    updateProgress();
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

async function allowMailSend() {

    const item = Office.context.mailbox.item;

    try {
        // 表示後に宛先・添付が変わっていないか再確認してから保存する
        const latest = await Okan.collect(item);

        if (latest.signature !== currentSignature) {
            await refresh();
            setMessage("宛先または添付ファイルが変更されました。もう一度確認してください。", "error");
            return;
        }

        // roamingSettings はメールボックス全体で永続するため使わない。
        // このメール（作成セッション）にだけ有効な sessionData に、署名を保存する。
        await Okan.setConfirmed(item, latest.signature);

        setMessage("確認を保存しました。メールの「送信」をもう一度押してください。", "ok");

    } catch (e) {
        console.error("allowMailSend", e);
        setMessage("確認を保存できませんでした。もう一度お試しください。", "error");
    }
}
