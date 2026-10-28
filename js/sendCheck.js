/* 送信確認ダイアログ（モーダル）。
   表示データは親（taskpane.js）から URL のハッシュで受け取る。
   ダイアログからは Outlook のメールを直接操作できないため、
   「確認した」「キャンセル」を親に通知するだけにして、保存と送信は親が行う。 */

/* ウィンドウが小さくなって下が切れたときの対応（下の watchClipping）で使う設定。
   Office.onReady のコールバックが読み込み直後に実行されても参照できるよう、先頭で宣言しておく。 */
const CLIP_RATIO = 0.98;         // 見えている割合がこれ未満なら「下が切れている」とみなす
const MIN_VISIBLE_RATIO = 0.3;   // これ未満は、非表示や画面外の可能性があるので無視する
const MIN_HEIGHT_PERCENT = 30;   // commands.js の DIALOG_MIN_HEIGHT_PERCENT と同じ値
const MAX_ATTEMPTS = 3;          // 開き直しを繰り返す回数の上限（無限ループ防止）
const SETTLE_MS = 600;           // ウィンドウをドラッグしている最中に何度も開き直さないための待ち時間

let attempt = 0;

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

    restoreChecked(data.checked); // 高さを小さくして開き直したときは、チェック状態を引き継ぐ
    updateProgress();
    watchClipping(data);
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
        none.className = "empty-row";
        none.textContent = "なし";
        container.appendChild(none);
        return;
    }

    recipients.forEach(r => {

        // 社内/社外は、送信者と同じドメインかどうかで親（shared.js）が判定済み。
        // 判定がない場合は、安全側に倒して社外として表示する
        const isExternal = r.external !== false;

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
        none.className = "empty-row";
        none.textContent = "添付ファイルなし";
        container.appendChild(none);
        return;
    }

    files.forEach(file => {

        const row = document.createElement("div");
        row.className = "recipient-row";

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

/* ---------- ウィンドウが小さくなって、ダイアログの下が切れたときの対応 ----------
   ダイアログの大きさは開くときにしか指定できず、あとからウィンドウが小さくされると、
   OWA 側のダイアログに収まらなくなって外側にもスクロールバーが出る（二重スクロール）。
   外側は変更できないため、「自分の何割が見えているか」を監視し、切れていたら
   親（commands.js）に高さを小さくして開き直すよう依頼する。 */

function watchClipping(data) {

    // 別オリジンの iframe からでも、IntersectionObserver で「見えている割合」は取得できる
    if (!("IntersectionObserver" in window)) return;

    attempt = Number(data.attempt) || 0;

    let timer = null;
    const thresholds = Array.from({ length: 51 }, (_, i) => i / 50);

    const observer = new IntersectionObserver(entries => {

        const ratio = entries[entries.length - 1].intersectionRatio;

        clearTimeout(timer);

        if (ratio >= CLIP_RATIO) {
            attempt = 0; // 収まっている。次に切れたときは、また最初から数える
            return;
        }

        if (ratio < MIN_VISIBLE_RATIO || attempt >= MAX_ATTEMPTS) return;

        timer = setTimeout(() => requestShrink(data, ratio), SETTLE_MS);

    }, { threshold: thresholds });

    // .container はダイアログ（iframe）の高さいっぱいに広がっている
    observer.observe(document.querySelector(".container"));
}

function requestShrink(data, ratio) {

    const current = Number(data.heightPercent) || 70;
    const next = Math.max(MIN_HEIGHT_PERCENT, Math.floor(current * ratio) - 2);

    if (next >= current) return;

    Office.context.ui.messageParent(JSON.stringify({
        type: "resize",
        heightPercent: next,
        attempt: attempt + 1,
        checked: [...document.querySelectorAll(".confirm")].map(cb => cb.checked)
    }));
}

function restoreChecked(checked) {

    const boxes = [...document.querySelectorAll(".confirm")];

    // 宛先・添付の数が合わないときは、引き継がない
    if (!Array.isArray(checked) || checked.length !== boxes.length) return;

    boxes.forEach((cb, i) => { cb.checked = !!checked[i]; });
}

function onCancel() {

    Office.context.ui.messageParent(JSON.stringify({ type: "cancel" }));
}
