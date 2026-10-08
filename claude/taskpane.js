/* 起動役のタスクペイン。
   Smart Alerts の「確認画面を開く」から開かれ、確認ダイアログ（sendCheck.html）をモーダル表示する。
   ダイアログで「確認して送信」が押されたら、確認済みフラグを保存して送信する。 */

let dialog = null;
let shownSignature = null; // ダイアログに表示した内容の署名

Office.onReady(() => {

    document.getElementById("openButton").addEventListener("click", openDialog);

    // 開いた時点で自動表示する（環境によっては許可の確認が出る。その場合はボタンから開く）
    openDialog();
});

async function openDialog() {

    if (dialog) return;

    setMessage("");

    const item = Office.context.mailbox.item;
    let state, subject;

    try {
        [state, subject] = await Promise.all([
            Okan.collect(item),
            Okan.call(cb => item.subject.getAsync(cb))
        ]);
    } catch (e) {
        setMessage("宛先と添付ファイルを読み込めませんでした。もう一度お試しください。", "error");
        return;
    }

    shownSignature = state.signature;

    const payload = {
        subject,
        recipients: state.recipients,
        files: state.files
    };

    // 宛先などをサーバーへ送らないよう、ハッシュ（#）で渡す
    const url = new URL("sendCheck.html", location.href);
    url.hash = encodeURIComponent(JSON.stringify(payload));

    Office.context.ui.displayDialogAsync(
        url.toString(),
        { width: 45, height: 85, displayInIframe: true }, // Outlook on the web では画面内のモーダルとして表示
        result => {

            if (result.status !== Office.AsyncResultStatus.Succeeded) {

                const code = result.error.code;
                setMessage(
                    code === 12009
                        ? "確認画面の表示が許可されませんでした。「確認画面を開く」を押して、許可してください。"
                        : `確認画面を開けませんでした（エラー ${code}）。`,
                    "error"
                );
                return;
            }

            dialog = result.value;
            dialog.addEventHandler(Office.EventType.DialogMessageReceived, onDialogMessage);
            dialog.addEventHandler(Office.EventType.DialogEventReceived, onDialogEvent);
        }
    );
}

function closeDialog() {

    if (!dialog) return;

    try {
        dialog.close();
    } catch (e) {
        // すでに閉じている場合は何もしない
    }
    dialog = null;
}

// ダイアログが×ボタンなどで閉じられた、または読み込みに失敗した
function onDialogEvent(arg) {

    dialog = null;

    setMessage(
        arg.error === 12006
            ? "確認画面が閉じられました。送信するには「確認画面を開く」を押してください。"
            : `確認画面でエラーが発生しました（エラー ${arg.error}）。`,
        "error"
    );
}

async function onDialogMessage(arg) {

    let msg;

    try {
        msg = JSON.parse(arg.message);
    } catch (e) {
        return;
    }

    closeDialog();

    if (msg.type === "cancel") {
        setMessage("キャンセルしました。送信するには「確認画面を開く」を押してください。");
        return;
    }

    if (msg.type === "confirm") {
        await saveAndSend();
    }
}

async function saveAndSend() {

    const item = Office.context.mailbox.item;

    try {
        // 確認画面を見せたあとに宛先・添付が変わっていないか再確認してから保存する
        const latest = await Okan.collect(item);

        if (latest.signature !== shownSignature) {
            setMessage("確認後に宛先または添付ファイルが変更されました。もう一度確認してください。", "error");
            return;
        }

        // roamingSettings はメールボックス全体で永続するため使わない。
        // このメール（作成セッション）にだけ有効な sessionData に、署名を保存する。
        await Okan.setConfirmed(item, latest.signature);

    } catch (e) {
        console.error("setConfirmed", e);
        setMessage("確認を保存できませんでした。もう一度お試しください。", "error");
        return;
    }

    // Mailbox 1.15 以降なら、このままここから送信する（確認済みフラグは保存済み）
    if (Office.context.requirements.isSetSupported("Mailbox", "1.15") &&
        typeof item.sendAsync === "function") {
        try {
            await Okan.call(cb => item.sendAsync(cb));
            return;
        } catch (e) {
            console.error("sendAsync", e);
            setMessage("確認は保存しましたが、送信できませんでした。メールの「送信」を押してください。", "error");
            return;
        }
    }

    // 1.15 未対応の環境：確認だけ保存し、ユーザーに送信してもらう
    setMessage("確認を保存しました。メールの「送信」をもう一度押してください。", "ok");
}

function setMessage(text, cls) {

    const m = document.getElementById("message");
    m.textContent = text;
    m.className = cls || "";
}
