/* 送信時イベント（OnMessageSend）と、確認ダイアログを開くコマンド。

   【試験】OPEN_DIALOG_IN_HANDLER = true のとき、「送信」を押した時点で、
   ハンドラの中から直接確認ダイアログを開く（「確認画面を開く」ボタンを押す手間をなくす）。
   - 「確認して送信」→ そのまま送信される
   - 「キャンセル」/ ×で閉じる → 送信を中止する
   - ダイアログを開けなかった場合 → 従来の「確認画面を開く」ボタン方式にフォールバックし、
     アラートの末尾にエラー番号を表示する */

Office.onReady();

const OPEN_DIALOG_IN_HANDLER = true;               // false にすると従来方式のみ
const CONFIRM_COMMAND_ID = "msgComposeConfirmButton"; // manifest.xml のボタンIDと一致させる

/* ---------- 送信時チェック ---------- */

async function onMessageSendHandler(event) {
  try {
    const item = Office.context.mailbox.item;
    const state = await Okan.collect(item);

    // 宛先なしはOutlook側のエラーに任せる
    if (state.recipients.length === 0) {
      event.completed({ allowEvent: true });
      return;
    }

    // 確認ダイアログで承認済み、かつ宛先・添付が変わっていなければ送信を許可
    const confirmed = await Okan.getConfirmed(item);
    if (confirmed === state.signature) {
      event.completed({ allowEvent: true });
      return;
    }

    let note = "";

    if (OPEN_DIALOG_IN_HANDLER) {
      const r = await askConfirmation(item, state);

      if (r.status === "confirm") {
        if (await verifyAndSave(item, state.signature)) {
          event.completed({ allowEvent: true }); // 保留中の送信がそのまま続行される
        } else {
          event.completed({
            allowEvent: false,
            errorMessage: "確認後に宛先または添付が変更されました。もう一度「送信」を押して確認してください。",
          });
        }
        return;
      }

      if (r.status === "cancel") {
        event.completed({ allowEvent: false, errorMessage: "送信をキャンセルしました。" });
        return;
      }

      if (r.status === "closed") {
        event.completed({
          allowEvent: false,
          errorMessage: "確認画面が閉じられたため、送信を中止しました。もう一度「送信」を押してください。",
        });
        return;
      }

      // ダイアログを開けなかった場合。原因調査のため、エラーをコンソールとアラートに残す
      console.warn("displayDialogAsync from OnMessageSend failed", r);
      note = `（ダイアログ表示エラー: ${r.code}）`;
    }

    // フォールバック：「確認画面を開く」ボタン方式
    event.completed({
      allowEvent: false,
      errorMessage:
        "送信前に宛先と添付ファイルの確認が必要です。" +
        `（宛先 ${state.recipients.length} 件、うち社外 ${state.externalCount} 件、添付 ${state.files.length} 件）` +
        note,
      cancelLabel: "確認画面を開く", // 「アクションを実行」ボタンのラベル（20文字以内）
      commandId: CONFIRM_COMMAND_ID,
    });
  } catch (e) {
    console.error("onMessageSendHandler", e);
    event.completed({
      allowEvent: false,
      errorMessage: "送信前チェックでエラーが発生しました。もう一度「送信」を押してください。",
    });
  }
}

/* ---------- 確認ダイアログ ---------- */

// ダイアログを開き、ユーザーの操作が終わるまで待つ。
// 戻り値: { status: "confirm" | "cancel" | "closed" | "error", code?, message? }
async function askConfirmation(item, state) {

  let subject;

  try {
    subject = await Okan.call(cb => item.subject.getAsync(cb));
  } catch (e) {
    return { status: "error", code: "subject", message: String(e && e.message) };
  }

  const payload = { subject, recipients: state.recipients, files: state.files };

  // 宛先などをサーバーへ送らないよう、ハッシュ（#）で渡す
  const url = new URL("sendCheck.html", location.href);
  url.hash = encodeURIComponent(JSON.stringify(payload));

  return new Promise(resolve => {
    try {
      Office.context.ui.displayDialogAsync(
        url.toString(),
        { width: 45, height: 85, displayInIframe: true }, // Outlook on the web では画面内のモーダルとして表示
        result => {

          if (result.status !== Office.AsyncResultStatus.Succeeded) {
            resolve({ status: "error", code: result.error.code, message: result.error.message });
            return;
          }

          const dialog = result.value;

          dialog.addEventHandler(Office.EventType.DialogMessageReceived, arg => {
            let msg = {};
            try { msg = JSON.parse(arg.message); } catch (e) { /* 無視 */ }

            try { dialog.close(); } catch (e) { /* すでに閉じている */ }

            resolve({ status: msg.type === "confirm" ? "confirm" : "cancel" });
          });

          // ×ボタンなどで閉じられた（12006）、または読み込みに失敗した
          dialog.addEventHandler(Office.EventType.DialogEventReceived, arg => {
            resolve(
              arg.error === 12006
                ? { status: "closed" }
                : { status: "error", code: arg.error }
            );
          });
        }
      );
    } catch (e) {
      // この実行環境で displayDialogAsync 自体が使えない場合
      resolve({ status: "error", code: "exception", message: String(e && e.message) });
    }
  });
}

// 確認画面を見せたあとに宛先・添付が変わっていないか再確認し、問題なければ確認済みフラグを保存する。
// roamingSettings はメールボックス全体で永続するため使わず、
// このメール（作成セッション）にだけ有効な sessionData に署名を保存する。
async function verifyAndSave(item, shownSignature) {

  const latest = await Okan.collect(item);

  if (latest.signature !== shownSignature) return false;

  await Okan.setConfirmed(item, latest.signature);
  return true;
}

/* ---------- 「確認画面を開く」コマンド（Smart Alerts のアクション / リボンボタン） ---------- */

async function openConfirmDialog(event) {

  const item = Office.context.mailbox.item;

  try {
    const state = await Okan.collect(item);
    const r = await askConfirmation(item, state);

    if (r.status === "confirm") {
      await saveAndSend(item, state.signature);
    } else if (r.status === "cancel") {
      notify("キャンセルしました。送信するには「送信前チェック」ボタンを押してください。", false);
    } else if (r.status === "closed") {
      notify("確認画面が閉じられました。送信するには「送信前チェック」ボタンを押してください。", true);
    } else {
      notify(`確認画面を開けませんでした（エラー ${r.code}）。`, true);
    }
  } catch (e) {
    console.error("openConfirmDialog", e);
    notify("確認画面でエラーが発生しました。もう一度お試しください。", true);
  } finally {
    event.completed();
  }
}

async function saveAndSend(item, shownSignature) {

  try {
    if (!(await verifyAndSave(item, shownSignature))) {
      notify("確認後に宛先または添付が変更されました。もう一度確認してください。", true);
      return;
    }
  } catch (e) {
    console.error("verifyAndSave", e);
    notify("確認を保存できませんでした。もう一度お試しください。", true);
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
      notify("確認は保存しましたが、送信できませんでした。「送信」を押してください。", true);
      return;
    }
  }

  // 1.15 未対応の環境：確認だけ保存し、ユーザーに送信してもらう
  notify("確認を保存しました。メールの「送信」をもう一度押してください。", false);
}

// 結果はメール作成画面の通知バーに出す（150文字以内）
function notify(message, isError) {

  const type = Office.MailboxEnums.ItemNotificationMessageType;

  const details = isError
    ? { type: type.ErrorMessage, message }
    : { type: type.InformationalMessage, message, icon: "Icon.16x16", persistent: false };

  Office.context.mailbox.item.notificationMessages.replaceAsync("okanStatus", details, () => {});
}

Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
Office.actions.associate("openConfirmDialog", openConfirmDialog);
