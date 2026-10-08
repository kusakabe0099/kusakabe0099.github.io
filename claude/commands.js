/* 送信時イベント（OnMessageSend）と、確認ダイアログを開くコマンド。
   Smart Alerts の「確認画面を開く」ボタンは、タスクペインではなくこのファイルの
   openConfirmDialog（manifest.xml の ExecuteFunction）を実行するので、タスクペインは表示されない。 */

Office.onReady();

const CONFIRM_COMMAND_ID = "msgComposeConfirmButton"; // manifest.xml のボタンIDと一致させる

let dialog = null;
let shownSignature = null; // ダイアログに表示した内容の署名
let pendingEvent = null;   // 実行中のコマンドのイベント（ダイアログが終わるまで保持する）

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

    event.completed({
      allowEvent: false,
      errorMessage:
        "送信前に宛先と添付ファイルの確認が必要です。" +
        `（宛先 ${state.recipients.length} 件、うち社外 ${state.externalCount} 件、添付 ${state.files.length} 件）`,
      cancelLabel: "確認画面を開く", // 「アクションを実行」ボタンのラベル（20文字以内）
      commandId: CONFIRM_COMMAND_ID,
    });
  } catch (e) {
    event.completed({
      allowEvent: false,
      errorMessage: "送信前チェックでエラーが発生しました。もう一度「送信」を押してください。",
    });
  }
}

/* ---------- 確認ダイアログ ---------- */

// 「確認画面を開く」（Smart Alerts のアクション / リボンボタン）
async function openConfirmDialog(event) {

  if (dialog) {
    event.completed(); // すでに開いている
    return;
  }

  pendingEvent = event;

  const item = Office.context.mailbox.item;
  let state, subject;

  try {
    [state, subject] = await Promise.all([
      Okan.collect(item),
      Okan.call(cb => item.subject.getAsync(cb)),
    ]);
  } catch (e) {
    notify("宛先と添付ファイルを読み込めませんでした。もう一度お試しください。", true);
    finishCommand();
    return;
  }

  shownSignature = state.signature;

  const payload = {
    subject,
    recipients: state.recipients,
    files: state.files,
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
        notify(
          code === 12009
            ? "確認画面の表示が許可されませんでした。「送信前チェック」ボタンから開いてください。"
            : `確認画面を開けませんでした（エラー ${code}）。`,
          true
        );
        finishCommand();
        return;
      }

      dialog = result.value;
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, onDialogMessage);
      dialog.addEventHandler(Office.EventType.DialogEventReceived, onDialogEvent);
    }
  );
}

// ダイアログが×ボタンなどで閉じられた、または読み込みに失敗した
function onDialogEvent(arg) {

  dialog = null;

  notify(
    arg.error === 12006
      ? "確認画面が閉じられました。送信するには「送信前チェック」ボタンを押してください。"
      : `確認画面でエラーが発生しました（エラー ${arg.error}）。`,
    true
  );

  finishCommand();
}

async function onDialogMessage(arg) {

  let msg;

  try {
    msg = JSON.parse(arg.message);
  } catch (e) {
    return;
  }

  closeDialog();

  if (msg.type === "confirm") {
    await saveAndSend();
  } else {
    notify("キャンセルしました。送信するには「送信前チェック」ボタンを押してください。", false);
  }

  finishCommand();
}

async function saveAndSend() {

  const item = Office.context.mailbox.item;

  try {
    // 確認画面を見せたあとに宛先・添付が変わっていないか再確認してから保存する
    const latest = await Okan.collect(item);

    if (latest.signature !== shownSignature) {
      notify("確認後に宛先または添付が変更されました。もう一度確認してください。", true);
      return;
    }

    // roamingSettings はメールボックス全体で永続するため使わない。
    // このメール（作成セッション）にだけ有効な sessionData に、署名を保存する。
    await Okan.setConfirmed(item, latest.signature);

  } catch (e) {
    console.error("setConfirmed", e);
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

function closeDialog() {

  if (!dialog) return;

  try {
    dialog.close();
  } catch (e) {
    // すでに閉じている場合は何もしない
  }
  dialog = null;
}

// コマンドの実行を終える（ダイアログが終わるまで event.completed を呼ばずに待つ）
function finishCommand() {

  if (!pendingEvent) return;

  pendingEvent.completed();
  pendingEvent = null;
}

// タスクペインがないので、結果はメール作成画面の通知バーに出す（150文字以内）
function notify(message, isError) {

  const type = Office.MailboxEnums.ItemNotificationMessageType;

  const details = isError
    ? { type: type.ErrorMessage, message }
    : { type: type.InformationalMessage, message, icon: "Icon.16x16", persistent: false };

  Office.context.mailbox.item.notificationMessages.replaceAsync("okanStatus", details, () => {});
}

Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
Office.actions.associate("openConfirmDialog", openConfirmDialog);
