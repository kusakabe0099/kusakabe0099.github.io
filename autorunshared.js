function onMessageSendHandler(event) {

    const confirmed =
        Office.context.roamingSettings.get(
            "sendCheckPassed"
        );
    
    Office.context.ui.displayDialogAsync(
        "https://kusakabe0099.github.io/sendcheck.html",
        { height: 70, width: 60 }
    );

    console.log("confirmed=", confirmed);

    if (!confirmed) {

        event.completed({
            allowEvent: false,
            errorMessage:
                "送信確認画面を開いて確認してください。"
        });

        return;
    }

    event.completed({
        allowEvent: true
    });
}

Office.onReady(() => {

    Office.actions.associate(
        "onMessageSendHandler",
        onMessageSendHandler
    );

});
