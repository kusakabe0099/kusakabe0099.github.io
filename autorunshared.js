function onMessageSendHandler(event) {

    const confirmed =
        Office.context.roamingSettings.get(
            "sendCheckPassed"
        );

    console.log("confirmed=", confirmed);

    if (!confirmed) {

        event.completed({
            allowEvent: false,
            errorMessage:
                "送信前チェックを完了してください"
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
