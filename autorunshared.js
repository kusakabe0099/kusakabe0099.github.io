function onMessageSendHandler(event) {

    const confirmed =
        Office.context.roamingSettings.get(
            "sendCheckPassed"
        );

    if (!confirmed) {

        event.completed({
            allowEvent: false,
            errorMessage:
                "送信前チェックを完了してください"
        });

        return;
    }

    Office.context.roamingSettings.remove(
        "sendCheckPassed"
    );

    Office.context.roamingSettings.saveAsync();

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
