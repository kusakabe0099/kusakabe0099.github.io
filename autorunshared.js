function onMessageSendHandler(event) {

    Office.context.mailbox.item.sessionData.getAsync(
        "sendCheckPassed",
        function(result) {

            console.log(
                "sessionData",
                result
            );

            if (
                result.status ===
                Office.AsyncResultStatus.Succeeded &&
                result.value === "true"
            ) {

                event.completed({
                    allowEvent: true
                });

                return;
            }

            event.completed({
                allowEvent: false,
                errorMessage:
                    "未確認です"
            });

        }
    );

}

Office.onReady(() => {

    Office.actions.associate(
        "onMessageSendHandler",
        onMessageSendHandler
    );

});
