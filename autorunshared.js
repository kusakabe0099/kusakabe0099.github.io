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
                    Office.AsyncResultStatus
                        .Succeeded &&
                result.value === "true"
            ) {

                Office.context.mailbox.item.sessionData.removeAsync(
                    "sendCheckPassed",
                    function() {

                        event.completed({
                            allowEvent: true
                        });

                    }
                );

                return;
            }

            event.completed({
                allowEvent: false,
                errorMessage:
                    "送信前確認が未完了です"
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
