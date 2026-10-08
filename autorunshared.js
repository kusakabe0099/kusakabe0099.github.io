function onMessageSendHandler(event) {

    Office.context.mailbox.item
        .loadCustomPropertiesAsync(function (result) {

            if (
                result.status !==
                Office.AsyncResultStatus.Succeeded
            ) {

                event.completed({
                    allowEvent: false,
                    errorMessage: "送信前確認が必要です"
                });

                return;
            }

            const props = result.value;

            const checked =
                props.get("sendCheckPassed");

            if (checked === "true") {

                props.remove("sendCheckPassed");

                props.saveAsync(function () {

                    event.completed({
                        allowEvent: true
                    });

                });

                return;
            }

            event.completed({
                allowEvent: false,
                errorMessage:
                  "送信前チェックを実施してください"
            });

        });

}

Office.onReady(() => {

    Office.actions.associate(
        "onMessageSendHandler",
        onMessageSendHandler
    );

});
