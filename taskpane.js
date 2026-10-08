let currentDialog = null;

function openDialog() {

    Office.context.ui.displayDialogAsync(
        "https://kusakabe0099.github.io/sendcheck.html",
        {
            height: 80,
            width: 70,
            displayInIframe: true
        },
        function (result) {

            if (
                result.status !==
                Office.AsyncResultStatus.Succeeded
            ) {

                console.error(result.error);
                return;
            }

            currentDialog = result.value;

            currentDialog.addEventHandler(
                Office.EventType.DialogMessageReceived,
                function (arg) {

                    const msg =
                        JSON.parse(arg.message);

                    console.log(
                        "dialog message=",
                        msg
                    );

                    if (
                        msg.action ===
                        "CONFIRMED"
                    ) {

                        Office.context.mailbox.item.sessionData.setAsync(
                            "sendCheckPassed",
                            "true",
                            function(result) {

                                console.log(
                                    "sessionData set",
                                    result.status
                                );

                                currentDialog.close();

                            }
                        );

                    }

                }
            );

        }
    );

}
