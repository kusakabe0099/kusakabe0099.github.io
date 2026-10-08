Office.onReady(() => {

    document
        .getElementById("openDialog")
        .addEventListener(
            "click",
            openDialog
        );

});

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

                console.error(
                    result.error
                );

                return;
            }

            const dialog =
                result.value;

            dialog.addEventHandler(
                Office.EventType.DialogMessageReceived,
                function (arg) {

                    const msg =
                        JSON.parse(
                            arg.message
                        );

                    if (
                        msg.action ===
                        "CONFIRMED"
                    ) {

                        Office.context
                            .mailbox
                            .item
                            .loadCustomPropertiesAsync(
                                function (r) {

                                    const props =
                                        r.value;

                                    props.set(
                                        "sendCheckPassed",
                                        "true"
                                    );

                                    props.saveAsync(
                                        function () {

                                            dialog.close();

                                            alert(
                                                "確認完了。送信してください。"
                                            );

                                        }
                                    );

                                }
                            );

                    }

                }
            );

        }
    );

}
