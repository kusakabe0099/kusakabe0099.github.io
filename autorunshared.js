function onMessageSendHandler(event) {

    console.log("SEND START");

    try {

        console.log(
            "Office.context.ui",
            Office.context.ui
        );

        Office.context.ui.displayDialogAsync(
            "https://kusakabe0099.github.io/sendcheck.html",
            {
                height: 80,
                width: 70,
                displayInIframe: true
            },
            function(result) {

                console.log(
                    "displayDialogAsync result",
                    result
                );

                if (
                    result.status ===
                    Office.AsyncResultStatus.Succeeded
                ) {

                    event.completed({
                        allowEvent: false,
                        errorMessage:
                            "Dialog OPEN SUCCESS"
                    });

                } else {

                    event.completed({
                        allowEvent: false,
                        errorMessage:
                            "Dialog OPEN FAILED"
                    });

                }

            }
        );

    } catch(ex) {

        console.error(ex);

        event.completed({
            allowEvent: false,
            errorMessage:
                ex.message
        });

    }

}

Office.onReady(() => {

    Office.actions.associate(
        "onMessageSendHandler",
        onMessageSendHandler
    );

});
