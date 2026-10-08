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
                    Office.AsyncResultStatus.Failed
                ) {
        
                    console.log(
                        "error",
                        result.error
                    );
        
                    console.log(
                        "code",
                        result.error.code
                    );
        
                    console.log(
                        "message",
                        result.error.message
                    );
        
                }
        
                event.completed({
                    allowEvent: false,
                    errorMessage: "検証中"
                });
        
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

function openSendCheck(event) {

    Office.context.ui.displayDialogAsync(
        "https://kusakabe0099.github.io/sendcheck.html",
        {
            width: 70,
            height: 80,
            displayInIframe: true
        }
    );

    event.completed();
}

Office.onReady(() => {

    Office.actions.associate(
        "openSendCheck",
        openSendCheck
    );

    Office.actions.associate(
        "onMessageSendHandler",
        onMessageSendHandler
    );

});
