function onMessageSendHandler(event) {

    console.log(
        "SEND EVENT START"
    );

    try {

        Office.context.mailbox.item.sessionData.getAsync(
            "test",
            function(result) {

                console.log(
                    "SESSION",
                    result
                );

                event.completed({
                    allowEvent: false,
                    errorMessage:
                        "値を確認してください"
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

Office.onReady(() => {

    Office.actions.associate(
        "onMessageSendHandler",
        onMessageSendHandler
    );

});
