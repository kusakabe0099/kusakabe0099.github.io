function onMessageSendHandler(event) {

    Office.context.mailbox.item
        .loadCustomPropertiesAsync(
            function(result) {

                console.log(
                    "LOAD",
                    result
                );

                if (
                    result.status !==
                    Office.AsyncResultStatus.Succeeded
                ) {

                    event.completed({
                        allowEvent: false,
                        errorMessage:
                            "LOAD FAILED"
                    });

                    return;
                }

                const props =
                    result.value;

                const flag =
                    props.get(
                        "sendCheckPassed"
                    );

                console.log(
                    "FLAG",
                    flag
                );

                event.completed({
                    allowEvent: false,
                    errorMessage:
                        `FLAG=${flag}`
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
