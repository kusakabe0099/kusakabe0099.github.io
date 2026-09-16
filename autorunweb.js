function onMessageSendHandler(event) {

    Office.context.mailbox.item.subject.getAsync(res => {

        const subject =
            res.value || "";

        if (subject.trim() === "") {

            event.completed({
                allowEvent: false,
                errorMessage: "件名が未入力です"
            });

            return;
        }

        event.completed({
            allowEvent: true
        });

    });

}

Office.onReady(() => {

    Office.actions.associate(
        "onMessageSendHandler",
        onMessageSendHandler
    );

});