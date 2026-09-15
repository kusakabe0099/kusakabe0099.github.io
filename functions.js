function onMessageSendHandler(event) {

    console.log("OnMessageSend fired");

    event.completed({
        allowEvent: true
    });
}

Office.onReady(() => {

    Office.actions.associate(
        "onMessageSendHandler",
        onMessageSendHandler
    );

});
