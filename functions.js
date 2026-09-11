function onMessageSendHandler(event) {
    console.log("Send event fired");

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