function onMessageSendHandler(event) {
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