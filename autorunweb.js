function onMessageSendHandler(event) {

    console.log("SEND EVENT");

    event.completed({
        allowEvent: false,
        errorMessage: "送信イベントが発火しました"
    });

}

Office.onReady(() => {

    Office.actions.associate(
        "onMessageSendHandler",
        onMessageSendHandler
    );

});
``
