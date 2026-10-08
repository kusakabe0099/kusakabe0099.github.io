function onMessageSendHandler(event) {

    console.log("SEND START");

    console.log(
        "sessionData",
        Office.context.mailbox.item.sessionData
    );

    console.log(
        "customProperties",
        typeof Office.context.mailbox.item.loadCustomPropertiesAsync
    );

    event.completed({
        allowEvent: false,
        errorMessage: "ログ確認"
    });
}

Office.onReady(() => {

    Office.actions.associate(
        "onMessageSendHandler",
        onMessageSendHandler
    );

});
