Office.onReady(() => {

    document
        .getElementById("sendBtn")
        .addEventListener(
            "click",
            sendMail
        );

});

function sendMail() {

    Office.context.ui.messageParent(
        JSON.stringify({
            action: "CONFIRMED"
        })
    );

}
