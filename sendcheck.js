function sendMail() {

    Office.context.ui.messageParent(
        JSON.stringify({
            action: "CONFIRMED"
        })
    );

}
