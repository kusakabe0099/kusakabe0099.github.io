Office.onReady(() => {

    document
        .getElementById(
            "confirmBtn"
        )
        .addEventListener(
            "click",
            confirmSend
        );

});

function confirmSend() {

    Office.context.ui.messageParent(
        JSON.stringify({
            action:
                "CONFIRMED"
        })
    );

}
