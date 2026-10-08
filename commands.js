function openSendCheck(event) {

    Office.context.ui.displayDialogAsync(
        "https://kusakabe0099.github.io/sendcheck.html",
        {
            width: 70,
            height: 80,
            displayInIframe: true
        }
    );

    event.completed();
}
