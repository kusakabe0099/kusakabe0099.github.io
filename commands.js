function openSendCheck(event) {

    Office.context.ui.displayDialogAsync(
        "https://kusakabe0099.github.io/sendcheck.html",
        {
            width: 70,
            height: 80,
            displayInIframe: true
        },
        function (result) {

            console.log(result);

        }
    );

    event.completed();
}

Office.onReady(() => {

    Office.actions.associate(
        "openSendCheck",
        openSendCheck
    );

});
