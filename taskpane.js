Office.onReady(() => {

    document
        .getElementById("openDialog")
        .addEventListener("click", openDialog);

});

function openDialog() {

    Office.context.ui.displayDialogAsync(
        "https://kusakabe0099.github.io/sendcheck.html",
        {
            height: 80,
            width: 70,
            displayInIframe: true
        },
        function (asyncResult) {
    
            const dialog = asyncResult.value;
    
            dialog.addEventHandler(
                Office.EventType.DialogMessageReceived,
                function (arg) {
    
                    console.log(
                        "dialog message=",
                        arg.message
                    );
    
                }
            );
    
        }
    );

}
