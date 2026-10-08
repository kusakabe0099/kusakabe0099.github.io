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
            
                    const msg =
                        JSON.parse(arg.message);
            
                    if (msg.action === "SEND") {
            
                        console.log(
                            "mailbox",
                            Office.context.mailbox
                        );
            
                        console.log(
                            "item",
                            Office.context.mailbox.item
                        );
            
                        console.log(
                            "sendAsync",
                            typeof Office.context.mailbox.item.sendAsync
                        );
            
                    }
            
                }
            );
    
        }
    );

}
