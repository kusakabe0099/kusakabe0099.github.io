Office.onReady(() => {
    
    console.log("Office");
    
    console.log(Office.context);
    
    console.log(Office.context.mailbox);
    
    document
        .getElementById("sendBtn")
        .addEventListener("click", () => {

            Office.context.ui.messageParent(
                JSON.stringify({
                    action: "SEND"
                })
            );

        });


});
