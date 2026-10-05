Office.onReady(() => {
    
    console.log("Office");
    
    console.log(Office.context);
    
    console.log(Office.context.mailbox);
    
    document
        .getElementById("sendBtn")
        .addEventListener(
            "click",
            sendMail
        );

});

function sendMail() {

    const item =
        Office.context.mailbox.item;

    console.log(item);

    item.sendAsync(result => {

        console.log(result);

        if (
            result.status ===
            Office.AsyncResultStatus.Succeeded
        ) {

            console.log(
                "SEND SUCCESS"
            );

        } else {

            console.error(
                result.error
            );

        }

    });

}
