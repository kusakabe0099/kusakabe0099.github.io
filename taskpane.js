Office.onReady(() => {

    document
        .getElementById("openDialog")
        .addEventListener("click", openDialog);

});

function openDialog() {

    Office.context.ui.displayDialogAsync(
        "https://kusakabe0099.github.io/sendcheck.html",
        {
            height: 70,
            width: 60,
            displayInIframe: true
        },
        result => {

            console.log(result);

            if (
                result.status !==
                Office.AsyncResultStatus.Succeeded
            ) {

                console.error(result.error);

            }

        }
    );

}
