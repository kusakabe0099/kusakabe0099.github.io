Office.onReady(() => {

    document
        .getElementById("testSave")
        .addEventListener(
            "click",
            testSessionSave
        );

    document
        .getElementById("testRead")
        .addEventListener(
            "click",
            testSessionRead
        );

});

function testSessionSave() {

    Office.context.mailbox.item.sessionData.setAsync(
        "test",
        "123",
        function(result) {

            console.log(
                "SAVE",
                result
            );

        }
    );

}

function testSessionRead() {

    Office.context.mailbox.item.sessionData.getAsync(
        "test",
        function(result) {

            console.log(
                "READ",
                result
            );

        }
    );

}
