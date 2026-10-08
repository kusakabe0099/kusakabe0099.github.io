Office.onReady(() => {

    document
        .getElementById("testSave")
        .addEventListener(
            "click",
            testPropsSave
        );

    document
        .getElementById("testRead")
        .addEventListener(
            "click",
            testSessionRead
        );

});

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

function testPropsSave() {

    Office.context.mailbox.item
        .loadCustomPropertiesAsync(
            function(result) {

                const props =
                    result.value;

                props.set(
                    "sendCheckPassed",
                    "true"
                );

                props.saveAsync(
                    function(r) {

                        console.log(
                            "SAVE",
                            r
                        );

                    }
                );

            }
        );

}
