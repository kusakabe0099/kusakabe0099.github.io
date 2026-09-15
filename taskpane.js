Office.onReady(() => {

    const item = Office.context.mailbox.item;

    Promise.all([
        checkTo(item),
        checkCc(item),
        checkBcc(item),
        checkSubject(item),
        checkAttachment(item)
    ]).then(updateButton);

});

function checkTo(item) {

    return new Promise(resolve => {

        item.to.getAsync(result => {

            const ok =
                result.status === Office.AsyncResultStatus.Succeeded &&
                result.value.length > 0;

            document.getElementById("toCheck").checked = ok;

            resolve();
        });

    });
}

function checkCc(item) {

    return new Promise(resolve => {

        item.cc.getAsync(result => {

            const ok =
                result.status === Office.AsyncResultStatus.Succeeded &&
                result.value.length > 0;

            document.getElementById("ccCheck").checked = ok;

            resolve();
        });

    });
}

function checkBcc(item) {

    return new Promise(resolve => {

        item.bcc.getAsync(result => {

            const ok =
                result.status === Office.AsyncResultStatus.Succeeded &&
                result.value.length > 0;

            document.getElementById("bccCheck").checked = ok;

            resolve();
        });

    });
}

function checkSubject(item) {

    return new Promise(resolve => {

        item.subject.getAsync(result => {

            const ok =
                result.status === Office.AsyncResultStatus.Succeeded &&
                result.value.trim() !== "";

            document.getElementById("subjectCheck").checked = ok;

            resolve();
        });

    });
}

function checkAttachment(item) {

    const ok = item.attachments.length > 0;

    document.getElementById("attachCheck").checked = ok;

    return Promise.resolve();
}

function updateButton() {

    const checks = document.querySelectorAll(
        "input[type='checkbox']"
    );

    const allOk =
        Array.from(checks).every(c => c.checked);

    const button =
        document.getElementById("sendButton");

    button.disabled = !allOk;

    button.textContent =
        allOk ? "送信OK" : "確認不足";
}
