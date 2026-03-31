const ALLOWED_DOMAIN = "test.dd";
const NOTIFICATION_ID = "externalRecipientsWarning";

Office.onReady(() => {
    window.onRecipientsChanged = onRecipientsChanged;
});

function onRecipientsChanged(event) {
    const item = Office.context.mailbox.item;

    Promise.all([
        getRecipients(item.to),
        getRecipients(item.cc),
        getRecipients(item.bcc)
    ])
    .then(([to, cc, bcc]) => {
        const all = [...to, ...cc, ...bcc];

        const external = all.filter(r =>
            isExternal(r.emailAddress)
        );

        if (external.length > 0) {
            showWarning(external);
        } else {
            clearWarning();
        }
    })
    .catch(console.error)
    .finally(() => event.completed());
}

function getRecipients(field) {
    return new Promise((resolve) => {
        field.getAsync(result => {
            resolve(result.value || []);
        });
    });
}

function isExternal(email) {
    if (!email) return false;
    return !email.toLowerCase().endsWith("@" + ALLOWED_DOMAIN);
}

function showWarning(externals) {
    const emails = externals.map(e => e.emailAddress).join(", ");

    Office.context.mailbox.item.notificationMessages.replaceAsync(NOTIFICATION_ID, {
        type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
        message: `Uwaga! Adresaci spoza domeny: ${emails}`,
        persistent: true
    });
}

function clearWarning() {
    Office.context.mailbox.item.notificationMessages.removeAsync(NOTIFICATION_ID);
}
