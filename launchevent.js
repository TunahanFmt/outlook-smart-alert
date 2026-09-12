/* global Office */

Office.onReady();

function onMessageSendHandler(event) {
    let isCompleted = false;

    // 3.5 saniye içinde işlem bitmezse mailin kilitlenmesini önlemek için otomatik tamamla
    const safetyTimeout = setTimeout(() => {
        if (!isCompleted) {
            isCompleted = true;
            event.completed({ allowEvent: true });
        }
    }, 3500);

    function safeComplete(args) {
        if (!isCompleted) {
            isCompleted = true;
            clearTimeout(safetyTimeout);
            event.completed(args);
        }
    }

    try {
        const mailbox = Office.context.mailbox;
        const userEmail = mailbox.userProfile ? mailbox.userProfile.emailAddress : "";
        const senderDomain = getDomain(userEmail);

        if (!senderDomain) {
            safeComplete({ allowEvent: true });
            return;
        }

        const item = mailbox.item;

        Promise.all([
            getRecipients(item.to),
            getRecipients(item.cc),
            getRecipients(item.bcc)
        ]).then((results) => {
            const allRecipients = results.flat();
            const externalRecipients = [];

            for (const email of allRecipients) {
                const domain = getDomain(email);
                if (domain && domain !== senderDomain) {
                    externalRecipients.push(email);
                }
            }

            if (externalRecipients.length > 0) {
                const uniqueList = [...new Set(externalRecipients)].join(", ");
                safeComplete({
                    allowEvent: false,
                    errorMessage: `GÜVENLİK UYARISI: Şirket dışı alıcı tespit edildi: [${uniqueList}]. Yalnızca @${senderDomain} adreslerine gönderim yapabilirsiniz.`,
                    sendModePrompt: Office.MailboxEnums.SendModePrompt.Block
                });
            } else {
                safeComplete({ allowEvent: true });
            }
        }).catch(() => {
            safeComplete({ allowEvent: true });
        });

    } catch (err) {
        safeComplete({ allowEvent: true });
    }
}

function getDomain(email) {
    if (!email || typeof email !== "string" || !email.includes("@")) return "";
    return email.split("@").pop().toLowerCase().trim();
}

function getRecipients(field) {
    return new Promise((resolve) => {
        if (!field || typeof field.getAsync !== "function") {
            resolve([]);
            return;
        }
        field.getAsync((result) => {
            if (result && result.status === Office.AsyncResultStatus.Succeeded && Array.isArray(result.value)) {
                const emails = result.value.map(r => r.emailAddress || r.address || "").filter(Boolean);
                resolve(emails);
            } else {
                resolve([]);
            }
        });
    });
}

Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
