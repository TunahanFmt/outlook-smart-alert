/* global Office */

/**
 * Office.js kütüphanesi hazır olduğunda tetiklenir
 */
Office.onReady(() => {
    // Office ortamı hazır
});

/**
 * E-posta Gönder butonuna basıldığında tetiklenen OnMessageSend olay işleyicisi
 * @param {Office.AddinCommands.Event} event - Office eklenti olayı
 */
function onMessageSendHandler(event) {
    try {
        const mailbox = Office.context.mailbox;
        
        // 1. Gönderen kullanıcının e-posta adresi ve domain bilgisi
        const userEmail = mailbox.userProfile ? mailbox.userProfile.emailAddress : "";
        const senderDomain = getDomainFromEmail(userEmail);

        if (!senderDomain) {
            // Gönderen domaini okunamadıysa güvenlik nedeniyle engelle
            event.completed({
                allowEvent: false,
                errorMessage: "Gönderen e-posta adresi doğrulanamadı. Lütfen oturumunuzu kontrol edin.",
                sendModePrompt: Office.MailboxEnums.SendModePrompt.Block
            });
            return;
        }

        const item = mailbox.item;

        // 2. TO, CC ve BCC alıcı listelerini eşzamanlı olarak al
        Promise.all([
            getRecipientsAsync(item.to),
            getRecipientsAsync(item.cc),
            getRecipientsAsync(item.bcc)
        ]).then(([toRecipients, ccRecipients, bccRecipients]) => {
            const allRecipients = [...toRecipients, ...ccRecipients, ...bccRecipients];

            if (allRecipients.length === 0) {
                // Alıcı yoksa gönderime izin ver (Outlook zaten uyaracaktır)
                event.completed({ allowEvent: true });
                return;
            }

            // 3. Gönderen domaininden farklı olan alıcıları filtrele
            const externalRecipients = [];

            for (const recipient of allRecipients) {
                const email = recipient.emailAddress || recipient.address || "";
                const recipientDomain = getDomainFromEmail(email);

                if (recipientDomain && recipientDomain !== senderDomain) {
                    externalRecipients.push(email);
                }
            }

            // 4. Farklı domain alıcısı bulunduysa gönderimi engelle
            if (externalRecipients.length > 0) {
                const uniqueExternal = [...new Set(externalRecipients)];
                const recipientListStr = uniqueExternal.join(", ");

                event.completed({
                    allowEvent: false,
                    errorMessage: `GÜVENLİK UYARISI: Şirket dışı alıcı tespit edildi: [${recipientListStr}]. Yalnızca @${senderDomain} uzantılı adreslere mail gönderebilirsiniz.`,
                    sendModePrompt: Office.MailboxEnums.SendModePrompt.Block
                });
            } else {
                // Tüm alıcılar şirket içi domain ile eşleşiyor
                event.completed({ allowEvent: true });
            }
        }).catch((error) => {
            // Beklenmeyen bir hata durumunda güvenli tarafta kalıp gönderimi engelle
            console.error("Alıcı kontrol hatası:", error);
            event.completed({
                allowEvent: false,
                errorMessage: "Alıcı adresleri doğrulanırken bir sistem hatası oluştu. Mail gönderilemedi.",
                sendModePrompt: Office.MailboxEnums.SendModePrompt.Block
            });
        });

    } catch (err) {
        console.error("onMessageSendHandler çalışma hatası:", err);
        event.completed({
            allowEvent: false,
            errorMessage: "Eklenti çalışırken bir hata oluştu.",
            sendModePrompt: Office.MailboxEnums.SendModePrompt.Block
        });
    }
}

/**
 * E-posta adresinden domain bilgisini ayıklar
 * @param {string} email 
 * @returns {string} Domain adı (küçük harflerle)
 */
function getDomainFromEmail(email) {
    if (!email || typeof email !== "string" || !email.includes("@")) {
        return "";
    }
    const parts = email.split("@");
    return parts[parts.length - 1].toLowerCase().trim();
}

/**
 * Outlook alıcı alanındaki (TO/CC/BCC) kişileri asenkron olarak çeker
 * @param {Office.Recipients} recipientField 
 * @returns {Promise<Array>} Alıcı nesneleri dizisi
 */
function getRecipientsAsync(recipientField) {
    return new Promise((resolve) => {
        if (!recipientField || typeof recipientField.getAsync !== "function") {
            resolve([]);
            return;
        }
        recipientField.getAsync((result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded && result.value) {
                resolve(result.value);
            } else {
                resolve([]);
            }
        });
    });
}

// Fonksiyonu Office Add-in olay eşlemesine kaydet
if (typeof Office !== "undefined" && Office.actions) {
    Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
}
