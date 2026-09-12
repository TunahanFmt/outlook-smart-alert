/* global Office */

console.log("[SmartAlert] Script yüklendi ve çalışma ortamı hazır.");

Office.onReady((info) => {
    console.log("[SmartAlert] Office.onReady tetiklendi. Host:", info.host, "Platform:", info.platform);
});

function onMessageSendHandler(event) {
    console.log("[SmartAlert] === OnMessageSendHandler Tetiklendi ===");
    let isCompleted = false;

    // 3.5 saniyelik güvenlik zaman aşımı
    const safetyTimeout = setTimeout(() => {
        if (!isCompleted) {
            console.warn("[SmartAlert] ⚠️ ZAMAN AŞIMI: 3.5s içinde yanıt alınamadı. Mail gönderimine otomatik izin veriliyor.");
            isCompleted = true;
            event.completed({ allowEvent: true });
        }
    }, 3500);

    function safeComplete(args, reason) {
        console.log(`[SmartAlert] safeComplete çağrıldı. Nedeni: [${reason}] | Tamamlanma Durumu: ${isCompleted}`);
        if (!isCompleted) {
            isCompleted = true;
            clearTimeout(safetyTimeout);
            console.log("[SmartAlert] event.completed() çalıştırılıyor. Gönderilen parametreler:", JSON.stringify(args));
            event.completed(args);
        } else {
            console.warn("[SmartAlert] ⚠️ DİKKAT: event.completed() zaten çağrılmıştı, mükerrer çağrı engellendi.");
        }
    }

    try {
        const mailbox = Office.context.mailbox;
        console.log("[SmartAlert] Mailbox nesnesi:", mailbox ? "Mevcut" : "NULL/Undefined");

        const userEmail = mailbox && mailbox.userProfile ? mailbox.userProfile.emailAddress : "";
        console.log("[SmartAlert] Gönderen e-posta adresi:", userEmail || "Bulunamadı");

        const senderDomain = getDomain(userEmail);
        console.log("[SmartAlert] Tespit edilen gönderen domaini:", senderDomain || "Bulunamadı");

        if (!senderDomain) {
            console.warn("[SmartAlert] Gönderen domaini okunamadığı için kontrol atlanıyor.");
            safeComplete({ allowEvent: true }, "Gönderen domain yok");
            return;
        }

        const item = mailbox.item;
        console.log("[SmartAlert] Alıcı adresleri (To, CC, BCC) çekiliyor...");

        Promise.all([
            getRecipients(item.to, "To"),
            getRecipients(item.cc, "Cc"),
            getRecipients(item.bcc, "Bcc")
        ]).then((results) => {
            const allRecipients = results.flat();
            console.log("[SmartAlert] Çekilen tüm alıcı listesi:", allRecipients);

            const externalRecipients = [];
            for (const email of allRecipients) {
                const domain = getDomain(email);
                console.log(`[SmartAlert] Alıcı Analizi -> E-posta: ${email} | Domain: ${domain}`);
                if (domain && domain !== senderDomain) {
                    externalRecipients.push(email);
                }
            }

            console.log("[SmartAlert] Harici (dış) alıcılar:", externalRecipients);

            if (externalRecipients.length > 0) {
                const uniqueList = [...new Set(externalRecipients)].join(", ");
                
                // Enum ve nesne kontrolleri için loglar
                console.log("[SmartAlert] Office.MailboxEnums nesnesi:", Office.MailboxEnums);
                let blockMode = 1; // Fallback SendModePrompt.Block değeri

                if (Office.MailboxEnums && Office.MailboxEnums.SendModePrompt) {
                    console.log("[SmartAlert] Office.MailboxEnums.SendModePrompt nesnesi mevcut:", Office.MailboxEnums.SendModePrompt);
                    blockMode = Office.MailboxEnums.SendModePrompt.Block;
                } else {
                    console.warn("[SmartAlert] ⚠️ Office.MailboxEnums.SendModePrompt bulunamadı! Yedek değer (1) kullanılıyor.");
                }

                console.log(`[SmartAlert] 🛑 ENGELLEME KARARI VERİLDİ. Kullanılan blockMode: ${blockMode}`);
                safeComplete({
                    allowEvent: false,
                    errorMessage: `GÜVENLİK UYARISI: Şirket dışı alıcı tespit edildi: [${uniqueList}]. Yalnızca @${senderDomain} adreslerine gönderim yapabilirsiniz.`,
                    sendModePrompt: blockMode
                }, "Harici alıcı tespit edildi");
            } else {
                console.log("[SmartAlert] ✅ TÜM ALICILAR İÇ DOMAİN. Gönderime izin veriliyor.");
                safeComplete({ allowEvent: true }, "Tüm alıcılar iç domain");
            }
        }).catch((err) => {
            console.error("[SmartAlert] ❌ Promise.all aşamasında hata yakalandı:", err);
            safeComplete({ allowEvent: true }, "Promise.all hatası");
        });

    } catch (err) {
        console.error("[SmartAlert] ❌ Ana try-catch bloğunda kritik hata:", err);
        safeComplete({ allowEvent: true }, "Ana try-catch hatası");
    }
}

function getDomain(email) {
    if (!email || typeof email !== "string" || !email.includes("@")) return "";
    return email.split("@").pop().toLowerCase().trim();
}

function getRecipients(field, fieldName) {
    return new Promise((resolve) => {
        if (!field || typeof field.getAsync !== "function") {
            console.log(`[SmartAlert] Alıcı alanı [${fieldName}] boş veya getAsync desteklemiyor.`);
            resolve([]);
            return;
        }
        console.log(`[SmartAlert] [${fieldName}] alanı için getAsync çağrılıyor...`);
        field.getAsync((result) => {
            console.log(`[SmartAlert] [${fieldName}] getAsync yanıtı alındı. Durum: ${result ? result.status : "null"}`);
            if (result && result.status === Office.AsyncResultStatus.Succeeded && Array.isArray(result.value)) {
                const emails = result.value.map(r => r.emailAddress || r.address || "").filter(Boolean);
                console.log(`[SmartAlert] [${fieldName}] alanından ayıklanan e-postalar:`, emails);
                resolve(emails);
            } else {
                console.warn(`[SmartAlert] [${fieldName}] alanı okunamadı veya boş:`, result ? result.error : "Sonuç yok");
                resolve([]);
            }
        });
    });
}

Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
