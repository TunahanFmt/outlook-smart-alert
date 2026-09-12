/* global Office */

console.log("[SmartAlert] Hızlı Tarama Modu Yüklendi.");

Office.onReady();

function onMessageSendHandler(event) {
    const startTime = performance.now();
    console.log("[SmartAlert] 🚀 Gönderim kontrolü başlatıldı.");
    let isCompleted = false;

    // 2.5 saniyelik sıkı güvenlik zaman aşımı (Mail kilitlenmesini tamamen önler)
    const safetyTimeout = setTimeout(() => {
        if (!isCompleted) {
            console.warn("[SmartAlert] ⚠️ ZAMAN AŞIMI: Güvenlik sınırı aşıldı, gönderime izin veriliyor.");
            isCompleted = true;
            event.completed({ allowEvent: true });
        }
    }, 2500);

    function safeComplete(args, reason) {
        if (!isCompleted) {
            isCompleted = true;
            clearTimeout(safetyTimeout);
            const duration = (performance.now() - startTime).toFixed(2);
            console.log(`[SmartAlert] ⏱️ İşlem Tamamlandı (${duration} ms). Sonuç: [${reason}]`);
            console.log("[SmartAlert] Yanıt:", JSON.stringify(args));
            event.completed(args);
        }
    }

    try {
        const mailbox = Office.context.mailbox;
        const userEmail = (mailbox && mailbox.userProfile) ? mailbox.userProfile.emailAddress : "";
        const senderDomain = getDomain(userEmail);

        console.log(`[SmartAlert] Gönderen: ${userEmail} | Domain: ${senderDomain}`);

        if (!senderDomain) {
            console.warn("[SmartAlert] Gönderen domaini okunamadı, kontrol atlanıyor.");
            safeComplete({ allowEvent: true }, "Gönderen domain yok");
            return;
        }

        const item = mailbox.item;

        // TO, CC, BCC alanlarını eşzamanlı (paralel) en hızlı şekilde oku
        Promise.all([
            getRecipientsFast(item.to, "To"),
            getRecipientsFast(item.cc, "Cc"),
            getRecipientsFast(item.bcc, "Bcc")
        ]).then((results) => {
            const externalRecipients = new Set();

            // Performans için tek geçişli düz döngü
            for (let i = 0; i < results.length; i++) {
                const list = results[i];
                for (let j = 0; j < list.length; j++) {
                    const email = list[j];
                    const domain = getDomain(email);
                    if (domain && domain !== senderDomain) {
                        externalRecipients.add(email);
                    }
                }
            }

            if (externalRecipients.size > 0) {
                const blockedList = Array.from(externalRecipients).join(", ");
                console.log(`[SmartAlert] 🛑 ENGELLEME: Harici alıcılar tespit edildi -> [${blockedList}]`);

                safeComplete({
                    allowEvent: false,
                    errorMessage: `GÜVENLİK UYARISI: Şirket dışı alıcı tespit edildi: [${blockedList}]. Yalnızca @${senderDomain} adreslerine gönderim yapabilirsiniz.`
                }, "Harici alıcı engellendi");
            } else {
                console.log("[SmartAlert] ✅ BAŞARILI: Tüm alıcılar kurum içi.");
                safeComplete({ allowEvent: true }, "Tüm alıcılar iç domain");
            }
        }).catch((err) => {
            console.error("[SmartAlert] ❌ Alıcı okuma hatası:", err);
            safeComplete({ allowEvent: true }, "Okuma hatası");
        });

    } catch (err) {
        console.error("[SmartAlert] ❌ Kritik işlem hatası:", err);
        safeComplete({ allowEvent: true }, "Kritik hata");
    }
}

// Hızlı Domain Ayıklama
function getDomain(email) {
    if (!email || typeof email !== "string") return "";
    const atIndex = email.lastIndexOf("@");
    return atIndex !== -1 ? email.substring(atIndex + 1).toLowerCase().trim() : "";
}

// Ultra Hızlı Asenkron Alıcı Okuyucu
function getRecipientsFast(field, fieldName) {
    return new Promise((resolve) => {
        if (!field || typeof field.getAsync !== "function") {
            resolve([]);
            return;
        }
        field.getAsync((result) => {
            if (result && result.status === Office.AsyncResultStatus.Succeeded && Array.isArray(result.value)) {
                const len = result.value.length;
                const emails = new Array(len);
                let validCount = 0;

                for (let i = 0; i < len; i++) {
                    const item = result.value[i];
                    const addr = (typeof item === "string") ? item : (item.emailAddress || item.address || "");
                    if (addr) {
                        emails[validCount++] = addr;
                    }
                }
                emails.length = validCount; // Diziyi gerçek eleman sayısına kırp
                console.log(`[SmartAlert] [${fieldName}] okundu: ${validCount} adres`);
                resolve(emails);
            } else {
                resolve([]);
            }
        });
    });
}

Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
