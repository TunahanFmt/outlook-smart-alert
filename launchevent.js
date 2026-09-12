/* global Office */

/**
 * ============================================================================
 * OUTLOOK SMART ALERT - DOKÜMAN VE FİRMA KONTROL EKLENTİSİ
 * ============================================================================
 * Kurallar:
 * 1. Şahsi veya Shared hesaplardan atılan mailde (To, CC, BCC) aynı anda 2 FARKLI müşteri firması bulunamaz.
 * 2. Domain uzantıları (.com, .com.tr vb.) bağımsız marka köküne göre kontrol yapılır.
 * ============================================================================
 */

console.log("[SmartAlert] Kurumsal Mail Güvenlik ve Eşleşme Modülü Yüklendi.");

Office.onReady();

// ============================================================================
// ⚙️ CONFIGURATION / YAPILANDIRMA ALANI
// ============================================================================

// 1. Kurum İçi Domain
const INTERNAL_DOMAIN = "fmtturkey.com";

// 2. Tanımlı Müşteri / Partner Firma Listesi
// Uzantı yazsanız bile (.com / .com.tr) sistem otomatik olarak marka kökünü ("akcansa", "onduline") alır.
const CLIENT_DOMAINS = [
    "akcansa.com.tr",
    "cci.com.tr",
    "abdiibrahim.com.tr",
    "onduline.com.tr",
    "allianz.com.tr",
    "roche.com"
];

// ============================================================================
// 🚀 ANA TETİKLEYİCİ FONKSİYON (OnMessageSend)
// ============================================================================

function onMessageSendHandler(event) {
    const startTime = performance.now();
    console.log("[SmartAlert] 🚀 Gönderim kontrolü başlatıldı.");
    let isCompleted = false;

    // 3.0 saniyelik güvenlik zaman aşımı
    const safetyTimeout = setTimeout(() => {
        if (!isCompleted) {
            console.warn("[SmartAlert] ⚠️ ZAMAN AŞIMI: 3.0s içinde yanıt alınamadı. Güvenlik nedeniyle gönderime izin veriliyor.");
            isCompleted = true;
            event.completed({ allowEvent: true });
        }
    }, 3000);

    function safeComplete(args, reason) {
        if (!isCompleted) {
            isCompleted = true;
            clearTimeout(safetyTimeout);
            const duration = (performance.now() - startTime).toFixed(2);
            console.log(`[SmartAlert] ⏱️ İşlem Tamamlandı (${duration} ms). Sonuç: [${reason}]`);
            console.log("[SmartAlert] Yanıt Detayı:", JSON.stringify(args));
            event.completed(args);
        }
    }

    try {
        // --- ADIM 1: GÖNDEREN E-POSTA VE DOMAIN TESPİTİ ---
        const userEmail = getSenderEmail();
        const senderDomain = getDomain(userEmail);

        console.log(`[SmartAlert] 👤 Gönderen Mail: [${userEmail}] | Domain: [${senderDomain}]`);

        if (!userEmail) {
            console.warn("[SmartAlert] Gönderen mail adresi tespit edilemedi, kontrol atlanıyor.");
            safeComplete({ allowEvent: true }, "Gönderen adresi yok");
            return;
        }

        const item = Office.context.mailbox.item;

        // --- ADIM 2: ALICILARI (TO, CC, BCC) PARALEL OLARAK ÇEK ---
        Promise.all([
            getRecipientsFast(item.to, "To"),
            getRecipientsFast(item.cc, "Cc"),
            getRecipientsFast(item.bcc, "Bcc")
        ]).then((results) => {
            const allRecipients = results.flat();
            console.log(`[SmartAlert] 📩 Toplam ${allRecipients.length} alıcı tespit edildi:`, allRecipients);

            if (allRecipients.length === 0) {
                console.log("[SmartAlert] ✅ Alıcı bulunamadı.");
                safeComplete({ allowEvent: true }, "Alıcı yok");
                return;
            }

            // --- ADIM 3: ALICI DOMAINLERINI VE MARKA KÖKLERİNİ ANALİZ ET ---
            // Tanımlı müşteri domainlerinin uzantısız marka isimlerini çıkar ("onduline.com.tr" -> "onduline")
            const clientBrands = CLIENT_DOMAINS.map(d => getBrandFromDomain(d));

            const recipientDomains = new Set();
            const detectedClientBrands = new Set();

            for (let i = 0; i < allRecipients.length; i++) {
                const email = allRecipients[i];
                const domain = getDomain(email);

                if (domain) {
                    recipientDomains.add(domain);

                    const brand = getBrandFromDomain(domain);
                    console.log(`[SmartAlert] 🏢 brand : `, brand);
                    
                    // Alıcının marka adı tanımlı listede varsa kaydet
                    if (clientBrands.includes(brand)) {
                        detectedClientBrands.add(brand);
                        console.log(`[SmartAlert] 🏢 detectedClientBrands : `, brand);
                    }
                }
            }

            console.log(`[SmartAlert] 🔍 Tespit Edilen Tüm Alıcı Domainleri:`, Array.from(recipientDomains));
            console.log(`[SmartAlert] 🏢 Tespit Edilen Müşteri Markaları:`, Array.from(detectedClientBrands));

            // --- KONTROL 1: ÇAPRAZ FİRMA KONTROLÜ (ÇOKLU FİRMA ENGELİ) ---
            if (detectedClientBrands.size > 1) {
                const clientList = Array.from(detectedClientBrands).map(b => b.toUpperCase()).join(", ");
                console.log(`[SmartAlert] 🛑 ENGELLEME (Çoklu Firma Çakışması): Mailde ${detectedClientBrands.size} farklı firma tespit edildi -> [${clientList}]`);

                safeComplete({
                    allowEvent: false,
                    errorMessage: `GÜVENLİK ENGELİ: Aynı e-postada birden fazla müşteri firması ekli olamaz! Tespit edilen firmalar: [${clientList}]. Lütfen her firma için ayrı mail oluşturun.`
                }, "Çoklu Müşteri Firması Engeli");
                return;
            }

            // --- HER ŞEY UYGUNSE GÖNDERİME İZİN VER ---
            console.log("[SmartAlert] ✅ TÜM KONTROLLER BAŞARILI. Gönderime izin veriliyor.");
            safeComplete({ allowEvent: true }, "Tüm kurallar doğrulandı");

        }).catch((err) => {
            console.error("[SmartAlert] ❌ Alıcı okuma aşamasında hata:", err);
            safeComplete({ allowEvent: true }, "Okuma hatası");
        });

    } catch (err) {
        console.error("[SmartAlert] ❌ Ana işlem hatası:", err);
        safeComplete({ allowEvent: true }, "Kritik ana hata");
    }
}

// ============================================================================
// 🛠️ YARDIMCI FONKSİYONLAR (HELPERS)
// ============================================================================

function getSenderEmail() {
    const mailbox = Office.context.mailbox;
    if (mailbox && mailbox.userProfile && mailbox.userProfile.emailAddress) {
        return mailbox.userProfile.emailAddress.toLowerCase().trim();
    }
    if (mailbox && mailbox.initialData && mailbox.initialData.userEmailAddress) {
        return mailbox.initialData.userEmailAddress.toLowerCase().trim();
    }
    return "";
}

function getDomain(email) {
    if (!email || typeof email !== "string") return "";
    const atIndex = email.lastIndexOf("@");
    return atIndex !== -1 ? email.substring(atIndex + 1).toLowerCase().trim() : "";
}

/**
 * Domain adının uzantısını temizleyerek ilk marka parçasını döndürür.
 * Örn: "onduline.com" -> "onduline"
 * Örn: "onduline.com.tr" -> "onduline"
 */
function getBrandFromDomain(domain) {
    if (!domain || typeof domain !== "string") return "";
    return domain.split(".")[0].toLowerCase().trim();
}

function getRecipientsFast(field, fieldName) {
    return new Promise((resolve) => {
        if (!field || typeof field.getAsync !== "function") {
            resolve([]);
            return;
        }
        field.getAsync((result) => {
            if (result && result.status === Office.AsyncResultStatus.Succeeded && Array.isArray(result.value)) {
                const len = result.value.length;
                const emails = [];

                for (let i = 0; i < len; i++) {
                    const item = result.value[i];
                    const addr = (typeof item === "string") ? item : (item.emailAddress || item.address || "");
                    if (addr) {
                        emails.push(addr.toLowerCase().trim());
                    }
                }
                console.log(`[SmartAlert] 📥 [${fieldName}] alanından ${emails.length} adres okundu.`);
                resolve(emails);
            } else {
                resolve([]);
            }
        });
    });
}

// Handler ilişkilendirmesi
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
