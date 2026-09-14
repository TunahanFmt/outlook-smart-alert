/* global Office */

/**
 * ============================================================================
 * OUTLOOK SMART ALERT - DOKÜMAN VE FİRMA KONTROL EKLENTİSİ
 * ============================================================================
 * Kurallar:
 * 1. Özel tanımlı gönderen mailleri (abdiibrahimfilo@..., ccifilo@... vb.) sadece kendi firmasına mail gönderebilir.
 * 2. Şahsi veya Genel hesaplardan atılan mailde (To, CC, BCC) aynı anda 2 FARKLI müşteri firması bulunamaz.
 * 3. Domain uzantıları (.com, .com.tr vb.) bağımsız marka köküne göre kontrol yapılır.
 * ============================================================================
 */

console.log("[SmartAlert] Kurumsal Mail Güvenlik ve Eşleşme Modülü Yüklendi.");

Office.onReady();

// ============================================================================
// ⚙️ CONFIGURATION / YAPILANDIRMA ALANI
// ============================================================================

// 1. Kurum İçi Domain
const INTERNAL_DOMAIN = "fmtturkey.com";

// 2. Sabit Gönderen E-Posta -> Müşteri Markası Eşleşme Haritası
const SENDER_BRAND_MAP = {
    "abdiibrahimfilo@fmtturkey.com": "abdiibrahim",
    "ccifilo@fmtturkey.com": "cci",
    "akcansafilo@fmtturkey.com": "akcansa",
    "ondulinefilo@fmtturkey.com": "onduline",
    "allianzfilo@fmtturkey.com": "allianz",
    "rochefilo@fmtturkey.com": "roche"
};

// 3. Tanımlı Müşteri / Partner Firma Domain Listesi
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
        const item = Office.context.mailbox.item;

        // --- GÖNDEREN VE ALICILARI PARALEL OLARAK ÇEK ---
        Promise.all([
            getSenderEmailAsync(item),
            getRecipientsFast(item.to, "To"),
            getRecipientsFast(item.cc, "Cc"),
            getRecipientsFast(item.bcc, "Bcc")
        ]).then(([userEmail, toRecipients, ccRecipients, bccRecipients]) => {
            const senderDomain = getDomain(userEmail);

            console.log(`[SmartAlert] 👤 Gönderen Mail: [${userEmail}] | Domain: [${senderDomain}]`);

            if (!userEmail) {
                console.warn("[SmartAlert] Gönderen mail adresi tespit edilemedi, kontrol atlanıyor.");
                safeComplete({ allowEvent: true }, "Gönderen adresi yok");
                return;
            }

            const allRecipients = [...toRecipients, ...ccRecipients, ...bccRecipients];
            console.log(`[SmartAlert] 📩 Toplam ${allRecipients.length} alıcı tespit edildi:`, allRecipients);

            if (allRecipients.length === 0) {
                console.log("[SmartAlert] ✅ Alıcı bulunamadı.");
                safeComplete({ allowEvent: true }, "Alıcı yok");
                return;
            }

            // --- ADIM 1: ALICI DOMAINLERINI VE MARKA KÖKLERİNİ ANALİZ ET ---
			// Tanımlı müşteri domainlerinin uzantısız marka isimlerini çıkar ("onduline.com.tr" -> "onduline")
            const clientBrands = CLIENT_DOMAINS.map(d => getBrandFromDomain(d));
            const recipientDomains = new Set();
            const detectedRecipientBrands = new Set();

            for (let i = 0; i < allRecipients.length; i++) {
                const email = allRecipients[i];
                const domain = getDomain(email);

                if (domain) {
                    recipientDomains.add(domain);
                    const brand = getBrandFromDomain(domain);
					console.log(`[SmartAlert] 🏢 brand : `, brand);
                    
                    // Alıcının marka adı tanımlı listede varsa kaydet
                    if (clientBrands.includes(brand)) {
                        detectedRecipientBrands.add(brand);
						console.log(`[SmartAlert] 🏢 detectedRecipientBrands : `, brand);
                    
                    }
                }
            }

            console.log(`[SmartAlert] 🔍 Tespit Edilen Tüm Alıcı Domainleri:`, Array.from(recipientDomains));
            console.log(`[SmartAlert] 🏢 Tespit Edilen Müşteri Markaları (Alıcılarda):`, Array.from(detectedRecipientBrands));

            // --- ADIM 2: GÖNDEREN E-POSTA ÖZEL EŞLEŞME KONTROLÜ ---
            const senderBrand = SENDER_BRAND_MAP[userEmail];

            if (senderBrand) {
                console.log(`[SmartAlert] 🎯 Gönderen adresi sabit listede tanımlı: [${userEmail}] -> Marka: [${senderBrand}]`);
                
                // Gönderenin tanımlı markası haricinde farklı bir müşteri markası alıcılarda var mı?
                const foreignBrands = Array.from(detectedRecipientBrands).filter(b => b !== senderBrand);

                if (foreignBrands.length > 0) {
                    const targetCompanies = foreignBrands.map(b => b.toUpperCase()).join(", ");
                    console.log(`[SmartAlert] 🛑 ENGELLEME: [${userEmail}] adresinden [${targetCompanies}] firmasına mail gönderilemez.`);

                    safeComplete({
                        allowEvent: false,
                        errorMessage: `GÜVENLİK ENGELİ: ${userEmail} adresinden ${targetCompanies} firmasına mail atamazsınız!`
                    }, "Gönderen-Alıcı Firma Uyuşmazlık Engeli");
                    return;
                }
            }

            // --- ADIM 3: ÇAPRAZ FİRMA KONTROLÜ (ÇOKLU FİRMA ENGELİ) ---
            if (detectedRecipientBrands.size > 1) {
                const clientList = Array.from(detectedRecipientBrands).map(b => b.toUpperCase()).join(", ");
                console.log(`[SmartAlert] 🛑 ENGELLEME (Çoklu Firma Çakışması): Mailde ${detectedRecipientBrands.size} farklı firma tespit edildi -> [${clientList}]`);

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
            console.error("[SmartAlert] ❌ Alıcı/Gönderen okuma aşamasında hata:", err);
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

/**
 * Paylaşılan kutu (Shared Mailbox) ve "Kimden" alanı değişikliklerini 
 * doğru şekilde yakalayan asenkron gönderen adresi okuyucu.
 */
function getSenderEmailAsync(item) {
    return new Promise((resolve) => {
        if (item && item.from && typeof item.from.getAsync === "function") {
            item.from.getAsync((result) => {
                if (result && result.status === Office.AsyncResultStatus.Succeeded && result.value) {
                    const addr = result.value.emailAddress || result.value.address || "";
                    if (addr) {
                        resolve(addr.toLowerCase().trim());
                        return;
                    }
                }
                resolve(getFallbackUserEmail());
            });
        } else {
            resolve(getFallbackUserEmail());
        }
    });
}

function getFallbackUserEmail() {
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
