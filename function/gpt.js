function getClientType() {
    const diagnostics = Office.context.diagnostics;
    const platform = diagnostics.platform;
    const userAgent = navigator.userAgent || "";

    // 📱 Mobile
    if (
        platform === Office.PlatformType.iOS ||
        platform === Office.PlatformType.Android
    ) {
        return "mobile";
    }

    // 🌐 OWA
    if (platform === Office.PlatformType.OfficeOnline) {
        return "owa";
    }

    // 💻 Desktop
    if (
        platform === Office.PlatformType.PC ||
        platform === Office.PlatformType.Mac
    ) {
        const isWebView =
            userAgent.includes("WebView") ||
            userAgent.includes("Edge");

        return isWebView ? "new_outlook" : "classic_outlook";
    }

    return "unknown";
}

async function onMessageSendHandler(event) {
    try {
        const client = getClientType();
        console.log("Client:", client);

        // ✅ 1. CLASSIC → total bypass
        if (client === "classic_outlook") {
            event.completed({ allowEvent: true });
            return;
        }

        // ✅ 2. MOBILE → też bypass (zalecane)
        if (client === "mobile") {
            event.completed({ allowEvent: true });
            return;
        }

        // 🔍 3. NEW OUTLOOK + OWA → Smart Alert
        const item = Office.context.mailbox.item;

        const hasAttachments =
            item.attachments && item.attachments.length > 0;

        if (!hasAttachments) {

            // 👉 Smart Alert (promptUser)
            event.completed({
                allowEvent: false,
                errorMessage: getSmartAlertMessage(),
                errorMessageMarkdown: getSmartAlertMarkdown()
            });

            return;
        }

        // ✅ OK
        event.completed({ allowEvent: true });

    } catch (err) {
        console.error(err);

        // fallback: nie blokuj usera
        event.completed({ allowEvent: true });
    }
}

function getSmartAlertMessage() {
    return "Attachment missing";
}

function getSmartAlertMarkdown() {
    return `
**Brak załącznika**

Wygląda na to, że wiadomość nie zawiera załącznika.

**Możliwe opcje:**
- Dodaj załącznik
- Wyślij mimo to

---
*This message may require an attachment.*
`;
}

//---------//

Office.onReady(() => {
    // wymagane w event-based addin
});

async function onMessageSendHandler(event) {
    try {
        const diagnostics = Office.context.diagnostics;
        const platform = diagnostics.platform; // PC, OfficeOnline, Mac, iOS, Android
        const host = diagnostics.host; // Outlook

        const userAgent = navigator.userAgent || "";

        let clientType = "unknown";

        // 📱 Mobile
        if (platform === Office.PlatformType.iOS || platform === Office.PlatformType.Android) {
            clientType = "mobile";
        }
        // 🌐 OWA (Outlook Web Access)
        else if (platform === Office.PlatformType.OfficeOnline) {
            clientType = "owa";
        }
        // 💻 Desktop (Windows / Mac)
        else if (platform === Office.PlatformType.PC || platform === Office.PlatformType.Mac) {

            // 👉 Próba wykrycia "New Outlook"
            if (userAgent.includes("WebView") || userAgent.includes("Edge")) {
                clientType = "new_outlook";
            } else {
                clientType = "classic_outlook";
            }
        }

        console.log("Client type:", clientType);

        // 👉 Twoja logika zależna od klienta
        switch (clientType) {
            case "mobile":
                // np. ograniczona walidacja
                break;

            case "owa":
                // pełna walidacja web
                break;

            case "new_outlook":
                // zachowuje się jak OWA, ale czasem różnice w UI
                break;

            case "classic_outlook":
                // desktop legacy
                break;
        }

        // ✅ pozwól wysłać
        event.completed({ allowEvent: true });

    } catch (error) {
        console.error(error);

        // ❌ blokuj wysyłkę w razie błędu
        event.completed({
            allowEvent: false,
            errorMessage: "Unexpected error in add-in"
        });
    }
}