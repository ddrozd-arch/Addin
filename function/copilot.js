function runSmartAlerts(event) {
    // Tu Twoja logika walidacji
    const shouldBlock = false; // przykład

    if (shouldBlock) {
        event.completed({
            allowEvent: false,
            errorMessage: "Wiadomość nie spełnia wymagań."
        });
    } else {
        event.completed({ allowEvent: true });
    }
}

function runSmartAlerts(event) {
    // Tu Twoja logika walidacji
    const shouldBlock = false; // przykład

    if (shouldBlock) {
        event.completed({
            allowEvent: false,
            errorMessage: "Wiadomość nie spełnia wymagań."
        });
    } else {
        event.completed({ allowEvent: true });
    }
}
function onMessageSend(event) {
    const client = Office.context.mailbox.diagnostics.hostName;
    const version = Office.context.mailbox.diagnostics.hostVersion;
    const platform = Office.context.platform;

    console.log("Client:", client);
    console.log("Version:", version);
    console.log("Platform:", platform);

    // --- Classic Outlook (Win32) ---
    if (client === "Outlook" && platform === Office.PlatformType.PC) {
        console.log("Classic Outlook detected → bypass Smart Alerts");
        event.completed({ allowEvent: true });
        return;
    }

    // --- Outlook Mobile (iOS/Android) ---
    if (platform === Office.PlatformType.iOS || platform === Office.PlatformType.Android) {
        console.log("Outlook Mobile detected → bypass Smart Alerts");
        event.completed({ allowEvent: true });
        return;
    }

    // --- OWA ---
    if (platform === Office.PlatformType.OfficeOnline) {
        console.log("OWA detected → Smart Alerts active");
        runSmartAlerts(event);
        return;
    }

    // --- New Outlook (Windows/Mac) ---
    if (platform === Office.PlatformType.Mac || platform === Office.PlatformType.Office) {
        console.log("New Outlook detected → Smart Alerts active");
        runSmartAlerts(event);
        return;
    }

    // --- Fallback ---
    console.log("Unknown client → allow");
    event.completed({ allowEvent: true });
}
