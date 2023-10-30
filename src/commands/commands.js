/* global global, Office, self, window */
Office.onReady(() => {
    // Office.js est prêt à être appelé si nécessaire
});

function getGlobal() {
    return typeof self !== "undefined"
        ? self
        : typeof window !== "undefined"
            ? window
            : typeof global !== "undefined"
                ? global
                : undefined;
}

const g = getGlobal();

function successNotification(msg) {
    var id = "0";
    var details = {
        type: "informationalMessage",
        iconUrl: "Icon.16x16.png",
        message: msg,
        persistent: false
    };
    Office.context.mailbox.item.notificationMessages.addAsync(id, details, function (value) {
    });
}

function failedNotification(msg) {
    var id = "0";
    var details = {
        type: "errorMessage",
        iconUrl: "Icon.16x16.png",
        message: msg,
        persistent: false
    };
    Office.context.mailbox.item.notificationMessages.addAsync(id, details, function (value) {
    });
}

function getItemRestId() {
    if (Office.context.mailbox.diagnostics.hostName === "OutlookIOS") {
        return Office.context.mailbox.item.itemId;
    } else {
        return Office.context.mailbox.convertToRestId(
            Office.context.mailbox.item.itemId,
            Office.MailboxEnums.RestVersion.v2_0
        );
    }
}

/* Transfert simple d'e-mail */
function simpleForwardEmail() {
    Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, function (result) {
        var accessToken = result.value;
        simpleForwardFunc(accessToken);
    });
}

function simpleForwardFunc(accessToken) {
    var itemId = getItemRestId();
    var forwardUrl = Office.context.mailbox.restUrl + "/v2.0/me/messages/" + itemId + "/forward";

    const forwardMeta = JSON.stringify({
        Comment: "FYI",
        ToRecipients: [
            {
                EmailAddress: {
                    Name: "ccmbercy",
                    Address: "gco-ccm@outlook.fr"
                }
            }
        ]
    });

    $.ajax({
        url: forwardUrl,
        type: "POST",
        dataType: "json",
        contentType: "application/json",
        data: forwardMeta,
        headers: { Authorization: "Bearer " + accessToken }
    }).always(function (response) {
        successNotification("Transfert d'e-mail réussi !");
        // Après le transfert réussi, déplacez l'e-mail vers le dossier "Junk" (Courrier indésirable)
        moveEmailToJunk(accessToken, itemId);
    });
}

/* Déplacer l'e-mail vers le dossier "Junk" (Courrier indésirable) */
function moveEmailToJunk(accessToken, emailId) {
    // Construisez l'URL REST pour effectuer l'opération de déplacement vers le dossier "Junk" (Courrier indésirable)
    var moveEmailUrl = Office.context.mailbox.restUrl + "/v2.0/me/messages/" + emailId + "/move";

    // Créez les informations pour le déplacement (dans le dossier "Junk" / Courrier indésirable)
    var moveInfo = JSON.stringify({
        DestinationId: "Junk"
    });

    $.ajax({
        url: moveEmailUrl,
        type: "POST",
        dataType: "json",
        contentType: "application/json",
        data: moveInfo,
        headers: { Authorization: "Bearer " + accessToken }
    }).done(function (response) {
        successNotification("E-mail déplacé vers le dossier Courrier indésirable.");
    }).fail(function (error) {
        failedNotification("Échec du déplacement de l'e-mail : " + error.responseText);
    });
}
