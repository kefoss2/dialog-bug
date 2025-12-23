Office.initialize = function () { };

const PROMPT_URL = "https://localhost:3000/Prompt.html";
let clickEvent;
let dialog;
let readWriteToken = "";

function deleteEmail(event){
    Office.onReady(async function (info) {
        console.log("Starting delete function");

        //get graph auth token
        await getGraphToken();
        
        //display dialog
        clickEvent = event;
        openDialog();

        console.log("Function finished");
    });
}

async function getGraphToken(){
    const msalConfig = {
        auth: {
            clientId: 'a652e239-c37b-4d4c-986b-7cefabda1b0a',
            authority: 'https://login.microsoftonline.com/f64e9def-96e2-421a-95bb-7ebc1f584e99'
        },
        system: {
            loggerOptions: {
                loggerCallback: (level, message, containsPii) => {
                    console.log(message)
                },
            }
        },
    };

    const myMSALObj = await msal.createNestablePublicClientApplication(msalConfig);

    const tokenRequest = {
        scopes: ['Mail.ReadWrite'],
        claims: undefined,
    };

    try{
        readWriteToken = (await myMSALObj.acquireTokenSilent(tokenRequest)).accessToken;
    }
    catch(e){
        console.log("Unable to get token silently")
        readWriteToken = (await myMSALObj.acquireTokenPopup(tokenRequest)).accessToken;
    }
}

function openDialog() {
    Office.context.ui.displayDialogAsync(PROMPT_URL, { height: 50, width: 50, displayInIframe: true }, dialogCallback);
}

function dialogCallback(asyncResult) {
    if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
        console.error("displayDialogAsync failed", asyncResult.error);
        if (clickEvent) {
            clickEvent.completed();
        }
        return;
    }

    dialog = asyncResult.value;

    // capture message "cancel" or "move" sent from script in Prompt.html
    dialog.addEventHandler(Office.EventType.DialogMessageReceived, messageHandler);

    // capture event (closing Prompt.html dialog with x button)
    dialog.addEventHandler(Office.EventType.DialogEventReceived, eventHandler);

}

function eventHandler(arg) {
    console.log(`Received event from prompt dialog.`);
    console.table(arg);
    clickEvent.completed();
}

async function messageHandler(arg) {
    console.log(`Received message "${arg.message}" from prompt dialog.`);
    dialog.close();
    switch (arg.message) {
        case "cancel":
            clickEvent.completed();
            break;
        case "move":
            const id = Office.context.mailbox.item?.itemId.replaceAll("/", "-");
            const emailAddress = Office.context.mailbox.userProfile.emailAddress;
            const path = `https://graph.microsoft.com/v1.0/users/${emailAddress}/messages/${id}/move`;
            let response = await fetch(path, {
                method: "POST",
                headers: {
                    Authorization: readWriteToken,
                    "Content-Type": "application/json"
                },
                body: JSON.stringify({destinationId:"deleteditems"})
            });
            clickEvent.completed();
            break;
        default:
            break;
    }
}