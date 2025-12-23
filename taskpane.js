Office.initialize = function () { };

const PROMPT_URL = "https://localhost:3000/Prompt.html";
let clickEvent;
let dialog;
let readWriteToken = "";

function deleteEmail(event){
    Office.onReady(async function (info) {
        console.log("Starting delete function");
        
        //display dialog
        clickEvent = event;
        let dialogResult = await openDialog();
        let dialogCallbackResult = await dialogCallback(dialogResult);

        //delete email if user selected "delete"
        if(dialogCallbackResult.event.message === 'move'){
            console.log("Deleting email");
            //get graph auth token
            await getGraphToken();
            
            await moveEmailToDelete();
        }
        
        console.log("Function finished");
        clickEvent.completed();

    });
}

async function openDialog() {
    return new Promise((resolve, _) => {
        Office.context.ui.displayDialogAsync(PROMPT_URL, { height: 50, width: 50, displayInIframe: true }, resolve);
    });
}

function dialogCallback(asyncResult) {
    if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
        console.error("displayDialogAsync failed", asyncResult.error);
        if (clickEvent) {
            clickEvent.completed();
        }
        return;
    }
    return new Promise((resolve, reject) => {
        try{
            dialog = asyncResult.value;

            // capture message "cancel" or "move" sent from script in Prompt.html
            dialog.addEventHandler(Office.EventType.DialogMessageReceived, function(arg) {
                console.log(`Received message "${arg.message}" from prompt dialog.`);
                dialog.close();
                resolve({ event: arg});
            });

            // capture event (closing Prompt.html dialog with x button)
            dialog.addEventHandler(Office.EventType.DialogEventReceived, function(arg){
                console.log(`Received event from prompt dialog.`);
                console.table(arg);
                resolve({ event: arg});
            });
        }
        catch (e) {
            console.log(e);
            reject(e);
        }
    });
}

async function getGraphToken(){
    const msalConfig = {
        auth: {
            clientId: '[APPLICATION ID]',
            authority: 'https://login.microsoftonline.com/[TENANT ID]'
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

async function moveEmailToDelete(arg) {
    const id = Office.context.mailbox.item?.itemId.replaceAll("/", "-");
    const emailAddress = Office.context.mailbox.userProfile.emailAddress;
    const path = `https://graph.microsoft.com/v1.0/users/${emailAddress}/messages/${id}/move`;
    let response = await fetch(path, {
        method: "POST",
        headers: {
            Authorization: readWriteToken,
            "Content-Type": "application/json"
        },
        body: JSON.stringify({destinationId: "deleteditems"})
    });
}