# Dialog API Bug

## Summary

Event handlers for dialogs displayed using `Office.context.ui.displayDialogAsync` do not capture `DialogMessageReceived` events after an email has been moved using a Graph move request. After moving an email, the add-in can open another dialog, but events are never captured in the dialog's callback function.

The `main` branch of this repo reproduces this error in the scenario closest to our actual code, where we use promises to get the dialog result in our main function and then move the email and complete the click event there. The `simplified` branch of this repo reproduces the issue using `displayDialogAsync` with a regular callback.

## Instructions to Reproduce

### Create an Entra app

This repo uses Microsoft Graph, so we need an Entra app to get auth tokens.

1. In Entra, go to "App registrations" and click "New registration". 
2. Add a name for the app. Add a redirect URI of type "Single-page application (SPA)": `brk-multihub://localhost:3000`. Click "Register".
3. On the "Overview" page for the new app, write down the "Application (client) ID" and "Directory (tenant) ID" numbers to be added to the local add-in later.
4. On the "Authentication" tab, click "Add Redirect URI" and click "Single-page application". Enter the redirect URI `http://localhost:3000/taskpane.html` and click "Configure".
5. On the "API permissions" tab, click "Add a permission", click "Microsoft Graph", and click "Delegated Permissions". Add the permission `Mail.ReadWrite`.
6. On the same page, click "Grant admin consent for MSFT".

### Install the add-in

1. Log into an email account from the tenant with the app registration from the last section in OWA.
2. Go to `aka.ms/olksideload`, and in the menu that appears, go to "My add-ins" and click "Add a custom add-in".
3. Install the manifest file `manifest-localhost.xml`. Empty the cache and do a hard reload.
4. Open an email and open the apps menu. An app with the name "Dialog API Bug Demo" should appear.

### Run the add-in locally and reproduce the bug

1. Clone this repo locally and navigate to it in a Terminal or bash shell.
2. In the file `taskpane.js`, replace `[APPLICATION ID]` and `[TENANT ID]` with the application and tenant IDs from the app registration.
3. Run the command `npm install`.
4. Run the command `npm start`.
    - From the [Microsoft sample add-in](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/hello-world/outlook-hello-world) used as the basis for this one: If you've never developed an Office add-in on this computer before or it has been more than 30 days since you last did, you'll be prompted to delete an old security certificate and/or install a new one. Agree to both prompts.
5. Open an email and open the apps menu. Click the "Dialog API Bug Demo" button. A dialog with two buttons (one to delete an email and one to cancel) should appear.
6. Click the button to delete the email and wait for the email to be deleted.
7. Open a different email, open the apps menu, and click the "Dialog API Bug Demo" button.
8. On the new dialog, click either button. The dialog does not respond and the email is not deleted.

## References

This repo was created using these two repos:

- https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/hello-world/outlook-hello-world
- https://github.com/andrewlamansky/Dialog-API-Bug/tree/main
