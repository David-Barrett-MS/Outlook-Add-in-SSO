# Outlook-Add-in-SSO

Shows how to make Graph calls from an Outlook add-in using SSO.  Client-side code only (no server besides static web server required), and no frameworks such as Node.  Code adapted from other Microsoft add-in samples.

The Graph calls require [an application registration](https://learn.microsoft.com/en-gb/office/dev/add-ins/develop/enable-nested-app-authentication-in-your-add-in#register-your-single-page-application).  Application information (tenant and app id) is entered into the TaskPane (authentication calls will not work prior to that).  This allows the add-in to be directly tested with the files served by Github.

## Running from this repository

This repository is Github Pages enabled, so the add-in can be served directly.  To use the add-in directly, you'll need to configure the application registration with a redirect URL of brk-multihub://david-barrett-ms.github.io/Outlook-Add-in-SSO (as that is where the add-in is hosted).  You can install the add-in using [the manifest that targets Github pages](https://github.com/David-Barrett-MS/Outlook-Add-in-SSO/blob/main/www/Outlook%20SSO%20Add-in%20Github.xml).

## Serving from your own server

To use, host the files (under www) on a suitable web server and update the references in the [app manifest](https://github.com/David-Barrett-MS/Outlook-Add-in-SSO/blob/main/www/Outlook%20SSO%20Add-in.xml) to point to that server.  You'll also need to [register an application in EntraID](https://learn.microsoft.com/en-gb/office/dev/add-ins/develop/enable-nested-app-authentication-in-your-add-in#register-your-single-page-application) and update [applicationId](https://github.com/David-Barrett-MS/Outlook-Add-in-SSO/blob/main/www/authConfig.js#L9) and [authority](https://github.com/David-Barrett-MS/Outlook-Add-in-SSO/blob/main/www/authConfig.js#L14) to your own (authority includes tenant Id, which will need to be changed).  Once done, install the add-in using your updated manifest.

## Testing the add-in

The add-in exposes a TaskPane with two buttons on (one to retrieve user information, the other to list the user's OneDrive files).  The taskpane is available on mail items in both read and compose mode.  The first time the TaskPane is opened for each mailbox, you'll need to configure the tenant and application Id at the bottom of the TaskPane.  These are saved into add-in settings so in future you'll see these correctly set.  Once the application information is set, test the buttons.

The add-in logs a lot of what it does to the console, so this can be monitored using Dev Tools.