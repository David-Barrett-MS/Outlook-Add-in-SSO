/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import { AccountManager } from "./authConfig.js";
import { makeGraphRequest } from "./msgraph-helper.js";

const accountManager = new AccountManager();
const sideloadMsg = document.getElementById("sideload-msg");
const appBody = document.getElementById("app-body");
const getUserDataButton = document.getElementById("getUserData");
const getUserFilesButton = document.getElementById("getUserFiles");
const tenantIdElement = document.getElementById("entraTenantId");
const appIdElement = document.getElementById("entraAppId");

/**
 * The add-in settings object.
 * @type {Office.RoamingSettings}
 */
let addinSettings;
let tenantId;
let applicationId;

Office.onReady((info) => {
  if (info.host == Office.HostType.Outlook) {
    if (sideloadMsg) {
      sideloadMsg.style.display = "none";
    }
    if (appBody) {
      appBody.style.display = "flex";
    }
    if (getUserDataButton) {
      getUserDataButton.onclick = getUserData;
    }
    if (getUserFilesButton) {
      getUserFilesButton.onclick = getUserFiles;
    }


    // Initialize the roaming settings object and retrieve client information.
    addinSettings = Office.context.roamingSettings;
    tenantId = addinSettings.get("tenantId");
    applicationId = addinSettings.get("applicationId");

    // Write the application information to the TaskPane and console.
    console.log("Application ID: " + applicationId);
    console.log("Tenant ID: " + tenantId);
    const appIdElement = document.getElementById("entraAppId");
    if (appIdElement) {
      appIdElement.value = applicationId;
      appIdElement.onchange = updateApplicationId;
    }
    if (tenantIdElement) {
      tenantIdElement.value = tenantId;
      tenantIdElement.onchange = updateTenantId;
    }

    initialiseAccountManager();

    applyOfficeTheme();
  }
});

function initialiseAccountManager() {
  addinSettings = Office.context.roamingSettings;
  tenantId = addinSettings.get("tenantId");
  applicationId = addinSettings.get("applicationId");

  console.log("Initializing account manager...");
  console.log("Application ID: " + applicationId);
  console.log("Tenant ID: " + tenantId);
  accountManager.initialize(applicationId, tenantId);
}

function applyOfficeTheme() {
  // Identify the current Office theme in use.
  const currentOfficeTheme = Office.context.officeTheme.themeId;

  if (currentOfficeTheme === undefined) {
    console.log("No Office theme detected.");
    return;
  }
  console.log("Current Office theme: " + currentOfficeTheme);

  console.log("Applying Office theme...");
  document.body.style.backgroundColor = Office.context.officeTheme.bodyBackgroundColor;
  document.body.style.color = Office.context.officeTheme.bodyForegroundColor;

  if (Office.context.officeTheme.isDarkTheme) {
    console.log("Dark theme detected.");
  }
}

async function updateTenantId() {
  const newTenantId = tenantIdElement.value;
  console.log("New tenant ID: " + newTenantId);
  addinSettings.set("tenantId", newTenantId);
  await addinSettings.saveAsync();
  console.log("Tenant ID saved.");
  initialiseAccountManager();
}

async function updateApplicationId() {
  const newApplicationId = appIdElement.value;
  console.log("New application ID: " + newApplicationId);
  addinSettings.set("applicationId", newApplicationId);
  await addinSettings.saveAsync();
  console.log("Application ID saved.");
  initialiseAccountManager();
}

/**
 * Gets the user data such as name and email and displays it
 * in the task pane.
 */
async function getUserData() {
  try {
    const userDataElement = document.getElementById("userData");
    const userAccount = await accountManager.ssoGetUserIdentity(["user.read"]);
    const idTokenClaims = userAccount.idTokenClaims;

    console.log(userAccount);

    if (userDataElement) {
      userDataElement.style.visibility = "visible";
    }
    if (userName) {
      userName.innerText = idTokenClaims.name ?? "";
    }
    if (userEmail) {
      userEmail.innerText = idTokenClaims.preferred_username ?? "";
    }
  } catch (error) {
    console.error(error);
  }
}

/**
 * Gets the first 10 item names (files or folders) from the user's OneDrive.
 * Displays the item names in the TaskPane.
 */
async function getUserFiles() {
  try {
    const names = await getFileNames();
    console.log(names.length + " items found.");

    const userFilesElement = document.getElementById("userFiles");
    if (userFilesElement) {
      userFilesElement.style.visibility = "visible";
      const userFilesListElement = document.getElementById("fileList");
      userFilesListElement.innerHTML = ""; // Clear previous list
      console.log("Writing file names to the taskpane...");
      names.forEach((name) => {
        const listItem = document.createElement("li");
        listItem.innerText = name;
        userFilesListElement.appendChild(listItem);
        console.log(name);
      });
    }

  } catch (error) {
    console.error(error);
  }
}

async function getFileNames(count = 10) {
  const accessToken = await accountManager.ssoGetToken(["Files.Read"]);
  const response = await makeGraphRequest(
    accessToken,
    "/me/drive/root/children",
    `?$select=name&$top=${count}`
  );

  const names = response.value.map(item => item.name);
  return names;
}