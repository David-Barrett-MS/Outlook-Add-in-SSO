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
const userName = document.getElementById("userName");
const userEmail = document.getElementById("userEmail");

Office.onReady((info) => {
  switch (info.host) {
    case Office.HostType.Outlook:
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
      console.log("Initializing account manager...");
      accountManager.initialize();
      applyOfficeTheme();
      break;
  }
});

function applyOfficeTheme() {
  // Identify the current Office theme in use.
  const currentOfficeTheme = Office.context.officeTheme.themeId;
  console.log("Current Office theme: " + currentOfficeTheme);

  if (!Office.context.officeTheme.isDarkTheme) {
      console.log("No changes required.");
  }

  console.log("Applying Office theme...");
  // Get the colors of the current Office theme.
  const bodyBackgroundColor = Office.context.officeTheme.bodyBackgroundColor;
  const bodyForegroundColor = Office.context.officeTheme.bodyForegroundColor;
  const controlBackgroundColor = Office.context.officeTheme.controlBackgroundColor;
  const controlForegroundColor = Office.context.officeTheme.controlForegroundColor;

  document.body.style.backgroundColor = bodyBackgroundColor;
  document.body.style.color = bodyForegroundColor;

  if (Office.context.officeTheme.isDarkTheme) {
    console.log("Dark theme detected.");
  }
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