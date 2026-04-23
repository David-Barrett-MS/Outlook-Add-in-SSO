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
const getSharedMailboxMessagesButton = document.getElementById("getSharedMailboxMessages");
const saveDraftAndGetViaGraphButton = document.getElementById("saveDraftAndGetViaGraph");
const sharedMailboxAddressElement = document.getElementById("sharedMailboxAddress");
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
    if (getSharedMailboxMessagesButton) {
      getSharedMailboxMessagesButton.onclick = getSharedMailboxMessages;
    }
    if (saveDraftAndGetViaGraphButton) {
      saveDraftAndGetViaGraphButton.onclick = saveDraftAndGetViaGraph;
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

/**
 * Gets the top 5 messages from a shared mailbox that the signed-in user can access.
 * Logs message subjects on success, or full HTTP response details on failure.
 */
async function getSharedMailboxMessages() {
  const sharedMailboxAddress = sharedMailboxAddressElement?.value?.trim();
  if (!sharedMailboxAddress) {
    console.error("Enter a shared mailbox address before running this test.");
    return;
  }

  try {
    const accessToken = await accountManager.ssoGetToken(["Mail.ReadWrite.Shared"]);
    const authorizationHeader = accessToken.startsWith("Bearer ") ? accessToken : `Bearer ${accessToken}`;
    const query = "?$select=subject&$top=5&$orderby=receivedDateTime desc";
    const requestUrl = `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(sharedMailboxAddress)}/messages${query}`;

    const response = await fetch(requestUrl, {
      headers: {
        Authorization: authorizationHeader,
      },
    });

    if (!response.ok) {
      await logGraphErrorResponse(response, sharedMailboxAddress);
      return;
    }

    const payload = await response.json();
    const subjects = (payload.value || []).map((item) => item.subject ?? "(no subject)");

    console.log(`Top ${subjects.length} messages from shared mailbox ${sharedMailboxAddress}:`);
    subjects.forEach((subject, index) => {
      console.log(`${index + 1}. ${subject}`);
    });

    // Display messages in the TaskPane
    const sharedMailboxMessagesElement = document.getElementById("sharedMailboxMessages");
    const messageListElement = document.getElementById("messageList");

    if (sharedMailboxMessagesElement && messageListElement) {
      messageListElement.innerHTML = ""; // Clear previous list

      if (subjects.length === 0) {
        const noMessageItem = document.createElement("li");
        noMessageItem.innerText = "No messages found in this shared mailbox.";
        messageListElement.appendChild(noMessageItem);
      } else {
        subjects.forEach((subject) => {
          const listItem = document.createElement("li");
          listItem.innerText = subject;
          messageListElement.appendChild(listItem);
        });
      }

      sharedMailboxMessagesElement.style.visibility = "visible";
    }
  } catch (error) {
    console.error("Error retrieving shared mailbox messages.", error);
  }
}

async function logGraphErrorResponse(response, sharedMailboxAddress) {
  const headers = {};
  response.headers.forEach((value, key) => {
    headers[key] = value;
  });

  const contentType = response.headers.get("content-type") || "";
  const body = contentType.includes("application/json")
    ? await response.json()
    : await response.text();

  console.error(`Shared mailbox request failed for ${sharedMailboxAddress}.`, {
    status: response.status,
    statusText: response.statusText,
    headers,
    body,
  });
}

async function saveDraftAndGetViaGraph() {
  try {
    const mailboxItem = Office.context.mailbox?.item;
    if (!mailboxItem || typeof mailboxItem.saveAsync !== "function") {
      throw new Error("This test requires an Outlook compose item that supports saveAsync.");
    }

    const operationStartTime = Date.now();
    console.log("Saving current item draft...");
    const savedItemId = await saveCurrentItemAsync(mailboxItem);
    console.log("Draft saved.", { savedItemId });

    const graphMessageId = convertItemIdForGraph(savedItemId);
    const mailboxContext = await getMailboxContextForGraph(mailboxItem);
    console.log("Retrieving saved draft from Graph...", {
      graphMessageId,
      isSharedMailbox: mailboxContext.isShared,
      mailboxAddress: mailboxContext.mailboxAddress,
    });

    const scopes = mailboxContext.isShared ? ["Mail.ReadWrite.Shared"] : ["Mail.ReadWrite"];
    const accessToken = await accountManager.ssoGetToken(scopes);
    const authorizationHeader = accessToken.startsWith("Bearer ") ? accessToken : `Bearer ${accessToken}`;
    const result = await getMessageViaGraphWithRetry(graphMessageId, authorizationHeader, mailboxContext);
    const totalTimeMs = Date.now() - operationStartTime;

    console.log("Saved draft retrieved from Graph.", {
      retriesNeeded: result.retriesNeeded,
      totalTimeMs,
      message: result.message,
    });

    // Display results in the TaskPane
    const draftGraphResultsElement = document.getElementById("draftGraphResults");
    const retriesNeededElement = document.getElementById("retriesNeeded");
    const totalTimeToGraphElement = document.getElementById("totalTimeToGraph");
    const resultItemIdElement = document.getElementById("resultItemId");
    const resultItemSubjectElement = document.getElementById("resultItemSubject");

    if (draftGraphResultsElement && retriesNeededElement && totalTimeToGraphElement && resultItemIdElement && resultItemSubjectElement) {
      retriesNeededElement.innerText = result.retriesNeeded;
      totalTimeToGraphElement.innerText = `${totalTimeMs} ms`;
      resultItemIdElement.innerText = result.message.id || "(no id)";
      resultItemSubjectElement.innerText = result.message.subject || "(no subject)";
      draftGraphResultsElement.style.visibility = "visible";
    }
  } catch (error) {
    console.error("Error saving draft and retrieving it via Graph.", error);
  }
}

function saveCurrentItemAsync(mailboxItem) {
  return new Promise((resolve, reject) => {
    mailboxItem.saveAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value);
        return;
      }

      reject(result.error || new Error("saveAsync failed."));
    });
  });
}

function convertItemIdForGraph(itemId) {
  const mailbox = Office.context.mailbox;
  if (mailbox && typeof mailbox.convertToRestId === "function") {
    return mailbox.convertToRestId(itemId, Office.MailboxEnums.RestVersion.v2_0);
  }

  return itemId;
}

async function getMailboxContextForGraph(mailboxItem) {
  if (!mailboxItem || typeof mailboxItem.getSharedPropertiesAsync !== "function") {
    return {
      isShared: false,
      mailboxAddress: null,
    };
  }

  try {
    const sharedProperties = await getSharedPropertiesAsync(mailboxItem);
    const mailboxAddress = sharedProperties?.targetMailbox?.trim();

    if (mailboxAddress) {
      return {
        isShared: true,
        mailboxAddress,
      };
    }
  } catch (error) {
    console.warn("Unable to read shared mailbox properties. Falling back to /me endpoint.", error);
  }

  return {
    isShared: false,
    mailboxAddress: null,
  };
}

function getSharedPropertiesAsync(mailboxItem) {
  return new Promise((resolve, reject) => {
    mailboxItem.getSharedPropertiesAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value);
        return;
      }

      reject(result.error || new Error("getSharedPropertiesAsync failed."));
    });
  });
}

function buildGraphMessageRequestUrl(messageId, mailboxContext) {
  const encodedMessageId = encodeURIComponent(messageId);
  if (mailboxContext?.isShared && mailboxContext.mailboxAddress) {
    const encodedMailboxAddress = encodeURIComponent(mailboxContext.mailboxAddress);
    return `https://graph.microsoft.com/v1.0/users/${encodedMailboxAddress}/messages/${encodedMessageId}`;
  }

  return `https://graph.microsoft.com/v1.0/me/messages/${encodedMessageId}`;
}

async function getMessageViaGraphWithRetry(messageId, authorizationHeader, mailboxContext) {
  const requestUrl = buildGraphMessageRequestUrl(messageId, mailboxContext);
  const startTime = Date.now();
  let attempt = 0;
  let lastNotFoundResponse;

  while (Date.now() - startTime <= 20000) {
    attempt += 1;
    const response = await fetch(requestUrl, {
      headers: {
        Authorization: authorizationHeader,
      },
    });

    if (response.ok) {
      const message = await response.json();
      return {
        retriesNeeded: attempt - 1,
        message,
      };
    }

    if (response.status !== 404) {
      throw await createGraphResponseError(response, `Graph lookup failed for saved draft on attempt ${attempt}.`);
    }

    lastNotFoundResponse = await cloneGraphResponseDetails(response, `Saved draft not available in Graph yet on attempt ${attempt}.`);
    console.warn(lastNotFoundResponse.message, lastNotFoundResponse.details);

    if (Date.now() - startTime > 19000) {
      break;
    }

    await delay(2000);
  }

  const timeoutError = new Error("Saved draft was not available through Graph within 20 seconds.");
  timeoutError.graphResponse = lastNotFoundResponse?.details;
  throw timeoutError;
}

async function createGraphResponseError(response, message) {
  const details = await cloneGraphResponseDetails(response, message);
  const error = new Error(message);
  error.graphResponse = details.details;
  return error;
}

async function cloneGraphResponseDetails(response, message) {
  const headers = {};
  response.headers.forEach((value, key) => {
    headers[key] = value;
  });

  const contentType = response.headers.get("content-type") || "";
  const body = contentType.includes("application/json")
    ? await response.json()
    : await response.text();

  return {
    message,
    details: {
      status: response.status,
      statusText: response.statusText,
      headers,
      body,
    },
  };
}

function delay(milliseconds) {
  return new Promise((resolve) => {
    setTimeout(resolve, milliseconds);
  });
}