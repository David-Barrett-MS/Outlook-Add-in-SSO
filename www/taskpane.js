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
const listDistributionListsButton = document.getElementById("listDistributionLists");
const distributionListContainerElement = document.getElementById("distributionListContainer");
const distributionListSelectElement = document.getElementById("distributionListSelect");
const expandDistributionListButton = document.getElementById("expandDistributionList");
const distributionListMembersContainerElement = document.getElementById("distributionListMembersContainer");
const distributionListMembersElement = document.getElementById("distributionListMembers");
const expandRecipientsButton = document.getElementById("expandRecipients");
const expandedRecipientsContainerElement = document.getElementById("expandedRecipientsContainer");
const expandedRecipientsListElement = document.getElementById("expandedRecipientsList");
const saveDraftStatusElement = document.getElementById("saveDraftStatus");
const saveDraftStatusTextElement = document.getElementById("saveDraftStatusText");
const sharedMailboxAddressElement = document.getElementById("sharedMailboxAddress");
const tenantIdElement = document.getElementById("entraTenantId");
const appIdElement = document.getElementById("entraAppId");
const useCommonEndpointRadio = document.getElementById("useCommonEndpoint");
const useTenantIdEndpointRadio = document.getElementById("useTenantIdEndpoint");
const enablePiiLoggingCheckbox = document.getElementById("enablePiiLogging");
const debugConsoleElement = document.getElementById("debugConsole");
const debugConsoleHandleElement = document.getElementById("debugConsoleHandle");
const debugConsoleOutputElement = document.getElementById("debugConsoleOutput");
const debugConsoleClearButton = document.getElementById("debugConsoleClear");

/**
 * The add-in settings object.
 * @type {Office.RoamingSettings}
 */
let addinSettings;
let tenantId;
let applicationId;

initializeDebugConsole();

/**
 * Sets up the bottom-docked debug console: hooks the console.log/info/warn/error/debug methods
 * (and uncaught error/rejection events) so their output is also displayed in the TaskPane, and
 * wires up the vertical drag-to-resize handle and "Clear" button.
 */
function initializeDebugConsole() {
  if (!debugConsoleElement || !debugConsoleOutputElement) {
    return;
  }

  setDebugConsoleHeight(window.innerHeight * 0.25);

  if (debugConsoleClearButton) {
    debugConsoleClearButton.onclick = () => {
      debugConsoleOutputElement.innerHTML = "";
    };
  }

  if (debugConsoleHandleElement) {
    debugConsoleHandleElement.addEventListener("mousedown", startDebugConsoleResize);
  }

  window.addEventListener("resize", () => {
    setDebugConsoleHeight(debugConsoleElement.getBoundingClientRect().height);
  });

  hookConsoleMethods();
}

/**
 * Sets the debug console's height (clamped to sensible min/max values) and adjusts the page's
 * bottom padding so the fixed-position console doesn't cover other TaskPane content.
 */
function setDebugConsoleHeight(heightPx) {
  const minHeight = 60;
  const maxHeight = window.innerHeight - 40;
  const clampedHeight = Math.min(Math.max(heightPx, minHeight), maxHeight);

  debugConsoleElement.style.height = `${clampedHeight}px`;
  document.body.style.paddingBottom = `${clampedHeight}px`;
}

/**
 * Begins a vertical drag-resize of the debug console, tracking mouse movement until mouseup.
 */
function startDebugConsoleResize(event) {
  event.preventDefault();

  const onMouseMove = (moveEvent) => {
    setDebugConsoleHeight(window.innerHeight - moveEvent.clientY);
  };
  const onMouseUp = () => {
    document.removeEventListener("mousemove", onMouseMove);
    document.removeEventListener("mouseup", onMouseUp);
  };

  document.addEventListener("mousemove", onMouseMove);
  document.addEventListener("mouseup", onMouseUp);
}

/**
 * Wraps console.log/info/warn/error/debug so their output is echoed to the debug console panel,
 * and also captures uncaught errors and unhandled promise rejections.
 */
function hookConsoleMethods() {
  ["log", "info", "warn", "error", "debug"].forEach((level) => {
    const originalMethod = console[level] ? console[level].bind(console) : null;
    console[level] = (...args) => {
      if (originalMethod) {
        originalMethod(...args);
      }
      appendDebugConsoleEntry(level, args);
    };
  });

  window.addEventListener("error", (event) => {
    appendDebugConsoleEntry("error", [event.message, event.error]);
  });
  window.addEventListener("unhandledrejection", (event) => {
    appendDebugConsoleEntry("error", ["Unhandled promise rejection:", event.reason]);
  });
}

/**
 * Appends a formatted log entry to the debug console output and scrolls it into view.
 */
function appendDebugConsoleEntry(level, args) {
  if (!debugConsoleOutputElement) {
    return;
  }

  const entry = document.createElement("div");
  entry.className = `debug-console-entry ${level}`;

  const timestamp = document.createElement("span");
  timestamp.className = "debug-console-timestamp";
  timestamp.innerText = new Date().toLocaleTimeString();

  const message = document.createElement("span");
  message.innerText = args.map(formatDebugConsoleArg).join(" ");

  entry.appendChild(timestamp);
  entry.appendChild(message);
  debugConsoleOutputElement.appendChild(entry);
  debugConsoleOutputElement.scrollTop = debugConsoleOutputElement.scrollHeight;
}

/**
 * Formats a single console argument for display as text in the debug console.
 */
function formatDebugConsoleArg(arg) {
  if (arg === undefined) {
    return "undefined";
  }
  if (arg === null) {
    return "null";
  }
  if (typeof arg === "string") {
    return arg;
  }
  if (arg instanceof Error) {
    return arg.stack || arg.message;
  }

  try {
    return JSON.stringify(arg, null, 2);
  } catch (error) {
    return String(arg);
  }
}

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
    if (listDistributionListsButton) {
      listDistributionListsButton.onclick = listDistributionLists;
    }
    if (distributionListSelectElement) {
      distributionListSelectElement.onchange = updateExpandDistributionListButtonState;
    }
    if (expandDistributionListButton) {
      expandDistributionListButton.onclick = expandDistributionList;
    }
    if (expandRecipientsButton) {
      expandRecipientsButton.onclick = expandRecipients;
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

    // Restore authority endpoint radio selection.
    const useCommon = addinSettings.get("useCommonEndpoint");
    const useCommonEndpoint = (useCommon === undefined) ? true : useCommon;
    if (useCommonEndpointRadio) {
      useCommonEndpointRadio.checked = useCommonEndpoint;
      useCommonEndpointRadio.onchange = updateAuthorityEndpoint;
    }
    if (useTenantIdEndpointRadio) {
      useTenantIdEndpointRadio.checked = !useCommonEndpoint;
      useTenantIdEndpointRadio.onchange = updateAuthorityEndpoint;
    }

    // Restore PII logging checkbox.
    const piiLogging = addinSettings.get("piiLogging") ?? false;
    if (enablePiiLoggingCheckbox) {
      enablePiiLoggingCheckbox.checked = piiLogging;
      enablePiiLoggingCheckbox.onchange = updatePiiLogging;
    }
    initialiseAccountManager();

    applyOfficeTheme();
  }
});

function initialiseAccountManager() {
  addinSettings = Office.context.roamingSettings;
  tenantId = addinSettings.get("tenantId");
  applicationId = addinSettings.get("applicationId");

  const useCommon = addinSettings.get("useCommonEndpoint");
  const useCommonEndpoint = (useCommon === undefined) ? true : useCommon;
  const effectiveTenantId = useCommonEndpoint ? undefined : tenantId;
  const piiLogging = addinSettings.get("piiLogging") ?? false;

  console.log("Initializing account manager...");
  console.log("Application ID: " + applicationId);
  console.log("PII logging: " + piiLogging);
  if (effectiveTenantId === undefined) {
    console.log("Tenant ID for auth: common");
  } else {
    console.log("Tenant ID for auth: " + effectiveTenantId);
  }
  accountManager.initialize(applicationId, effectiveTenantId, piiLogging);
}

async function updatePiiLogging() {
  const piiLogging = enablePiiLoggingCheckbox?.checked ?? false;
  console.log("PII logging changed: " + piiLogging);
  addinSettings.set("piiLogging", piiLogging);
  await addinSettings.saveAsync();
  initialiseAccountManager();
}

async function updateAuthorityEndpoint() {
  const useCommonEndpoint = useCommonEndpointRadio?.checked ?? true;
  console.log("Authority endpoint changed. Use common: " + useCommonEndpoint);
  addinSettings.set("useCommonEndpoint", useCommonEndpoint);
  await addinSettings.saveAsync();
  initialiseAccountManager();
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

  console.log("Attempting to retrieve messages from shared mailbox: " + sharedMailboxAddress);

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

/**
 * Retrieves the personal distribution lists from the mailbox contacts (via the Graph beta
 * distributionList resource) and populates the drop-down list with the results.
 */
async function listDistributionLists() {
  try {
    const accessToken = await accountManager.ssoGetToken(["Contacts.Read"]);
    const distributionLists = await getDistributionLists(accessToken);

    distributionListSelectElement.innerHTML = ""; // Clear previous list

    if (distributionLists.length === 0) {
      const noListsOption = document.createElement("option");
      noListsOption.text = "No distribution lists found";
      noListsOption.disabled = true;
      distributionListSelectElement.appendChild(noListsOption);
    } else {
      distributionLists.forEach((distributionList) => {
        const option = document.createElement("option");
        option.value = distributionList.id;
        option.text = distributionList.displayName || "(no name)";
        distributionListSelectElement.appendChild(option);
      });
    }

    if (distributionListContainerElement) {
      distributionListContainerElement.style.visibility = "visible";
    }
    if (distributionListMembersContainerElement) {
      distributionListMembersContainerElement.style.visibility = "hidden";
    }
    updateExpandDistributionListButtonState();
  } catch (error) {
    console.error("Error listing distribution lists.", error);
  }
}

/**
 * Calls the Microsoft Graph beta endpoint to retrieve the personal distribution lists
 * defined in the signed-in user's mailbox contacts.
 */
async function getDistributionLists(accessToken) {
  const authorizationHeader = accessToken.startsWith("Bearer ") ? accessToken : `Bearer ${accessToken}`;
  const requestUrl = "https://graph.microsoft.com/beta/me/distributionLists?$select=id,displayName";

  const response = await fetch(requestUrl, {
    headers: { Authorization: authorizationHeader },
  });

  if (!response.ok) {
    await logGraphErrorResponse(response, "distribution lists");
    throw new Error(`Failed to retrieve distribution lists: ${response.statusText}`);
  }

  const payload = await response.json();
  return payload.value || [];
}

/**
 * Enables the "Expand List" button only when a distribution list is selected in the drop-down.
 */
function updateExpandDistributionListButtonState() {
  if (!expandDistributionListButton) {
    return;
  }

  const selectedOption = distributionListSelectElement?.selectedOptions?.[0];
  expandDistributionListButton.disabled = !selectedOption || !selectedOption.value;
}

/**
 * Recursively expands the selected distribution list (resolving nested distribution lists)
 * and displays the resolved contacts in the members list box.
 */
async function expandDistributionList() {
  const selectedOption = distributionListSelectElement?.selectedOptions?.[0];
  const distributionListId = selectedOption?.value;
  if (!distributionListId) {
    return;
  }

  try {
    const accessToken = await accountManager.ssoGetToken(["Contacts.Read"]);
    const visitedDistributionListIds = new Set();
    const contacts = await expandDistributionListMembers(distributionListId, accessToken, visitedDistributionListIds);

    distributionListMembersElement.innerHTML = ""; // Clear previous list

    if (contacts.length === 0) {
      const noMembersOption = document.createElement("option");
      noMembersOption.text = "No contacts found in this distribution list.";
      noMembersOption.disabled = true;
      distributionListMembersElement.appendChild(noMembersOption);
    } else {
      contacts.forEach((contact) => {
        const option = document.createElement("option");
        option.text = contact.emailAddress ? `${contact.displayName} <${contact.emailAddress}>` : contact.displayName;
        distributionListMembersElement.appendChild(option);
      });
    }

    if (distributionListMembersContainerElement) {
      distributionListMembersContainerElement.style.visibility = "visible";
    }
  } catch (error) {
    console.error("Error expanding distribution list.", error);
  }
}

/**
 * Retrieves the members of a distribution list, recursively expanding any nested (private)
 * distribution lists, and returns a flat, de-duplicated collection of resolved contacts.
 */
async function expandDistributionListMembers(distributionListId, accessToken, visitedDistributionListIds, resolvedContacts = [], seenContactKeys = new Set()) {
  if (visitedDistributionListIds.has(distributionListId)) {
    return resolvedContacts;
  }
  visitedDistributionListIds.add(distributionListId);

  const authorizationHeader = accessToken.startsWith("Bearer ") ? accessToken : `Bearer ${accessToken}`;
  const requestUrl = `https://graph.microsoft.com/beta/me/distributionLists/${encodeURIComponent(distributionListId)}?$expand=members`;

  const response = await fetch(requestUrl, {
    headers: { Authorization: authorizationHeader },
  });

  if (!response.ok) {
    await logGraphErrorResponse(response, `distribution list ${distributionListId}`);
    throw new Error(`Failed to retrieve distribution list members: ${response.statusText}`);
  }

  const distributionList = await response.json();
  const members = distributionList.members || [];

  for (const member of members) {
    if (member.type === "privateDL" && member.memberId) {
      // Nested distribution list: recurse to resolve its members too.
      await expandDistributionListMembers(member.memberId, accessToken, visitedDistributionListIds, resolvedContacts, seenContactKeys);
      continue;
    }

    const emailAddress = member.contact?.emailAddresses?.[0]?.address;
    const displayName = member.displayName || member.contact?.displayName || "(no name)";
    const contactKey = member.memberId || emailAddress || displayName;

    if (seenContactKeys.has(contactKey)) {
      continue;
    }
    seenContactKeys.add(contactKey);

    resolvedContacts.push({ displayName, emailAddress });
  }

  return resolvedContacts;
}

/**
 * Retrieves the current item's To/Cc/Bcc recipients via Office.js, expands any personal
 * distribution list recipients via Graph (recursively resolving nested lists), and displays
 * the flattened, de-duplicated set of contacts in the recipients list box.
 */
async function expandRecipients() {
  try {
    const mailboxItem = Office.context.mailbox?.item;
    if (!mailboxItem) {
      throw new Error("No item is currently open.");
    }

    const [toRecipients, ccRecipients, bccRecipients] = await Promise.all([
      getRecipientsAsync(mailboxItem.to),
      getRecipientsAsync(mailboxItem.cc),
      getRecipientsAsync(mailboxItem.bcc),
    ]);
    const allRecipients = [...toRecipients, ...ccRecipients, ...bccRecipients];

    const accessToken = await accountManager.ssoGetToken(["Contacts.Read"]);
    const visitedDistributionListIds = new Set();
    const seenContactKeys = new Set();
    const expandedContacts = [];

    for (const recipient of allRecipients) {
      if (recipient.recipientType === Office.MailboxEnums.RecipientType.DistributionList) {
        const distributionListId = await findDistributionListIdByDisplayName(recipient.displayName, accessToken);
        if (distributionListId) {
          await expandDistributionListMembers(distributionListId, accessToken, visitedDistributionListIds, expandedContacts, seenContactKeys);
          continue;
        }
        console.warn(`Could not resolve distribution list "${recipient.displayName}" via Graph. Adding as-is.`);
      }

      addUniqueContact(expandedContacts, seenContactKeys, recipient.displayName, recipient.emailAddress);
    }

    expandedRecipientsListElement.innerHTML = ""; // Clear previous list

    if (expandedContacts.length === 0) {
      const noRecipientsOption = document.createElement("option");
      noRecipientsOption.text = "No recipients found on this item.";
      noRecipientsOption.disabled = true;
      expandedRecipientsListElement.appendChild(noRecipientsOption);
    } else {
      expandedContacts.forEach((contact) => {
        const option = document.createElement("option");
        option.text = contact.emailAddress ? `${contact.displayName} <${contact.emailAddress}>` : contact.displayName;
        expandedRecipientsListElement.appendChild(option);
      });
    }

    if (expandedRecipientsContainerElement) {
      expandedRecipientsContainerElement.style.visibility = "visible";
    }
  } catch (error) {
    console.error("Error expanding recipients.", error);
  }
}

/**
 * Wraps the Office.js recipients field getAsync call (used for the to/cc/bcc fields of a
 * compose item) in a Promise. Returns an empty array if the field isn't available.
 */
function getRecipientsAsync(recipientsField) {
  return new Promise((resolve, reject) => {
    if (!recipientsField || typeof recipientsField.getAsync !== "function") {
      resolve([]);
      return;
    }

    recipientsField.getAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value || []);
        return;
      }

      reject(result.error || new Error("Failed to retrieve recipients."));
    });
  });
}

/**
 * Looks up a personal distribution list's Graph id by its display name (as shown in an
 * Office.js recipient), since Office.js does not expose the underlying Graph id directly.
 */
async function findDistributionListIdByDisplayName(displayName, accessToken) {
  if (!displayName) {
    return null;
  }

  const authorizationHeader = accessToken.startsWith("Bearer ") ? accessToken : `Bearer ${accessToken}`;
  const escapedDisplayName = displayName.replace(/'/g, "''");
  const filter = `displayName eq '${escapedDisplayName}'`;
  const requestUrl = `https://graph.microsoft.com/beta/me/distributionLists?$select=id,displayName&$filter=${encodeURIComponent(filter)}`;

  const response = await fetch(requestUrl, {
    headers: { Authorization: authorizationHeader },
  });

  if (!response.ok) {
    await logGraphErrorResponse(response, `distribution list lookup for "${displayName}"`);
    return null;
  }

  const payload = await response.json();
  const matches = payload.value || [];
  return matches.length > 0 ? matches[0].id : null;
}

/**
 * Adds a contact to the resolved list, skipping duplicates (matched by email address, falling
 * back to display name when no email address is available).
 */
function addUniqueContact(contacts, seenContactKeys, displayName, emailAddress) {
  const contactKey = emailAddress || displayName;
  if (seenContactKeys.has(contactKey)) {
    return;
  }
  seenContactKeys.add(contactKey);
  contacts.push({ displayName: displayName || "(no name)", emailAddress });
}

async function saveDraftAndGetViaGraph() {
  setSaveDraftStatus(true, "Saving draft...");
  if (saveDraftAndGetViaGraphButton) {
    saveDraftAndGetViaGraphButton.disabled = true;
  }

  try {
    const mailboxItem = Office.context.mailbox?.item;
    if (!mailboxItem || typeof mailboxItem.saveAsync !== "function") {
      throw new Error("This test requires an Outlook compose item that supports saveAsync (saveAsync function is not available).");
    }

    console.log("Saving current item draft...");
    const savedItemId = await saveCurrentItemAsync(mailboxItem);
    console.log("Draft saved.", { savedItemId });
    const operationStartTime = Date.now();

    const graphMessageId = convertItemIdForGraph(savedItemId);
    const mailboxContext = await getMailboxContextForGraph(mailboxItem);
    console.log("Retrieving saved draft from Graph...", {
      graphMessageId,
      isSharedMailbox: mailboxContext.isShared,
      mailboxAddress: mailboxContext.mailboxAddress,
    });

    setSaveDraftStatus(true, "Waiting for draft to become available via Graph...");
    const scopes = mailboxContext.isShared ? ["Mail.ReadWrite.Shared"] : ["Mail.ReadWrite"];
    const accessToken = await accountManager.ssoGetToken(scopes);
    const authorizationHeader = accessToken.startsWith("Bearer ") ? accessToken : `Bearer ${accessToken}`;
    const result = await getMessageViaGraphWithRetry(graphMessageId, authorizationHeader, mailboxContext, (attempt) => {
      setSaveDraftStatus(true, `Waiting for draft to become available via Graph... (attempt ${attempt})`);
    });
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
  } finally {
    setSaveDraftStatus(false);
    if (saveDraftAndGetViaGraphButton) {
      saveDraftAndGetViaGraphButton.disabled = false;
    }
  }
}

/**
 * Shows or hides the spinning status indicator next to the "Save draft and get via Graph"
 * button, optionally updating its status text.
 */
function setSaveDraftStatus(visible, statusText) {
  if (saveDraftStatusTextElement && statusText) {
    saveDraftStatusTextElement.innerText = statusText;
  }
  if (saveDraftStatusElement) {
    saveDraftStatusElement.style.visibility = visible ? "visible" : "hidden";
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

async function getMessageViaGraphWithRetry(messageId, authorizationHeader, mailboxContext, onAttempt) {
  const requestUrl = buildGraphMessageRequestUrl(messageId, mailboxContext);
  const startTime = Date.now();
  let attempt = 0;
  let lastNotFoundResponse;

  while (Date.now() - startTime <= 20000) {
    attempt += 1;
    if (typeof onAttempt === "function") {
      onAttempt(attempt);
    }
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