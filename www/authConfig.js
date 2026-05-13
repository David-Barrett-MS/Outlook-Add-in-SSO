// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/* This file provides MSAL auth configuration to get access token through nested app authentication. */

import "./lib/msal-browser.js";
export { AccountManager };


const msalConfig = {
  auth: {
    clientId: "0",
    authority: "https://login.microsoftonline.com/",
    supportsNestedAppAuth: true,
  },
  system: {
    loggerOptions: {
      loggerCallback: (level, message, containsPii) => {
        if (containsPii) return;
        switch (level) {
          case msal.LogLevel.Error:
            console.error(`[MSAL] Error: ${message}`);
            break;
          case msal.LogLevel.Warning:
            console.warn(`[MSAL] Warning: ${message}`);
            break;
          case msal.LogLevel.Info:
            console.info(`[MSAL] Info: ${message}`);
            break;
          case msal.LogLevel.Verbose:
            console.debug(`[MSAL] Verbose: ${message}`);
            break;
          case msal.LogLevel.Trace:
            console.debug(`[MSAL] Trace: ${message}`);
            break;
        }
      },
      logLevel: msal.LogLevel.Trace,
      piiLoggingEnabled: false,
    },
  },
};

// Encapsulate functions for getting user account and token information.
class AccountManager {
  pca = undefined;

  // Initialize MSAL public client application.
  async initialize(EntraApplicationId, EntraTenantId) {
    // Initialize the public client application.
    if (EntraTenantId !== undefined) {
      msalConfig.auth.authority = `https://login.microsoftonline.com/${EntraTenantId}`;
    }
    else {
      msalConfig.auth.authority = 'https://login.microsoftonline.com/common';
    }
    if (EntraApplicationId !== undefined) {
      msalConfig.auth.clientId = EntraApplicationId;
    }
    else {
      throw new Error("EntraApplicationId is not defined!");
    }

    try {
      this.pca = await msal.createNestablePublicClientApplication(msalConfig);
      console.log(`Successfully created pca with authority: ${msalConfig.auth.authority} and clientId: ${msalConfig.auth.clientId}`);
    } catch (error) {
      // All console.log statements write to the runtime log. For more information, see https://learn.microsoft.com/office/dev/add-ins/testing/runtime-logging
      console.log(`Error creating pca with authority: ${msalConfig.auth.authority} and clientId: ${msalConfig.auth.clientId}: ${error}`);
    }
  }

  async ssoGetToken(scopes) {
    const userAccount = await this.ssoGetUserIdentity(scopes);
    return userAccount.accessToken;
  }

  /**
   * Uses MSAL and nested app authentication to get the user account from Office SSO.
   * This demonstrates how to work with user identity from the token.
   *
   * @returns The user account data (identity).
   */
  async ssoGetUserIdentity(scopes) {
    if (this.pca === undefined) {
      throw new Error("AccountManager is not initialized!");
    }

    // Specify minimum scopes needed for the access token.
    const tokenRequest = {
      scopes: scopes
    };

    try {
      console.log("Trying to acquire token silently...");
      const userAccount = await this.pca.acquireTokenSilent(tokenRequest);
      console.log("Acquired token silently.");
      return userAccount;
    } catch (error) {
      console.log(`Unable to acquire token silently: ${error}`);
    }

    // Acquire token silent failure. Send an interactive request via popup.
    try {
      console.log("Trying to acquire token interactively...");
      const userAccount = await this.pca.acquireTokenPopup(tokenRequest);
      console.log("Acquired token interactively.");
      return userAccount;
    } catch (popupError) {
      // Acquire token interactive failure.
      console.log(`Unable to acquire token interactively: ${popupError}`);
      throw new Error(`Unable to acquire access token: ${popupError}`);
    }
  }
}