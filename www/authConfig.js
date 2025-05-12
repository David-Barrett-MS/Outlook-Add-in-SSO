// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/* This file provides MSAL auth configuration to get access token through nested app authentication. */

import "./lib/msal-browser.js";
export { applicationId, msalConfig, AccountManager };

const applicationId = "9ccd58e7-ce5f-4cbc-b411-80d9b5195d21";

const msalConfig = {
  auth: {
    clientId: applicationId,
    authority: "https://login.microsoftonline.com/77275f64-d6b0-4d6d-b04b-8580417d20a6",
    supportsNestedAppAuth: true,
  },
};

// Encapsulate functions for getting user account and token information.
class AccountManager {
  pca = undefined;

  // Initialize MSAL public client application.
  async initialize() {
    // Initialize the public client application.
    try {
      this.pca = await msal.createNestablePublicClientApplication(msalConfig);
    } catch (error) {
      // All console.log statements write to the runtime log. For more information, see https://learn.microsoft.com/office/dev/add-ins/testing/runtime-logging
      console.log(`Error creating pca: ${error}`);
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