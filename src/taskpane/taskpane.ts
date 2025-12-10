/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import { AppController } from "./services/AppController";
import { UiService } from "./services/UiService";

// Global initialization
let initialized = false;

const initializeTaskpane = async () => {
  if (initialized) return;

  if (typeof document !== "undefined" && document.getElementById("analyzeButton")) {
    try {
      initialized = true;
      // console.log("[Followup Suggester] Initializing task pane..."); // Removed for production readiness

      const uiService = new UiService();
      const appController = new AppController(uiService);
      
      await appController.initialize();
    } catch (e) {
      console.error("Taskpane initialization failed:", e);
      initialized = false;
    }
  }
};

if (typeof Office !== "undefined" && typeof Office.onReady === "function") {
  Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
      initializeTaskpane();
    } else {
        // Fallback for development/testing outside Outlook
        initializeTaskpane();
    }
  });
} else {
  // Browser fallback
  if (typeof document !== "undefined") {
    if (document.readyState === "loading") {
      document.addEventListener("DOMContentLoaded", initializeTaskpane);
    } else {
      initializeTaskpane();
    }
  }
}
