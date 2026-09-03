/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office */

Office.onReady(() => {
  // If needed, Office.js is ready to be called.
});

function action(event: Office.AddinCommands.Event) {
  event.completed();
}

// Register the function with Office.
Office.actions.associate("action", action);
