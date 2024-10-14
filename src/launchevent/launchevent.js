/* eslint-disable no-undef */
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

function onMessageSendHandler(event) {
  event.completed({
    allowEvent: false,
    cancelLabel: "Save",
    commandId: "msgComposeOpenPaneButton",
    errorMessageMarkdown: "**Would You Like To Save A Copy Of This Email Before Sending?**",
  });
}

// IMPORTANT: To ensure your add-in is supported in Outlook, remember to map the event handler name specified in the manifest to its JavaScript counterpart.
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
