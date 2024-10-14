/* eslint-disable no-undef */
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

export async function run() {
  const item = Office.context.mailbox.item;

  // Get the current date and time
  const currentDate = new Date();
  const formattedDate = currentDate.toLocaleDateString() + " " + currentDate.toLocaleTimeString();

  // Get the first recipient from the 'To' field asynchronously
  item.to.getAsync(function (toResult) {
    if (toResult.status === Office.AsyncResultStatus.Succeeded) {
      // Extract the first name from the recipient's display name
      const recipient =
        toResult.value.length > 0
          ? toResult.value[0].displayName.split(" ")[0].replace(/\s+/g, "_")
          : "UnknownRecipient";

      // Get Subject Asynchronously
      item.subject.getAsync(function (subjectResult) {
        if (subjectResult.status === Office.AsyncResultStatus.Succeeded) {
          // Replace spaces in the subject with underscores
          const subject = subjectResult.value.replace(/\s+/g, "_");

          // Get CC and BCC recipients asynchronously
          item.cc.getAsync(function (ccResult) {
            if (ccResult.status === Office.AsyncResultStatus.Succeeded) {
              const ccRecipients = ccResult.value.map((recipient) => recipient.emailAddress);

              item.bcc.getAsync(function (bccResult) {
                if (bccResult.status === Office.AsyncResultStatus.Succeeded) {
                  const bccRecipients = bccResult.value.map((recipient) => recipient.emailAddress);

                  // Get the body of the message
                  item.body.getAsync(Office.CoercionType.Text, function (bodyResult) {
                    if (bodyResult.status === Office.AsyncResultStatus.Succeeded) {
                      const body = bodyResult.value;

                      // Create the email data content without "Body:" and leave space under BCC
                      const emailContent = `
Date: ${formattedDate}
Subject: ${subject}
CC: ${ccRecipients.join(", ")}
BCC: ${bccRecipients.join(", ")}

${body}
                      `;

                      // Create a file name using only the first name of the recipient and the subject, replacing invalid characters with "_"
                      const fileName = `${recipient}|${subject}`.replace(/[^a-zA-Z0-9_-]/g, "_");

                      // Download the file with the dynamic name
                      downloadEmail(emailContent, fileName);
                    } else {
                      console.error("Error retrieving body: " + bodyResult.error.message);
                    }
                  });
                } else {
                  console.error("Error retrieving BCC: " + bccResult.error.message);
                }
              });
            } else {
              console.error("Error retrieving CC: " + ccResult.error.message);
            }
          });
        } else {
          console.error("Error retrieving subject: " + subjectResult.error.message);
        }
      });
    } else {
      console.error("Error retrieving recipient: " + toResult.error.message);
    }
  });
}

function downloadEmail(content, fileName) {
  // Create a Blob from the content and trigger a download
  const blob = new Blob([content], { type: "text/plain" });
  const link = document.createElement("a");
  link.href = window.URL.createObjectURL(blob);
  link.download = `${fileName}.msg`; // You can also use .eml if needed
  link.click();
}
