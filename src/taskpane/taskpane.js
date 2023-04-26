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

async function someRandomFunction() {
  const tVal1 = Office.context.roamingSettings.get("tkey1");
  const tVal2 = Office.context.roamingSettings.get("tkey2");
  console.log(`Taskpane - Vals from roaming setting: tkey1 = ${tVal1} & tkey2 = ${tVal2}`);

  const lVal1 = Office.context.roamingSettings.get("lkey1");
  const lVal2 = Office.context.roamingSettings.get("lkey2");
  console.log(`LaunchEvent - Vals from roaming setting: lkey1 = ${lVal1} & lkey2 = ${lVal2}`);
}

export async function run() {
  /**
   * Insert your Outlook code here
   */
  console.log("Within Taskpane");
  Office.context.roamingSettings.set("tkey1", "Taskpane_Val_1");
  Office.context.roamingSettings.set("tkey2", "Taskpane_Val_2");

  await someRandomFunction();
  console.log("Exist Taskpane");
}
