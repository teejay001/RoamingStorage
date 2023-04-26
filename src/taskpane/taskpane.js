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

function saveRoamingSetings(){
  return new Promise((resolve, reject) => {
    Office.context.roamingSettings.saveAsync((asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.error(`Action failed with message ${asyncResult.error.message}`);
        reject(asyncResult.error);
      } else {
        console.log(`Roaming Settings saved with status: ${asyncResult.value}`);
        resolve(asyncResult.value);
      }
    });
  });
}

async function someRandomFunction() {
  try {
    const tVal1 = Office.context.roamingSettings.get("tkey1");
    const tVal2 = Office.context.roamingSettings.get("tkey2");
    console.log(`Taskpane - Vals from roaming setting: tkey1 = ${tVal1} & tkey2 = ${tVal2}`);
  
    const lVal1 = Office.context.roamingSettings.get("lkey1");
    const lVal2 = Office.context.roamingSettings.get("lkey2");
    console.log(`LaunchEvent - Vals from roaming setting: lkey1 = ${lVal1} & lkey2 = ${lVal2}`);
  } catch (error) {
    console.error(error);
  }
}

export async function run() {
  /**
   * Insert your Outlook code here
   */
  console.log("Within Taskpane");
  try {
    Office.context.roamingSettings.set("tkey1", "Taskpane_Val_1");
    Office.context.roamingSettings.set("tkey2", "Taskpane_Val_2");
    await saveRoamingSetings();

    await someRandomFunction();
  } catch (error) {
    console.error(error);
  }
  console.log("Exist Taskpane");
}
