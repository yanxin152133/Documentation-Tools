/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("btn1").onclick = btn1;
  }
});

export async function btn1() {
  return Word.run(async (context) => {
    /**
     * Insert your Word code here
     */

    // insert a paragraph at the end of the document.
    // const paragraph = context.document.body.insertParagraph("Hello World", Word.InsertLocation.end);

    // change the paragraph color to blue.
    // paragraph.font.color = "blue";

    // await context.sync();
    // window.alert is not supported in Office add-ins. Update the taskpane DOM instead.
    const status = document.getElementById("status-msg") || (function () {
      const el = document.createElement("div");
      el.id = "status-msg";
      el.style.cssText = "padding:8px;margin:8px 0;background:#e6f4ff;border:1px solid #b3d7ff;border-radius:4px;";
      document.getElementById("app-body").prepend(el);
      return el;
    })();
    status.textContent = "按钮1被点击";
    console.log("按钮1被点击");
  });
}
