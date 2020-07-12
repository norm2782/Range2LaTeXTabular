/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady(info => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

function transpose(xs) {
  return xs[0].map((_, colIndex) => xs.map(row => row[colIndex]));
}

function widestWidth(xs) {
  return xs.map(x => x.length)
           .reduce(((acc, n) => Math.max(acc, n)), 0)
}

function fillCells(maxLength, cellData) {
  return cellData + "&nbsp;".repeat(maxLength - cellData.length)
}

function mkRow(widths, row) {
  return row.map((cellData, idx) => fillCells(widths[idx], cellData))
            .join(" & ") + " \\\\<br />"
}

export async function run() {
  try {
    await Excel.run(async context => {
      /**
       * Insert your Excel code here
       */
      const range = context.workbook.getSelectedRange();

      range.load("values");

      context.sync()
        .then(function () {
          let originalRange = range.values;
          let widths        = transpose(originalRange).map(widestWidth);
          let latexTabular  = "\\begin{tabular}{}<br />"
                            + originalRange.map(row => mkRow(widths, row)).join("")
                            + "\\end{tabular}";

          document.getElementById("latex-target").innerHTML = latexTabular;
        });

    });
  } catch (error) {
    console.error(error);
  }
}
