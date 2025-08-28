const numRows = 20,
  numCols = 10;
let sheets = [];
let currentSheetIndex = 0;
let selectedCell = null;
let selectedCellCoord = { row: 1, col: 1 };
let clipboard = { value: null, style: {} };
let zoomLevel = 1.0;
function createSheet(name) {
  let data = [];
  for (let r = 0; r <= numRows; r++) {
    data[r] = [];
    for (let c = 0; c <= numCols; c++) {
      if (r === 0 && c === 0) data[r][c] = "";
      else if (r === 0) data[r][c] = String.fromCharCode(64 + c);
      else if (c === 0) data[r][c] = r;
      else data[r][c] = "";
    }
  }
  return {
    name: name || "Sheet" + (sheets.length + 1),
    data: data,
    formulas: {},
    cellStyles: {},
    merged: {},
    printArea: null,
    comments: {},
    tableStyle: false,
    filter: null,
  };
}

sheets.push(createSheet("Sheet1"));

const ribbonContents = {
  home: `
        <div class="ribbonGroup" id="groupClipboard">
          <h4>Clipboard</h4>
          <button id="cutBtn">Cut</button>
          <button id="copyBtn">Copy</button>
          <button id="pasteBtn">Paste</button>
        </div>
        <div class="ribbonGroup" id="groupFont">
          <h4>Font</h4>
          <select id="fontSelect">
            <option>Aptos Narrow</option>
            <option>Calibri</option>
            <option>Times New Roman</option>
          </select>
          <select id="fontSizeSelect">
            <option>11</option>
            <option>12</option>
            <option>14</option>
            <option>16</option>
          </select>
          <button id="ribbonBoldBtn"><b>B</b></button>
          <button id="italicBtn"><i>I</i></button>
          <button id="underlineBtn"><u>U</u></button>
        </div>
        <div class="ribbonGroup" id="groupAlignment">
          <h4>Alignment</h4>
          <button id="alignLeftBtn">Left</button>
          <button id="alignCenterBtn">Center</button>
          <button id="alignRightBtn">Right</button>
          <button id="mergeCellsBtn">Merge</button>
        </div>
        <div class="ribbonGroup" id="groupNumber">
          <h4>Number</h4>
          <button id="currencyFormatBtn">$</button>
          <button id="percentageFormatBtn">%</button>
          <button id="dateFormatBtn">Date</button>
        </div>
        <div class="ribbonGroup" id="groupStyles">
          <h4>Styles</h4>
          <button id="cellStylesBtn">Cell Styles</button>
          <button id="conditionalFormatBtn">Cond. Format</button>
        </div>
        <div class="ribbonGroup" id="groupFormulas">
          <h4>Formulas</h4>
          <button id="insertFormulaBtn">Insert Formula (SUM)</button>
        </div>
      `,
  insert: `
        <div class="ribbonGroup">
          <h4>Insert</h4>
          <button id="insertChartBtn">Chart</button>
          <button id="insertTableBtn">Table</button>
          <button id="insertPictureBtn">Picture</button>
        </div>
      `,
  pagelayout: `
        <div class="ribbonGroup">
          <h4>Page Layout</h4>
          <button id="pageSetupBtn">Page Setup</button>
          <button id="printAreaBtn">Print Area</button>
        </div>
      `,
  formulas: `
        <div class="ribbonGroup">
          <h4>Formulas</h4>
          <button id="autoSumBtn">AutoSum</button>
          <button id="functionLibraryBtn">Function Library</button>
        </div>
      `,
  data: `
        <div class="ribbonGroup">
          <h4>Data</h4>
          <button id="sortBtn">Sort</button>
          <button id="filterBtn">Filter</button>
          <input type="text" id="filterInput" placeholder="Filter value" style="display:none; width:60px;"/>
          <button id="applyFilterBtn" style="display:none;">Apply</button>
        </div>
      `,
  review: `
        <div class="ribbonGroup">
          <h4>Review</h4>
          <button id="commentsBtn">Comments</button>
        </div>
      `,
  view: `
        <div class="ribbonGroup">
          <h4>View</h4>
          <button id="zoomInBtn">Zoom In</button>
          <button id="zoomOutBtn">Zoom Out</button>
        </div>
      `,
};

document.getElementById("ribbonContent").innerHTML = ribbonContents.home;
attachHomeTabEventListeners();

const ribbonTabs = document.querySelectorAll(".ribbonTab");
ribbonTabs.forEach((tab) => {
  tab.addEventListener("click", function () {
    ribbonTabs.forEach((t) => t.classList.remove("active"));
    this.classList.add("active");
    const tabName = this.getAttribute("data-tab");
    document.getElementById("ribbonContent").innerHTML =
      ribbonContents[tabName] || "";
    if (tabName === "home") attachHomeTabEventListeners();
    else if (tabName === "insert") attachInsertTabEventListeners();
    else if (tabName === "pagelayout") attachPageLayoutTabEventListeners();
    else if (tabName === "formulas") attachFormulasTabEventListeners();
    else if (tabName === "data") attachDataTabEventListeners();
    else if (tabName === "review") attachReviewTabEventListeners();
    else if (tabName === "view") attachViewTabEventListeners();
  });
});

function attachHomeTabEventListeners() {
  document.getElementById("cutBtn").addEventListener("click", function () {
    if (!selectedCell) return;
    clipboard.value = selectedCell.textContent;
    clipboard.style = getCellStyle(
      selectedCellCoord.row,
      selectedCellCoord.col
    );
    updateCell(selectedCellCoord.row, selectedCellCoord.col, "");
    setStatus("Cell cut.");
  });
  document.getElementById("copyBtn").addEventListener("click", function () {
    if (!selectedCell) return;
    clipboard.value = selectedCell.textContent;
    clipboard.style = getCellStyle(
      selectedCellCoord.row,
      selectedCellCoord.col
    );
    setStatus("Cell copied.");
  });
  document.getElementById("pasteBtn").addEventListener("click", function () {
    if (!selectedCell || clipboard.value === null) return;
    updateCell(selectedCellCoord.row, selectedCellCoord.col, clipboard.value);
    setCellStyle(selectedCellCoord.row, selectedCellCoord.col, clipboard.style);
    setStatus("Cell pasted.");
  });
  document.getElementById("fontSelect").addEventListener("change", function () {
    setCellStyle(selectedCellCoord.row, selectedCellCoord.col, {
      fontFamily: this.value,
    });
  });
  document
    .getElementById("fontSizeSelect")
    .addEventListener("change", function () {
      setCellStyle(selectedCellCoord.row, selectedCellCoord.col, {
        fontSize: this.value + "px",
      });
    });
  document
    .getElementById("ribbonBoldBtn")
    .addEventListener("click", function () {
      let current = getCellStyle(
        selectedCellCoord.row,
        selectedCellCoord.col
      ).fontWeight;
      setCellStyle(selectedCellCoord.row, selectedCellCoord.col, {
        fontWeight: current === "bold" ? "normal" : "bold",
      });
    });
  document.getElementById("italicBtn").addEventListener("click", function () {
    let current = getCellStyle(
      selectedCellCoord.row,
      selectedCellCoord.col
    ).fontStyle;
    setCellStyle(selectedCellCoord.row, selectedCellCoord.col, {
      fontStyle: current === "italic" ? "normal" : "italic",
    });
  });
  document
    .getElementById("underlineBtn")
    .addEventListener("click", function () {
      let current = getCellStyle(
        selectedCellCoord.row,
        selectedCellCoord.col
      ).textDecoration;
      setCellStyle(selectedCellCoord.row, selectedCellCoord.col, {
        textDecoration: current === "underline" ? "none" : "underline",
      });
    });
  document
    .getElementById("alignLeftBtn")
    .addEventListener("click", function () {
      setCellStyle(selectedCellCoord.row, selectedCellCoord.col, {
        textAlign: "left",
      });
    });
  document
    .getElementById("alignCenterBtn")
    .addEventListener("click", function () {
      setCellStyle(selectedCellCoord.row, selectedCellCoord.col, {
        textAlign: "center",
      });
    });
  document
    .getElementById("alignRightBtn")
    .addEventListener("click", function () {
      setCellStyle(selectedCellCoord.row, selectedCellCoord.col, {
        textAlign: "right",
      });
    });
  document
    .getElementById("mergeCellsBtn")
    .addEventListener("click", function () {
      mergeCells();
    });
  document
    .getElementById("currencyFormatBtn")
    .addEventListener("click", function () {
      let val = parseFloat(selectedCell.textContent);
      if (!isNaN(val))
        updateCell(
          selectedCellCoord.row,
          selectedCellCoord.col,
          "$" + val.toFixed(2)
        );
    });
  document
    .getElementById("percentageFormatBtn")
    .addEventListener("click", function () {
      let val = parseFloat(selectedCell.textContent);
      if (!isNaN(val))
        updateCell(
          selectedCellCoord.row,
          selectedCellCoord.col,
          (val * 100).toFixed(0) + "%"
        );
    });
  document
    .getElementById("dateFormatBtn")
    .addEventListener("click", function () {
      let d = new Date(selectedCell.textContent);
      if (d.toString() !== "Invalid Date")
        updateCell(
          selectedCellCoord.row,
          selectedCellCoord.col,
          d.toLocaleDateString()
        );
    });
  document
    .getElementById("cellStylesBtn")
    .addEventListener("click", function () {
      showExtraPanel(`<label>Cell Fill:</label>
          <input type="color" id="bgColorPicker" value="#ffffff" />
          <button id="applyBgColor">Apply</button>
          <button id="closeExtraPanel">Close</button>`);
      document
        .getElementById("applyBgColor")
        .addEventListener("click", function () {
          const color = document.getElementById("bgColorPicker").value;
          setCellStyle(selectedCellCoord.row, selectedCellCoord.col, {
            backgroundColor: color,
          });
          hideExtraPanel();
        });
      document
        .getElementById("closeExtraPanel")
        .addEventListener("click", hideExtraPanel);
    });
  document
    .getElementById("conditionalFormatBtn")
    .addEventListener("click", function () {
      let val = parseFloat(selectedCell.textContent);
      if (!isNaN(val) && val < 0)
        setCellStyle(selectedCellCoord.row, selectedCellCoord.col, {
          backgroundColor: "#ffcccc",
        });
      else
        setCellStyle(selectedCellCoord.row, selectedCellCoord.col, {
          backgroundColor: "",
        });
    });
  document
    .getElementById("insertFormulaBtn")
    .addEventListener("click", function () {
      document.getElementById("formulaBar").value = "=SUM(";
      document.getElementById("formulaBar").focus();
    });
}

function attachInsertTabEventListeners() {
  document
    .getElementById("insertChartBtn")
    .addEventListener("click", function () {
      if (!selectedCell) return;
      let col = selectedCellCoord.col;
      let values = [];
      const sheet = sheets[currentSheetIndex];
      for (let r = 1; r < sheet.data.length; r++) {
        let val = parseFloat(sheet.data[r][col]);
        if (!isNaN(val)) values.push(val);
      }
      if (values.length === 0) {
        setStatus("No numeric data in column.");
        return;
      }
      showChart(values);
    });
  document
    .getElementById("insertTableBtn")
    .addEventListener("click", function () {
      let sheet = sheets[currentSheetIndex];
      sheet.tableStyle = !sheet.tableStyle;
      setStatus("Table style " + (sheet.tableStyle ? "applied." : "removed."));
      renderSpreadsheet();
    });
  document
    .getElementById("insertPictureBtn")
    .addEventListener("click", function () {
      showExtraPanel(`<label>Image URL:</label>
          <input type="text" id="imgUrl" style="width:200px;" />
          <button id="insertImgBtn">Insert Image</button>
          <button id="closeExtraPanel">Close</button>`);
      document
        .getElementById("insertImgBtn")
        .addEventListener("click", function () {
          let url = document.getElementById("imgUrl").value;
          if (url)
            updateCell(
              selectedCellCoord.row,
              selectedCellCoord.col,
              `<img src="${url}" style="max-width:100%;max-height:100%;" />`
            );
          hideExtraPanel();
        });
      document
        .getElementById("closeExtraPanel")
        .addEventListener("click", hideExtraPanel);
    });
}

function attachPageLayoutTabEventListeners() {
  document
    .getElementById("pageSetupBtn")
    .addEventListener("click", function () {
      showExtraPanel(`
          <label>Top Margin (in):</label><input type="number" id="topMargin" value="1" step="0.1"/><br/>
          <label>Bottom Margin (in):</label><input type="number" id="bottomMargin" value="1" step="0.1"/><br/>
          <label>Left Margin (in):</label><input type="number" id="leftMargin" value="0.75" step="0.1"/><br/>
          <label>Right Margin (in):</label><input type="number" id="rightMargin" value="0.75" step="0.1"/><br/>
          <button id="applyPageSetup">Apply</button>
          <button id="closeExtraPanel">Close</button>
        `);
      document
        .getElementById("applyPageSetup")
        .addEventListener("click", function () {
          let sheet = sheets[currentSheetIndex];
          sheet.printSetup = {
            top: document.getElementById("topMargin").value,
            bottom: document.getElementById("bottomMargin").value,
            left: document.getElementById("leftMargin").value,
            right: document.getElementById("rightMargin").value,
          };
          setStatus("Page setup applied.");
          hideExtraPanel();
        });
      document
        .getElementById("closeExtraPanel")
        .addEventListener("click", hideExtraPanel);
    });
  document
    .getElementById("printAreaBtn")
    .addEventListener("click", function () {
      let sheet = sheets[currentSheetIndex];
      let startRow = numRows,
        endRow = 0,
        startCol = numCols,
        endCol = 0;
      for (let r = 1; r < sheet.data.length; r++) {
        for (let c = 1; c < sheet.data[r].length; c++) {
          if (sheet.data[r][c] !== "") {
            if (r < startRow) startRow = r;
            if (r > endRow) endRow = r;
            if (c < startCol) startCol = c;
            if (c > endCol) endCol = c;
          }
        }
      }
      if (endRow === 0) {
        setStatus("No print area defined.");
        return;
      }
      sheet.printArea = { startRow, endRow, startCol, endCol };
      setStatus("Print area set.");
      renderSpreadsheet();
    });
}

function attachFormulasTabEventListeners() {
  document.getElementById("autoSumBtn").addEventListener("click", function () {
    let col = selectedCellCoord.col;
    let sum = 0;
    const sheet = sheets[currentSheetIndex];
    for (let r = selectedCellCoord.row - 1; r >= 1; r--) {
      let val = parseFloat(sheet.data[r][col]);
      if (isNaN(val)) break;
      sum += val;
    }
    updateCell(selectedCellCoord.row, col, sum);
    setStatus("AutoSum applied.");
  });
  document
    .getElementById("functionLibraryBtn")
    .addEventListener("click", function () {
      showExtraPanel(`
          <p>Available Functions:</p>
          <ul>
            <li>SUM(range)</li>
            <li>AVERAGE(range)</li>
            <li>MIN(range)</li>
            <li>MAX(range)</li>
            <li>ROUND(number, places)</li>
            <!-- Add more as needed -->
          </ul>
          <button id="closeExtraPanel">Close</button>
        `);
      document
        .getElementById("closeExtraPanel")
        .addEventListener("click", hideExtraPanel);
    });
}

function attachDataTabEventListeners() {
  document.getElementById("sortBtn").addEventListener("click", function () {
    let col = selectedCellCoord.col;
    let sheet = sheets[currentSheetIndex];
    let header = sheet.data[0];
    let rows = sheet.data.slice(1);
    rows.sort((a, b) => {
      let aVal = a[col],
        bVal = b[col];
      return isNaN(aVal) || isNaN(bVal)
        ? aVal.localeCompare(bVal)
        : parseFloat(aVal) - parseFloat(bVal);
    });
    sheet.data = [header, ...rows];
    setStatus("Data sorted by column " + String.fromCharCode(64 + col) + ".");
    renderSpreadsheet();
  });
  document.getElementById("filterBtn").addEventListener("click", function () {
    let inp = document.getElementById("filterInput");
    let btn = document.getElementById("applyFilterBtn");
    if (inp.style.display === "none") {
      inp.style.display = "inline-block";
      btn.style.display = "inline-block";
    } else {
      inp.style.display = "none";
      btn.style.display = "none";
      sheets[currentSheetIndex].filter = null;
      renderSpreadsheet();
    }
  });
  document
    .getElementById("applyFilterBtn")
    .addEventListener("click", function () {
      let val = document.getElementById("filterInput").value;
      sheets[currentSheetIndex].filter = {
        col: selectedCellCoord.col,
        value: val,
      };
      setStatus(
        "Filter applied on column " +
          String.fromCharCode(64 + selectedCellCoord.col) +
          "."
      );
      renderSpreadsheet();
    });
}

function attachReviewTabEventListeners() {
  document.getElementById("commentsBtn").addEventListener("click", function () {
    let sheet = sheets[currentSheetIndex];
    let key = getCellKey(selectedCellCoord.row, selectedCellCoord.col);
    let existing = sheet.comments[key] || "";
    showExtraPanel(`
          <label>Comment for ${key}:</label><br/>
          <textarea id="commentEditor" rows="3" style="width:300px;">${existing}</textarea><br/>
          <button id="saveCommentBtn">Save Comment</button>
          <button id="closeExtraPanel">Close</button>
        `);
    document
      .getElementById("saveCommentBtn")
      .addEventListener("click", function () {
        sheet.comments[key] = document.getElementById("commentEditor").value;
        setStatus("Comment saved for " + key + ".");
        hideExtraPanel();
        renderSpreadsheet();
      });
    document
      .getElementById("closeExtraPanel")
      .addEventListener("click", hideExtraPanel);
  });
}

function attachViewTabEventListeners() {
  document.getElementById("zoomInBtn").addEventListener("click", function () {
    zoomLevel += 0.1;
    updateZoom();
  });
  document.getElementById("zoomOutBtn").addEventListener("click", function () {
    zoomLevel = Math.max(0.5, zoomLevel - 0.1);
    updateZoom();
  });
}
function updateZoom() {
  document.getElementById("spreadsheetContainer").style.transform =
    "scale(" + zoomLevel + ")";
  document.getElementById("zoomControl").textContent =
    Math.round(zoomLevel * 100) + "%";
}

function getCellKey(row, col) {
  return String.fromCharCode(64 + col) + row;
}
function getCellStyle(row, col) {
  const key = getCellKey(row, col);
  return sheets[currentSheetIndex].cellStyles[key] || {};
}
function setCellStyle(row, col, newStyles) {
  const key = getCellKey(row, col);
  let current = sheets[currentSheetIndex].cellStyles[key] || {};
  sheets[currentSheetIndex].cellStyles[key] = { ...current, ...newStyles };
  renderSpreadsheet();
}

function mergeCells() {
  const sheet = sheets[currentSheetIndex];
  const key = getCellKey(selectedCellCoord.row, selectedCellCoord.col);
  if (sheet.merged[key]) {
    delete sheet.merged[key];
    setStatus("Cell unmerged.");
    renderSpreadsheet();
    return;
  }
  if (selectedCellCoord.col >= numCols) {
    setStatus("Cannot merge: no right neighbor.");
    return;
  }
  const rightKey = getCellKey(selectedCellCoord.row, selectedCellCoord.col + 1);
  sheet.merged[key] = { colspan: 2, rowspan: 1 };
  setStatus("Cells merged.");
  renderSpreadsheet();
}

function updateCell(row, col, value) {
  sheets[currentSheetIndex].data[row][col] = value;
  delete sheets[currentSheetIndex].formulas[getCellKey(row, col)];
  renderSpreadsheet();
}

function renderSpreadsheet() {
  const sheet = sheets[currentSheetIndex];
  const table = document.getElementById("spreadsheet");
  table.innerHTML = "";
  for (let r = 0; r < sheet.data.length; r++) {
    const tr = document.createElement("tr");
    for (let c = 0; c < sheet.data[r].length; c++) {
      const currentKey = getCellKey(r, c);
      let skip = false;
      for (let mk in sheet.merged) {
        let mergeInfo = sheet.merged[mk];
        let mRow = parseInt(mk.slice(1));
        let mCol = mk.charCodeAt(0) - 64;
        if (r === mRow && c > mCol && c < mCol + mergeInfo.colspan) {
          skip = true;
          break;
        }
      }
      if (skip) continue;
      let cellElem = document.createElement(r === 0 || c === 0 ? "th" : "td");
      cellElem.dataset.row = r;
      cellElem.dataset.col = c;
      if (r === 0 || c === 0) {
        cellElem.textContent = sheet.data[r][c];
      } else {
        const key = getCellKey(r, c);
        let cellContent = sheet.formulas[key]
          ? evaluateFormula(sheet.formulas[key], sheet)
          : sheet.data[r][c];
        cellElem.innerHTML = cellContent;
        let styleObj = sheet.cellStyles[key];
        if (styleObj) {
          Object.assign(cellElem.style, styleObj);
        }
        if (sheet.comments[key]) {
          let commSpan = document.createElement("span");
          commSpan.textContent = "💬";
          commSpan.style.position = "absolute";
          commSpan.style.bottom = "0";
          commSpan.style.right = "0";
          commSpan.style.fontSize = "8px";
          cellElem.appendChild(commSpan);
        }
        if (
          sheet.printArea &&
          r >= sheet.printArea.startRow &&
          r <= sheet.printArea.endRow &&
          c >= sheet.printArea.startCol &&
          c <= sheet.printArea.endCol
        ) {
          cellElem.style.border = "2px solid #4285F4";
        }
        cellElem.addEventListener("click", cellClickHandler);
        if (sheet.tableStyle && r % 2 === 0) {
          cellElem.style.backgroundColor = "#f9f9f9";
        }
      }
      if (sheet.merged[currentKey]) {
        cellElem.colSpan = sheet.merged[currentKey].colspan;
        cellElem.rowSpan = sheet.merged[currentKey].rowspan;
      }
      if (r === selectedCellCoord.row || c === selectedCellCoord.col) {
        cellElem.classList.add("highlighted");
      }
      if (
        selectedCell &&
        parseInt(cellElem.dataset.row) === selectedCellCoord.row &&
        parseInt(cellElem.dataset.col) === selectedCellCoord.col
      ) {
        cellElem.classList.add("selected");
      }
      tr.appendChild(cellElem);
    }
    table.appendChild(tr);
  }
}

function evaluateFormula(formula, sheet = sheets[currentSheetIndex]) {
  if (formula[0] === "=") formula = formula.slice(1);
  formula = formula.replace(/([A-Z]+)(\d+)/g, function (match, p1, p2) {
    const col = p1.charCodeAt(0) - 64;
    const row = parseInt(p2);
    const key = p1 + row;
    let value = sheet.formulas[key]
      ? evaluateFormula(sheet.formulas[key], sheet)
      : sheet.data[row][col];
    return value === "" ? 0 : value;
  });
  try {
    const SUM = (...args) =>
      args.flat().reduce((a, b) => a + parseFloat(b || 0), 0);
    return eval(formula);
  } catch (e) {
    return "#ERROR";
  }
}

function cellClickHandler(e) {
  if (selectedCell) selectedCell.classList.remove("selected");
  selectedCell = e.currentTarget;
  selectedCell.classList.add("selected");
  selectedCellCoord = {
    row: parseInt(selectedCell.dataset.row),
    col: parseInt(selectedCell.dataset.col),
  };
  const key = getCellKey(selectedCellCoord.row, selectedCellCoord.col);
  document.getElementById("formulaBar").value =
    sheets[currentSheetIndex].formulas[key] ||
    sheets[currentSheetIndex].data[selectedCellCoord.row][
      selectedCellCoord.col
    ];
  document.getElementById("nameBox").value = key;
  renderSpreadsheet();
}

document.getElementById("formulaBar").addEventListener("keydown", function (e) {
  if (e.key === "Enter" && selectedCell) {
    const row = selectedCellCoord.row;
    const col = selectedCellCoord.col;
    const key = getCellKey(row, col);
    const value = this.value;
    if (value.startsWith("=")) {
      sheets[currentSheetIndex].formulas[key] = value;
      sheets[currentSheetIndex].data[row][col] = "";
    } else {
      sheets[currentSheetIndex].data[row][col] = value;
      delete sheets[currentSheetIndex].formulas[key];
    }
    renderSpreadsheet();
  }
});

function showExtraPanel(html) {
  const panel = document.getElementById("extraPanel");
  panel.innerHTML = html;
  panel.style.display = "block";
}
function hideExtraPanel() {
  document.getElementById("extraPanel").style.display = "none";
}

function showChart(values) {
  const modal = document.getElementById("chartModal");
  modal.style.display = "block";
  const canvas = document.getElementById("chartCanvas");
  const ctx = canvas.getContext("2d");
  ctx.clearRect(0, 0, canvas.width, canvas.height);
  const maxVal = Math.max(...values);
  const barWidth = canvas.width / values.length;
  values.forEach((val, i) => {
    let barHeight = (val / maxVal) * (canvas.height - 20);
    ctx.fillStyle = "#4CAF50";
    ctx.fillRect(
      i * barWidth,
      canvas.height - barHeight,
      barWidth - 2,
      barHeight
    );
    ctx.fillStyle = "#000";
    ctx.font = "10px Arial";
    ctx.fillText(val, i * barWidth + 2, canvas.height - barHeight - 2);
  });
}
document.getElementById("closeChartBtn").addEventListener("click", function () {
  document.getElementById("chartModal").style.display = "none";
});

function renderSheetTabs() {
  const sheetTabsDiv = document.getElementById("sheetTabs");
  sheetTabsDiv
    .querySelectorAll("button.sheetTab")
    .forEach((btn) => btn.remove());
  sheets.forEach((sheet, index) => {
    const btn = document.createElement("button");
    btn.textContent = sheet.name;
    btn.className = "sheetTab" + (index === currentSheetIndex ? " active" : "");
    btn.setAttribute("data-index", index);
    btn.addEventListener("click", function () {
      currentSheetIndex = parseInt(this.getAttribute("data-index"));
      selectedCell = null;
      selectedCellCoord = { row: 1, col: 1 };
      document.getElementById("formulaBar").value = "";
      document.getElementById("nameBox").value = "A1";
      renderSpreadsheet();
      renderSheetTabs();
    });
    sheetTabsDiv.insertBefore(btn, document.getElementById("addSheetBtn"));
  });
}
document.getElementById("addSheetBtn").addEventListener("click", function () {
  sheets.push(createSheet());
  currentSheetIndex = sheets.length - 1;
  renderSheetTabs();
  renderSpreadsheet();
});
renderSheetTabs();

document.getElementById("saveBtn").addEventListener("click", exportCSV);
document.getElementById("exportCsvBtn").addEventListener("click", exportCSV);
document.getElementById("importCsvBtn").addEventListener("click", function () {
  document.getElementById("importFile").click();
});
document.getElementById("importFile").addEventListener("change", function (e) {
  const file = e.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = function (event) {
    importCSV(event.target.result);
  };
  reader.readAsText(file);
});
document.getElementById("undoBtn").addEventListener("click", function () {
  setStatus("Undo not implemented.");
});
document.getElementById("redoBtn").addEventListener("click", function () {
  setStatus("Redo not implemented.");
});

function exportCSV() {
  const sheet = sheets[currentSheetIndex];
  let csvContent = "";
  for (let r = 0; r < sheet.data.length; r++) {
    csvContent += sheet.data[r].join(",") + "\n";
  }
  const blob = new Blob([csvContent], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = sheets[currentSheetIndex].name + ".csv";
  a.click();
  URL.revokeObjectURL(url);
  setStatus("CSV exported.");
}
function importCSV(csv) {
  const rows = csv
    .trim()
    .split("\n")
    .map((row) => row.split(","));
  const sheet = sheets[currentSheetIndex];
  sheet.data = rows;
  sheet.formulas = {};
  sheet.cellStyles = {};
  renderSpreadsheet();
  setStatus("CSV imported.");
}

function setStatus(msg) {
  document.getElementById("statusText").textContent = msg;
  setTimeout(() => {
    document.getElementById("statusText").textContent = "Ready";
  }, 3000);
}

renderSpreadsheet();