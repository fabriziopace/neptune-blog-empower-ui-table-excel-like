function enableExcelFnToUiTable() {
  pageMaster.addEventDelegate({
    onAfterRendering: function () {
      const tableColumns = tableExcel.getColumns();

      let objTableData = {};
      let arrayTableData = [];

      // reset table data
      modeltableExcel.setData([]);

      // arrange some example data
      let peoplesArr = ["Fabrizio Pace", "User Example", "Another People", "Other One", "Hello World"];
      for (var a = 0; a < 5; a++) {
        objTableData.id = a;
        objTableData.firstName = peoplesArr[a].split(" ")[0];
        objTableData.lastName = peoplesArr[a].split(" ")[1];
        objTableData.country = "Italy";
        objTableData.interests = "Neptune Software, SAPUI5..";
        arrayTableData.push({ ...objTableData });
      }

      tableExcel.bindRows("/");
      modeltableExcel.setData(arrayTableData);

      // onkeydown event for arrow keys / enter navigation
      document.getElementById(tableExcel.sId).onkeydown = function (e) {
        // on key down check if the focused element is the input
        if (e.target) {
          if (e.target.classList) {
            if (e.target.classList.contains("sapMInputBaseInner")) {
              if (e.key === "ArrowLeft" || e.key === "ArrowUp" || e.key === "ArrowRight" || e.key === "ArrowDown" || e.key === "Enter") {
                // The closest() method of the Element interface traverses the element and its parents (heading toward the document root)
                // until it finds a node that matches the specified CSS selector.
                // Ref. https://developer.mozilla.org/en-US/docs/Web/API/Element/closest
                let currentRowIndex = e.target.closest(".sapUiTableRow") ? e.target.closest(".sapUiTableRow").rowIndex - 1 : 0;
                let currentCellIndex = e.target.closest(".sapUiTableCell") ? e.target.closest(".sapUiTableCell").cellIndex : 0;

                let newRowIndex = 0;
                let newCellIndex = 0;

                // calculate the new index based on the current row / cell index selected
                // key arrow left
                if (e.key === "ArrowLeft") {
                  newRowIndex = currentRowIndex;
                  newCellIndex = currentCellIndex > 0 ? currentCellIndex - 1 : currentCellIndex;
                }

                // key arrow up
                if (e.key === "ArrowUp") {
                  newRowIndex = currentRowIndex > 0 ? currentRowIndex - 1 : currentRowIndex;
                  newCellIndex = currentCellIndex;
                }

                // key arrow right
                if (e.key === "ArrowRight") {
                  newRowIndex = currentRowIndex;
                  newCellIndex = currentCellIndex < tableColumns.length ? currentCellIndex + 1 : currentCellIndex;
                }

                // key arrow down / key enter
                if (e.key === "ArrowDown" || e.key === "Enter") {
                  newRowIndex = currentRowIndex + 1;
                  newCellIndex = currentCellIndex;
                }

                if (newRowIndex !== -1 && newCellIndex !== -1 && tableExcel.getRows()[newRowIndex]) {
                  const newTableRowToFocus = tableExcel.getRows()[newRowIndex];
                  if (newTableRowToFocus) {
                    const newTableCellToFocus = newTableRowToFocus.getCells()[newCellIndex];
                    if (newTableCellToFocus) {
                      let newInputToFocusHtml = $(`#${newTableCellToFocus.sId}`);
                      let newInputToFocus = sap.ui.getCore().byId(newTableCellToFocus.sId);
                      if (newInputToFocus && newInputToFocus.getEnabled()) {
                        // reset old selections
                        $(".customFocusExcelStyle").removeClass("customFocusExcelStyle");
                        newInputToFocusHtml.addClass("customFocusExcelStyle");
                        newInputToFocus.focus();
                      }
                    }
                  }
                }
              } else {
                return;
              }
            }
          }
        }
      };

      let startInputSel = null;

      // onmousedown event for excel like cells selection
      document.getElementById(tableExcel.sId).onmousedown = function (e) {
        startInputSel = null;

        // reset old selections
        $(".customFocusExcelStyle").removeClass("customFocusExcelStyle");

        // check if the element is the input
        if (e.target) {
          if (e.target.classList) {
            if (e.target.classList.contains("sapMInputBaseInner")) {
              startInputSel = e.target;
            }
          }
        }
      };

      // onmousedown event for excel like cells selection
      document.getElementById(tableExcel.sId).onmouseup = function (e) {
        // check if the element is the input
        // when the mouse is dragged over the cells
        // this code will select all cells (for bulk edit or other functions)
        if (e.target) {
          if (e.target.classList) {
            if (e.target.classList.contains("sapMInputBaseInner") && startInputSel) {
              let fromRowIndex = startInputSel.closest(".sapUiTableRow") ? startInputSel.closest(".sapUiTableRow").rowIndex - 1 : null;
              let toRowIndex = e.target.closest(".sapUiTableRow") ? e.target.closest(".sapUiTableRow").rowIndex - 1 : null;
              let fromCellIndex = startInputSel.closest(".sapUiTableCell") ? startInputSel.closest(".sapUiTableCell").cellIndex : null;
              let toCellIndex = e.target.closest(".sapUiTableCell") ? e.target.closest(".sapUiTableCell").cellIndex : null;
              let newTableRowToFocus = null;
              let newTableCellToFocus = null;

              if (fromRowIndex !== -1 && toRowIndex !== -1 && fromCellIndex !== -1 && toCellIndex !== -1) {
                let startRowIndex = fromRowIndex > toRowIndex ? toRowIndex : fromRowIndex;
                let endRowIndex = toRowIndex > fromRowIndex ? toRowIndex : fromRowIndex;
                let startCellIndex = fromCellIndex > toCellIndex ? toCellIndex : fromCellIndex;
                let endCellIndex = toCellIndex > fromCellIndex ? toCellIndex : fromCellIndex;
                let newInputToFocus = null;
                let counter = 0;

                for (var a = startRowIndex; a <= endRowIndex; a++) {
                  counter++;
                  // rows
                  newTableRowToFocus = tableExcel.getRows()[a];
                  if (newTableRowToFocus) {
                    for (var b = startCellIndex; b <= endCellIndex; b++) {
                      // cells
                      newTableCellToFocus = newTableRowToFocus.getCells()[b];
                      if (newTableCellToFocus) {
                        let newInputToFocusHtml = $(`#${newTableCellToFocus.sId}`);
                        newInputToFocus = sap.ui.getCore().byId(newTableCellToFocus.sId);
                        if (newInputToFocusHtml) {
                          newInputToFocusHtml.addClass("customFocusExcelStyle");
                        }
                      }
                      if (b === endCellIndex && a === endRowIndex && newInputToFocus && counter > 1) {
                        // on last cell of last row selected open the popover
                        popoverBulkEdit.openBy(newInputToFocus);
                      }
                    }
                  }
                }
              }
            }
          }
        }
      };
    },
  });
}
