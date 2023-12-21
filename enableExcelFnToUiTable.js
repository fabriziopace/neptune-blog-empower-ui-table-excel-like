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
      // bind rows and set model data
      tableExcel.bindRows("/");
      modeltableExcel.setData(arrayTableData);

      const tableExcelObj = document.getElementById(tableExcel.sId);

      var startInputSel = null;
      var isMouseUp = true;
      var lastInputSel = null;

      if (tableExcelObj) {
        // onkeydown event for arrow keys / enter navigation
        tableExcelObj.onkeydown = function (e) {
          // close popover
          popoverBulkEdit.close();

          // on key down check if the focused element is a input
          const currentTarget = e.target;

          // get table rows
          const tableRows = tableExcel.getRows();

          if (currentTarget) {
            if (currentTarget.classList) {
              if (currentTarget.classList.contains("sapMInputBaseInner")) {
                if (e.key === "ArrowLeft" || e.key === "ArrowUp" || e.key === "ArrowRight" || e.key === "ArrowDown" || e.key === "Enter") {
                  // The closest() method of the Element interface traverses the element and its parents (heading toward the document root)
                  // until it finds a node that matches the specified CSS selector.
                  // Ref. https://developer.mozilla.org/en-US/docs/Web/API/Element/closest

                  // navigation arrows between inputs
                  const targetClosestTableCell = currentTarget.closest(".sapUiTableCell");

                  let currentRowIndex = 0;
                  let currentColIndex = 0;

                  if (targetClosestTableCell.id) {
                    // based on the cell id get current row and column indexes
                    // example table-rows-row1-col1-fixed
                    $.each(targetClosestTableCell.id.split("-"), (i, idEl) => {
                      if (idEl.toString().substr(0, 4) === "rows") {
                        return true;
                      }
                      if (idEl.toString().substr(0, 3) === "row") {
                        currentRowIndex = idEl.replace("row", "");
                      }
                      if (idEl.toString().substr(0, 3) === "col") {
                        currentColIndex = idEl.replace("col", "");
                      }
                    });

                    currentRowIndex = parseInt(currentRowIndex) ? parseInt(currentRowIndex) : 0;
                    currentColIndex = parseInt(currentColIndex) ? parseInt(currentColIndex) : 0;
                  }

                  let newRowIndex = 0;
                  let newColIndex = 0;

                  // calculate the new index based on the current row / cell index selected
                  // key arrow left
                  if (e.key === "ArrowLeft") {
                    newRowIndex = currentRowIndex;
                    newColIndex = currentColIndex - 1;
                  }

                  // key arrow up
                  if (e.key === "ArrowUp") {
                    newRowIndex = currentRowIndex > 0 ? currentRowIndex - 1 : currentRowIndex;
                    newColIndex = currentColIndex;
                  }

                  // key arrow right
                  if (e.key === "ArrowRight") {
                    newRowIndex = currentRowIndex;
                    newColIndex = currentColIndex < tableColumns.length ? currentColIndex + 1 : currentColIndex;
                  }

                  // key arrow down / key enter
                  if (e.key === "ArrowDown" || e.key === "Enter") {
                    newRowIndex = currentRowIndex + 1;
                    newColIndex = currentColIndex;
                  }

                  if (newRowIndex !== -1 && newColIndex !== -1 && tableRows[newRowIndex]) {
                    const newTableRowToFocus = tableRows[newRowIndex];
                    if (newTableRowToFocus) {
                      const newTableCellToFocus = newTableRowToFocus.getCells()[newColIndex];
                      if (newTableCellToFocus) {
                        let newInputToFocusHtml = $(`#${newTableCellToFocus.sId}`);
                        let newInputToFocus = sap.ui.getCore().byId(newTableCellToFocus.sId);
                        if (newInputToFocus && newInputToFocus.getEnabled()) {
                          // reset old selections
                          $(".customFocusExcelStyle").removeClass("customFocusExcelStyle");

                          // focus new input
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

        // onmousedown event for excel-like cells selection
        tableExcelObj.onmousedown = function (e) {
          if (e.button !== 0) {
            // if the mouse button pressed is not the left exit
            return;
          }

          // called when the drag start
          const currentTarget = e.target;
          startInputSel = null;
          isMouseUp = false;

          // reset old selections if the ctrl button is not pressed
          if (!e.ctrlKey) {
            $(".customFocusExcelStyle").removeClass("customFocusExcelStyle");
          }

          // check if the element is a input
          if (currentTarget) {
            if (currentTarget.classList) {
              if (currentTarget.classList.contains("sapMInputBaseInner")) {
                startInputSel = currentTarget;

                // select first cell
                let newTableCellToFocus = startInputSel.closest(".sapUiTableCell");
                if (newTableCellToFocus) {
                  let newInputToFocusHtml = $(`#${newTableCellToFocus.id}`);
                  if (newInputToFocusHtml) {
                    newInputToFocusHtml.addClass("customFocusExcelStyle");

                    // if ctrl key is not pressed
                    // and we have more than 1 cell selected
                    // show the popover for bulk edit
                    let numInputFocused = document.getElementsByClassName("customFocusExcelStyle");
                    let htmlInputFocused = startInputSel.closest(".sapMInputBase");
                    if (htmlInputFocused) {
                      htmlInputFocused.classList.add("customFocusExcelStyle");
                    }
                    if (e.ctrlKey && numInputFocused.length > 1) {
                      setTimeout(function () {
                        let lastInputFocused = sap.ui.getCore().byId(htmlInputFocused.id);
                        popoverBulkEdit.openBy(lastInputFocused);
                      }, 100);
                    }
                  }
                }
              }
            }
          }
        };

        // onmousemove event for excel-like cells selection
        tableExcelObj.onmousemove = function (e) {
          // when the mouse is dragged over the cells
          // this code will select all cells (for bulk edit etc)
          const currentTarget = e.target;

          if (isMouseUp) {
            return;
          }

          // get table rows
          const tableRows = tableExcel.getRows();

          if (currentTarget) {
            if (currentTarget.classList) {
              if (currentTarget.classList.contains("sapMInputBaseInner") && startInputSel) {
                const startTargetClosestTableCell = startInputSel.closest(".sapUiTableCell");
                const endTargetClosestTableCell = currentTarget.closest(".sapUiTableCell");
                let fromRowIndex = 0;
                let toRowIndex = 0;
                let fromColIndex = 0;
                let toColIndex = 0;

                let newTableRowToFocus = null;
                if (startTargetClosestTableCell.id) {
                  // based on the cell id get current row and column indexes
                  // example table-rows-row1-col1-fixed
                  $.each(startTargetClosestTableCell.id.split("-"), (i, idEl) => {
                    if (idEl.toString().substr(0, 4) === "rows") {
                      return true;
                    }
                    if (idEl.toString().substr(0, 3) === "row") {
                      fromRowIndex = idEl.replace("row", "");
                    }
                    if (idEl.toString().substr(0, 3) === "col") {
                      fromColIndex = idEl.replace("col", "");
                    }
                  });

                  fromRowIndex = parseInt(fromRowIndex) ? parseInt(fromRowIndex) : 0;
                  fromColIndex = parseInt(fromColIndex) ? parseInt(fromColIndex) : 0;
                }
                if (endTargetClosestTableCell.id) {
                  // based on the cell id get current row and column indexes
                  // example table-rows-row1-col1-fixed
                  $.each(endTargetClosestTableCell.id.split("-"), (i, idEl) => {
                    if (idEl.toString().substr(0, 4) === "rows") {
                      return true;
                    }
                    if (idEl.toString().substr(0, 3) === "row") {
                      toRowIndex = idEl.replace("row", "");
                    }
                    if (idEl.toString().substr(0, 3) === "col") {
                      toColIndex = idEl.replace("col", "");
                    }
                  });

                  toRowIndex = parseInt(toRowIndex) ? parseInt(toRowIndex) : 0;
                  toColIndex = parseInt(toColIndex) ? parseInt(toColIndex) : 0;
                }

                if (fromRowIndex !== -1 && toRowIndex !== -1 && fromColIndex !== -1 && toColIndex !== -1) {
                  let startRowIndex = fromRowIndex > toRowIndex ? toRowIndex : fromRowIndex;
                  let endRowIndex = toRowIndex > fromRowIndex ? toRowIndex : fromRowIndex;

                  let startColIndex = fromColIndex > toColIndex ? toColIndex : fromColIndex;
                  let endColIndex = toColIndex > fromColIndex ? toColIndex : fromColIndex;

                  let newInputToFocus = null;
                  let counterCells = 0;

                  // loop each cell in each row for the massive selection
                  for (var a = startRowIndex; a <= endRowIndex; a++) {
                    // rows
                    newTableRowToFocus = tableRows[a];
                    if (newTableRowToFocus) {
                      for (var b = startColIndex; b <= endColIndex; b++) {
                        counterCells++;

                        // columns
                        let newTableCellToFocus = newTableRowToFocus.getCells()[b];
                        if (newTableCellToFocus) {
                          let newInputToFocusHtml = $(`#${newTableCellToFocus.sId}`);
                          newInputToFocus = sap.ui.getCore().byId(newTableCellToFocus.sId);

                          if (newInputToFocusHtml) {
                            newInputToFocusHtml.addClass("customFocusExcelStyle");
                          }
                        }
                        if (b === endColIndex && a === endRowIndex && newInputToFocus && counterCells > 1) {
                          // store the last cell of last row selected to a variable
                          lastInputSel = newInputToFocus;
                        }
                      }
                    }
                  }
                }
              }
            }
          }
        };

        // onmouseup event for excel-like cells selection
        document.onmouseup = function (e) {
          // called when the drag ends
          // reset variable
          isMouseUp = true;

          if (lastInputSel) {
            // on last cell of last row selected open the popover
            popoverBulkEdit.openBy(lastInputSel);

            // reset variable
            lastInputSel = null;
          }
        };
      }
    },
  });
}
