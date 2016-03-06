/// <reference path="../App.js" />

(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();
            
            $('#get-data-from-selection').click(getDataFromSelection);
        });
    };

    // Reads data from current document selection and displays a notification
    function getDataFromSelection() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    app.showNotification('The selected text is:', '"' + result.value + '"');
                } else {
                    app.showNotification('Error:', result.error.message);
                }
            }
        );
    }

    function showMessage(msg) {
        app.showNotification(msg);
    }
    function SetSeletedText() {
        Office.context.document.setSelectedDataAsync("Hello World!",
                                                  function (asyncResult) {
                                                      if (asyncResult.status == "failed") {
                                                          showMessage("Action failed with error: " + asyncResult.error.message);
                                                      } else {
                                                          showMessage("Success! Click the Next button to move on.");
                                                      }
                                                  });
    }

    function ReadSelectedText() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
                                             function (asyncResult) {
                                                 if (asyncResult.status == "failed") {
                                                     showMessage("Action failed with error: " + asyncResult.error.message);
                                                 }
                                                 else {
                                                     showMessage("Selected data: " + asyncResult.value +
                                                     " Click the Next button to choose a new tutorial.");
                                                 }
                                             });
    }


    function WriteRange() {
        var myMatrix = [["1", "2", "3"], ["4", "5", "6"], ["7", "8", "9"]];

        // Set myMatrix in the document.
        Office.context.document.setSelectedDataAsync(myMatrix, { coercionType: Office.CoercionType.Matrix },
                                                     function (asyncResult) {
                                                         if (asyncResult.status == "failed") {
                                                             showMessage("Action failed with error: " + asyncResult.error.message);
                                                         } else {
                                                             showMessage("You successfully wrote a matrix! Click Next to learn how to read one.");
                                                         }
                                                     });
    }


    function ReadRange() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Matrix,
                                             function (asyncResult) {
                                                 if (asyncResult.status == "failed") {
                                                     showMessage("Action failed with error: " + asyncResult.error.message);
                                                 }
                                                 else {
                                                     showMessage("Selected data: " + asyncResult.value);
                                                 }
                                             });
    }

    function WriteTable() {
        var myTable = new Office.TableData();
        myTable.headers = ["First Name", "Last Name", "Grade"];
        myTable.rows = [["Brittney", "Booker", "A"], ["Sanjit", "Pandit", "C"],
                       ["Naomi", "Peacock", "B"]];

        // Set the myTable in the document.
        Office.context.document.setSelectedDataAsync(myTable, { coercionType: Office.CoercionType.Table },
                                                     function (asyncResult) {
                                                         if (asyncResult.status == "failed") {
                                                             showMessage("Action failed with error: " + asyncResult.error.message);
                                                         } else {
                                                             showMessage("Check out your new table, then click next to learn another API call.");
                                                         }
                                                     });

    }
    function readTable() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Table,
                                             function (asyncResult) {
                                                 if (asyncResult.status == "failed") {
                                                     showMessage("Action failed with error: " + asyncResult.error.message);
                                                 }
                                                 else {
                                                     showMessage("Headers: " + asyncResult.value.headers + " Rows: " +
                                                     asyncResult.value.rows);
                                                 }
                                             });
    }

    function BindToTable() {
        //Bind to the table in the document from user current selection
        Office.context.document.bindings.addFromSelectionAsync(Office.BindingType.Table,
                                                              { id: "MyTableBinding" },
                                                              function (asyncResult) {
                                                                  if (asyncResult.status == "failed") {
                                                                      showMessage("Action failed with error: " + asyncResult.error.message);
                                                                  }
                                                                  else {
                                                                      showMessage("Added new binding with type: " + asyncResult.value.type +
                                                                " and id: " + asyncResult.value.id);
                                                                  }
                                                              });

    }
    function AddRow() {

        //Creating the row to update 
        var table = new Office.TableData();
        table.rows = ["Seattle", "WA"];
        var rowToUpdate = 2;

        //Getting the table binding and setting data in 3rd row.
        Office.select("bindings#MyTableBinding", onBindingNotFound).setDataAsync(table,
                                                     {
                                                         coercionType: Office.CoercionType.Table,
                                                         startRow: rowToUpdate
                                                     },
                                                     function (asyncResult) {
                                                         if (asyncResult.status == "failed") {
                                                             showMessage("Action failed with error: " + asyncResult.error.message);
                                                         }
                                                         else {
                                                             showMessage("Updated row number 3 with this data: " + table.rows);
                                                         }
                                                     });

        //Show error message in case the binding object wasn't found
        function onBindingNotFound() {
            showMessage("The binding object was not found. " +
          "Please return to previous step to create the binding");
        }
    }

    function AddColumnTable() {
        //Create a table with a single column
        var populationTable = new Office.TableData();
        populationTable.headers = [["Population"]];
        populationTable.rows = [["1593659"], ["416468"], ["616627"], ["645169"]];

        //Finding the binding by its id
        Office.context.document.bindings.getByIdAsync("MyTableBinding",
                                                      function (asyncResult) {
                                                          if (asyncResult.status == "failed") {
                                                              showMessage("Action failed with error: " + asyncResult.error.message);
                                                          }
                                                          else {
                                                              //getByIdAsync returns a binding object. 
                                                              //If the binding object was found, add a column to it
                                                              asyncResult.value.addColumnsAsync(populationTable, function (result) {
                                                                  if (result.status == "failed") {
                                                                      showMessage("Action failed with error: " + result.error.message);
                                                                  }
                                                                  else {
                                                                      showMessage("Successfully added Population column!");
                                                                  }
                                                              });
                                                          }
                                                      });
    }

    function GetSelectedCordinate() {
        /* Click Run Code to add an event handler to the Matrix binding. 
Then select different cells in the Matrix to trigger the event 
and read the current selected cell. */

        //Get the binding and add an event handler to detect selection change events
        Office.select("bindings#MyMatrixBinding", onBindingNotFound).
          addHandlerAsync(Office.EventType.BindingSelectionChanged,
                          onBindingSelectionChanged,
                          function (AsyncResult) {
                              showMessage("Event handler was added successfully!" +
                          " Change the matrix current selection to trigger the event");
                          });

        //Trigger on selection change, get partial data from the matrix
        function onBindingSelectionChanged(eventArgs) {
            eventArgs.binding.getDataAsync(
              {
                  CoercionType: Office.CoercionType.Matrix,
                  startRow: eventArgs.startRow,
                  startColumn: eventArgs.startColumn,
                  rowCount: 1, columnCount: 1
              },
              function (asyncResult) {
                  if (asyncResult.status == "failed") {
                      showMessage("Action failed with error: " + asyncResult.error.message);
                  }
                  else {
                      showMessage(asyncResult.value[0].toString());
                  }
              });
        }

        //Show error message in case the binding object wasn't found
        function onBindingNotFound() {
            showMessage("The binding object was not found." +
          " Please return to previous step to create the binding");
        }
    }

    function persistSetting() {
        // Set a setting in the document
        Office.context.document.settings.set("mySetting", "mySetting value");
        showMessage("You have saved a new setting. Click next to retrieve it.");
        //Get a setting previously set in the document
        var settingsValue = Office.context.document.settings.get("mySetting");
        showMessage("mySetting value is: " + settingsValue);

        //Save a setting in the document to make it available in future sessions
        Office.context.document.settings.saveAsync(function (asyncResult) {
            if (asyncResult.status == "failed") {
                showMessage("Action failed with error: " + asyncResult.error.message);
            }
            else {
                showMessage("Settings saved with status: " + asyncResult.status);
            }
        });

    }

    function SelectionChanged() {
        //For more practice, try changing this function to use Office.context.document.setSelectedDataAsync creatively!
        Globals.documentSelectionHandler = function (args) {
            Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, function (asyncResult) {
                if (asyncResult.status == "succeeded") {
                    showMessage("DocumentSelectionChanged: " + asyncResult.value);
                }
            });
        }

        Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, Globals.documentSelectionHandler, function (asyncResult) {
            if (asyncResult.status == "failed") {
                showMessage("Action failed with error: " + asyncResult.error.message);
            }
            else {
                showMessage("DocumentSelectionChanged handler added successfully." +
                  " Click Next to learn how to remove it.");
            }
        });

        /* Click Run Code to unregister the active view changed event */

        Office.context.document.removeHandlerAsync(Office.EventType.DocumentSelectionChanged, { handler: Globals.documentSelectionHandler }, function (asyncResult) {
            if (asyncResult.status == "failed") {
                showMessage("Action failed with error: " + asyncResult.error.message);
            }
            else {
                showMessage("DocumentSelectionChanged handler remove succeeded");
            }
        });
    }


    function navigaeToBinding() {
        // Create a TableData object.
        var myTable = new Office.TableData();
        myTable.headers = ["First Name", "Last Name", "Balance"];
        myTable.rows = [["Brittney", "Booker", "1223.10"], ["Sanjit", "Pandit", "34234.99"],
                        ["Naomi", "Peacock", "-50.78"]];

        // Set the myTable in the document.
        Office.context.document.setSelectedDataAsync(myTable, { coercionType: Office.CoercionType.Table },
            function (asyncResult) {
                if (asyncResult.status == "failed") {
                    showMessage("Action failed with error: " + asyncResult.error.message);
                }
            });

        //Create a new table binding for the selected table.
        Office.context.document.bindings.addFromSelectionAsync(Office.CoercionType.Table, { id: "MyTableBinding" }, function (asyncResult) {
            if (asyncResult.status == "failed") {
                showMessage("Action failed with error: " + asyncResult.error.message);
            } else {
                showMessage("Added new binding with type: " + asyncResult.value.type + " and id: " + asyncResult.value.id +
                ". Click next to learn how to navigate to this new binding.");
            }
        });

        //Go to binding by ID. Scroll so the binding is off-screen, then click Run Code
        Office.context.document.goToByIdAsync("MyTableBinding", Office.GoToType.Binding, function (asyncResult) {
            if (asyncResult.status == "failed") {
                showMessage("Action failed with error: " + asyncResult.error.message);
            }
            else {
                showMessage("Navigation successful!");
            }
        });
    }

    function getFileLoaction() {
        // Make the async call to get the file properties
        //Note: This will return undefined when the document is embedded in a webpage.
        Office.context.document.getFilePropertiesAsync(
            function (asyncResult) {
                if (asyncResult.status == "failed") {
                    showMessage("Action failed with error: " + asyncResult.error.message);
                } else {
                    showMessage("The document location is: " + asyncResult.value.url);
                }
            });
    }

    function setFormatting() {
        // Create a TableData object.
        var myTable = new Office.TableData();
        myTable.headers = ["First Name", "Last Name", "Balance"];
        myTable.rows = [["Brittney", "Booker", "1223.10"], ["Sanjit", "Pandit", "34234.99"],
                       ["Naomi", "Peacock", "-50.78"]];

        // Set the myTable in the document.
        Office.context.document.setSelectedDataAsync(myTable, {
            coercionType: Office.CoercionType.Table,
            cellFormat: [
                //Set the font color to yellow in the header
                { cells: Office.Table.Headers, format: { fontColor: "yellow" } },
                //Set the data cells to gray background with blue font color
                { cells: Office.Table.Data, format: { fontColor: "blue", backgroundColor: "gray" } },
                //Set the number format for Currency
                { cells: { column: 2 }, format: { numberFormat: "$#,##0.00_);[Red]($#,##0.00)" } },
                //Auto fit the column widths
                { cells: Office.Table.All, format: { width: "auto fit" } }
            ]
        },
            function (asyncResult) {
                if (asyncResult.status == "failed") {
                    showMessage("Action failed with error: " + asyncResult.error.message);
                } else {
                    showMessage("Check out your fancy new table!");
                }
            });
    }

})();
