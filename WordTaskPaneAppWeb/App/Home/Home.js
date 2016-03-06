/// <reference path="../App.js" />

(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();

            $('#get-data-from-selection').click(getDataFromSelection);
            $('#bindContentControl').click(bindContentControl);
            $('#getAllBinding').click(getAllBinding);
            
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

    function bindContentControl() {
        Office.context.document.bindings.addFromNamedItemAsync('Field',
       Office.BindingType.Text, { id: 'firstName1' },
       function (result) {
           if (result.status === Office.AsyncResultStatus.Succeeded) {
               write('Control bound. Binding.id: '
                   + result.value.id + ' Binding.type: ' + result.value.type);
           } else {
               write('Error:', result.error.message);
           }
       });
    }

    function getAllBinding() {
        Office.context.document.bindings.getAllAsync(function (asyncResult) {
            var bindingString = '';
            for (var i in asyncResult.value) {
                bindingString += asyncResult.value[i].id + '\n';
            }
            write('Existing bindings: ' + bindingString);
        });

        Office.context.document.bindings.getByIdAsync('firstName', function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                write('Action failed. Error: ' + asyncResult.error.message);
            }
            else {
                write('Retrieved binding with type: ' + asyncResult.value.type + ' and id: ' + asyncResult.value.id);
            }
        });

        Office.select("bindings#firstName", function onError() { }).getDataAsync(function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                write('Action failed. Error: ' + asyncResult.error.message);
            } else {
                write(asyncResult.value);
            }
        });

        Office.select("bindings#firstName").getDataAsync( function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                write('Action failed. Error: ' + asyncResult.error.message);
            } else {
                write(asyncResult.value);
            }
        });

       
    }

  
    function write(message) {
        app.showNotification("", message)
        document.getElementById('message').innerText += message;
    }
})();