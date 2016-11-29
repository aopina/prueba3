/// <reference path="../App.js" />
// global app
var ord;

(function () {
    'use strict';

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();

            $('#get-data-from-selection').click(getDataFromSelection);
        });
    };

	function obtener_orden(orden){
		ord = orden;
		
	}
	function regresar_orden(){
		return ord;
	}

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
})();