/* 
Copyright 2008 Mark Logic Corporation

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
*/
/* $Id: debug.js,v 1.6 2008-12-29 20:08:57 paven Exp $ */

// Last resort error handler
Event.observe(window, "error", function(ex) {
	alert(ex);
});


// Global flag to enable debugging
var _debug = false;

// Global log function. By default it does nothing.
var _l = function() {}

if(_debug) {

	if(!window.console) {
		window.console = (function() {
			var log = document.createElement("textarea");
			log.id = "ML-Debug-Log";
			log.readOnly = "readOnly";
			return {
				init: function() {
					document.body.appendChild(log);
					log = $(log);
					log.setStyle({
						"position": "fixed",
						"bottom": "2em",
						"left": "0",
						"width": "20em",
						"height": "10em",
					    "z-index": "10"
					});
				},
				log: function(message) {
					log.value = message + "\n" + log.value; 
				}
			}
		})();
	}
	_l = window.console.log;


	document.observe("dom:loaded", function() {
		window.console.init();
		// Create a "Refresh" button
		var refresh = document.createElement("button");
		refresh.id = "ML-Debug-Refresh";
		refresh.appendChild(document.createTextNode("Refresh"));
		document.body.appendChild(refresh);
		refresh = $(refresh); 
		refresh.setStyle({
			"position": "fixed",
			"bottom": "2em",
			"right": "0",
			"width": "8em",
		    "z-index": "10"
		});
		refresh.observe("click", function(){ 
				window.location.reload();
		});
	});
}
