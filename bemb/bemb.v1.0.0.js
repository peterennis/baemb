/* **************
	v1.0.0, jleach 2015-02

	To be loaded by the VBA project after the original DOM is initialized
	Provides base means for Access/VBA to interact with the JS
	To help alleviate VBA string quote management, only single quotes are used in this script
	
	ERRORS:
		10: jQuery not present in host page
	
*****************/


function browseEmbedBridge() {

}


browseEmbedBridge.prototype = {

	initFramework: function() {
		
		try {
			createEventBridgeElement();
			initJQuery();
		} catch (err) {
			alert(err.message);
		}
		
		
		function initJQuery() {
			// for now, I know the test embed has jquery so I'll just use that
			// after the initial tests are set up I'll revisit to dynamically load, etc
			if (typeof jQuery != 'undefined') {
				bemb_$ = jQuery;
			} else {
				throw new bembException(10, "jQuery not present, current version requires jQuery in the hosted page");
			}
		}
		
		function createEventBridgeElement() {
			// add an element to the DOM that we'll subscribe to for events
			var elem = document.createElement('button');
			elem.setAttribute('id', 'bemb-event-element');
			elem.setAttribute('style', 'display: none;');
			document.body.appendChild(elem);
			this.bridgeElement = elem;
		}
		
	},
	
	/////////////////////////////////////
	//
	//	runCode (callString, postFunc)
	//
	//		arg examples: 
	//			callString: function() { return addValues(1, 5); }
	//			postFunc:	function() { bembData += 6 }
	//
	//		example VBA call (internal version):
	//			.Exec "mvbj.runCode(function() { return addValues(1, 5)}, function() { return mvbjData += 6; });"
	//		
	//		example VBA call (dev interface version):
	//			.RunCode("addValues(1, 5);", "mvbjData += 6;")
	//
	//		example return: 12
	//
	//	Run some function the user provides and stuff the value into bembData
	//	optionally, run a postFunc where they can manipulate that data after
	//  the JS fills it (this is useful for turning JS data into readable VBA data
	//  without having to alter the core function(s) at all)
	//
	/////////////////////////////////////
	runCode: function(callString, postFunc) {
	
		// run the function they provided, pass the result to the global data var
		bembData = callString();
		
		// if they gave us a function to run, do it...
		if (typeof postFunc != 'undefined') {	
			postFunc();
		}
	
	},
	
	// links an event on jquery selected elements... callback to bemb_eventListener
	// func is an optional function passed by the VBA that will modify the mvbjData after the
	// event places its info... useful for coercing complex JS data into readable VB strings
	// (see mvbj.runCode() for a similar example)
	addEventListener: function(elementSelector, eventName, eventID, func) {
		var el = bemb_$(elementSelector);
		el.on(eventName, function(event) {bemb_eventListener(eventID, event, func);});
	},
	
	jQueryNativeVersion: '',
	bridgeElement: null
	
}


var bemb = new browseEmbedBridge();	// what we'll call upon to interact with our VBA project
var bembData;						// for transferring data to VBA
var bembRaisedEventID;				// don't want to group the event ID with the event object to transfer, we'll hold this separate
var bembEvent;						// we'll stuff the stringified event data into bembData, but maintain an actual object here...
var bemb_$;							// our very own jQuery... doesn't seem to like being defined in the prototype...


function bembException(number, message) {
	this.message = message;
	this.number = number;
}


// doesn't seem to like being defined in the prototype...
function bemb_eventListener(eventID, event, func) {
	
	bembRaisedEventID = eventID;
		
	// this JSON gets mangled somehow, and I believe it's by the charting library
	// it's problematic... should see about parameterizing the json conversion so it's only used when needed
	// or maybe some error handling...
	//bembData = JSON.stringify(event);
	
	bembEvent = event;
	bembData = event;	// store the event object by default, which they can then use the optional func to translate
	
	// if they gave us a function to run, do it...
	if (typeof func != 'undefined') {	
		func();
	}

	// raise an event on the event bridge so it's received by the subscribing VBA project
	// if the VBA listening method doesn't provide error handling, errors will be propped here
	try {
		this.bridgeElement.onclick();
	} catch (err) {
		alert(err.message);
	}
	
}
