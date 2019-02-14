
var __extends = this.__extends || function (d, b) {
	for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p];
	function __() { this.constructor = d; }
	__.prototype = b.prototype;
	d.prototype = new __();
};

var OfficeExtension;
(function (OfficeExtension) {
	var ClientObject = (function () {
		function ClientObject() {
			/// <field name="isNullObject" type="Boolean">Returns a boolean value for whether the corresponding object is a null object. You must call "context.sync()" before reading the isNullObject property.</field>
		}
		return ClientObject;
	})();
	OfficeExtension.ClientObject = ClientObject;
})(OfficeExtension || (OfficeExtension = {__proto__: null}));

var OfficeExtension;
(function (OfficeExtension) {
	var ClientRequestContext = (function () {
		function ClientRequestContext(url) {
			/// <summary>
			/// An abstract RequestContext object that facilitates requests to the host Office application. The "Excel.run" and "Word.run" methods provide a request context.
			/// </summary>
			/// <field name="trackedObjects" type="OfficeExtension.TrackedObjects"> Collection of objects that are tracked for automatic adjustments based on surrounding changes in the document. </field>
			/// <field name="requestHeaders" type="Object">Request headers.</field>
			this.requestHeaders = {
				__proto__: null,
			};
		}
		ClientRequestContext.prototype.load = function (object, option) {
			/// <summary>
			/// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
			/// </summary>
			/// <param name="object" type="OfficeExtension.ClientObject" />
			/// <param name="option" type="string|string[]|{select?, expand?, top?, skip?}" />
		};
		ClientRequestContext.prototype.loadRecursive = function (object, options, maxDepth) {
			/// <summary>
			/// Queues up a command to recursively load the specified properties of the object and its navigation properties. You must call "context.sync()" before reading the properties.
			/// </summary>
			/// <param name="object" type="OfficeExtension.ClientObject">The object to be loaded.</param>
			/// <param name="option" type="string|string[]|{select?, expand?, top?, skip?}">
			///     The key-value pairing of load options for the types, such as { "Workbook": "worksheets,tables",  "Worksheet": "tables",  "Tables": "name" }
			/// </param>
			/// <param name="maxDepth" type="Number" optional="true">The maximum recursive depth.</param>
		};
		ClientRequestContext.prototype.trace = function (message) {
			/// <summary>
			/// Adds a trace message to the queue. If the promise returned by "context.sync()" is rejected due to an error, this adds a ".traceMessages" array to the OfficeExtension.Error object, containing all trace messages that were executed. These messages can help you monitor the program execution sequence and detect the cause of the error.
			/// </summary>
			/// <param name="message" type="String" />
		};
		ClientRequestContext.prototype.sync = function (passThroughValue) {
			/// <summary>
			/// Synchronizes the state between JavaScript proxy objects and the Office document, by executing instructions queued on the request context and retrieving properties of loaded Office objects for use in your code. This method returns a promise, which is resolved when the synchronization is complete.
			/// </summary>
			/// <param name="passThroughValue" optional="true" />
			return new OfficeExtension.Promise();
		};
		ClientRequestContext.prototype.__proto__ = null;
		return ClientRequestContext;
	})();
	OfficeExtension.ClientRequestContext = ClientRequestContext;
})(OfficeExtension || (OfficeExtension = {__proto__: null}));

var OfficeExtension;
(function (OfficeExtension) {
	var ClientResult = (function () {
		function ClientResult() {
			/// <summary>
			/// Contains the result for methods that return primitive types. The object's value property is retrieved from the document after "context.sync()" is invoked.
			/// </summary>
			/// <field name="value">
			/// The value of the result that is retrieved from the document after "context.sync()" is invoked.
			/// </field>
		}
		ClientResult.prototype.__proto__ = null;
		return ClientResult;
	})();
	OfficeExtension.ClientResult = ClientResult;
})(OfficeExtension || (OfficeExtension = {__proto__: null}));

var OfficeExtension;
(function (OfficeExtension) {
	var DebugInfo = (function () {
		function DebugInfo() {
			/// <summary>
			/// Debug info (useful for detailed logging of the error, i.e., via JSON.stringify(...)).
			/// </summary>
			/// <field name="code" type="String">
			/// Error code string, such as "InvalidArgument".
			/// </field>
			/// <field name="message" type="String">
			/// The error message passed through from the host Office application.
			/// </field>
			/// <field name="innerError" type="DebugInfo|String">
			/// Inner error, if applicable.
			/// </field>

			/// <field name="errorLocation" type="String">
			/// The object type and property or method name (or similar information), if available.
			/// </field>
		}
		DebugInfo.prototype.__proto__ = null;
		return DebugInfo;
	})();
	OfficeExtension.DebugInfo = DebugInfo;

	var Error = (function () {
		function Error() {
			/// <summary>
			/// The error object returned by "context.sync()", if a promise is rejected due to an error while processing the request.
			/// </summary>
			/// <field name="name" type="String">
			/// Error name: "OfficeExtension.Error"
			/// </field>
			/// <field name="message" type="String">
			/// The error message passed through from the host Office application.
			/// </field>
			/// <field name="stack" type="String">
			/// Stack trace, if applicable.
			/// </field>
			/// <field name="code" type="String">
			/// Error code string, such as "InvalidArgument".
			/// </field>
			/// <field name="traceMessages" type="Array" elementType="string">
			/// Trace messages (if any) that were added via a "context.trace()" invocation before calling "context.sync()". If there was an error, this contains all trace messages that were executed before the error occurred. These messages can help you monitor the program execution sequence and detect the case of the error.
			/// </field>
			/// <field name="debugInfo" type="OfficeExtension.DebugInfo">
			/// Debug info (useful for detailed logging of the error, i.e., via JSON.stringify(...)).
			/// </field>
			/// <field name="innerError" type="Error">
			/// Inner error, if applicable.
			/// </field>
		}
		Error.prototype.__proto__ = null;
		return Error;
	})();
	OfficeExtension.Error = Error;
})(OfficeExtension || (OfficeExtension = {__proto__: null}));

var OfficeExtension;
(function (OfficeExtension) {
	var ErrorCodes = (function () {
		function ErrorCodes() {
		}
		ErrorCodes.__proto__ = null;
		ErrorCodes.accessDenied = "";
		ErrorCodes.generalException = "";
		ErrorCodes.activityLimitReached = "";
		ErrorCodes.invalidObjectPath = "";
		ErrorCodes.propertyNotLoaded = "";
		ErrorCodes.valueNotLoaded = "";
		ErrorCodes.invalidRequestContext = "";
		ErrorCodes.invalidArgument = "";
		ErrorCodes.runMustReturnPromise = "";
		ErrorCodes.cannotRegisterEvent = "";
		ErrorCodes.apiNotFound = "";
		ErrorCodes.connectionFailure = "";
		return ErrorCodes;
	})();
	OfficeExtension.ErrorCodes = ErrorCodes;
})(OfficeExtension || (OfficeExtension = {__proto__: null}));

var OfficeExtension;
(function (OfficeExtension) {
	var Promise = (function () {
		/// <summary>
		/// Creates a promise that resolves when all of the child promises resolve.
		/// </summary>
		Promise.all = function (promises) { return [new OfficeExtension.Promise()]; };
		/// <summary>
		/// Creates a promise that is resolved.
		/// </summary>
		Promise.resolve = function (value) { return new OfficeExtension.Promise(); };
		/// <summary>
		/// Creates a promise that is rejected.
		/// </summary>
		Promise.reject = function (error) { return new OfficeExtension.Promise(); };
		/// <summary>
		/// A Promise object that represents a deferred interaction with the host Office application. The publically-consumable OfficeExtension.Promise is available starting in ExcelApi 1.2 and WordApi 1.2. Promises can be chained via ".then", and errors can be caught via ".catch". Remember to always use a ".catch" on the outer promise, and to return intermediary promises so as not to break the promise chain. When a "native" Promise implementation is available, OfficeExtension.Promise will switch to use the native Promise instead.
		/// </summary>
		Promise.prototype.then = function (onFulfilled, onRejected) {
			/// <summary>
			/// This method will be called once the previous promise has been resolved.
			/// Both the onFulfilled on onRejected callbacks are optional.
			/// If either or both are omitted, the next onFulfilled/onRejected in the chain will be called called.
			/// Returns a new promise for the value or error that was returned from onFulfilled/onRejected.
			/// </summary>
			/// <param name="onFulfilled" type="Function" optional="true"></param>
			/// <param name="onRejected" type="Function" optional="true"></param>
			/// <returns type="OfficeExtension.Promise"></returns>
			onRejected(new Error());
		}
		Promise.prototype.catch = function (onRejected) {
			/// <summary>
			/// Catches failures or exceptions from actions within the promise, or from an unhandled exception earlier in the call stack.
			/// </summary>
			/// <param name="onRejected" type="Function" optional="true">function to be called if or when the promise rejects.</param>
			/// <returns type="OfficeExtension.Promise"></returns>
			onRejected(new Error());
		}
		Promise.prototype.__proto__ = null;
	})
	OfficeExtension.Promise = Promise;
})(OfficeExtension || (OfficeExtension = {__proto__: null}));

var OfficeExtension;
(function (OfficeExtension) {
	var TrackedObjects = (function () {
		function TrackedObjects() {
			/// <summary>
			/// Collection of tracked objects, contained within a request context. See "context.trackedObjects" for more information.
			/// </summary>
		}
		TrackedObjects.prototype.add = function (object) {
			/// <summary>
			/// Track a new object for automatic adjustment based on surrounding changes in the document. Only some object types require this. If you are using an object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
			/// </summary>
			/// <param name="object" type="OfficeExtension.ClientObject|OfficeExtension.ClientObject[]"></param>
		};
		TrackedObjects.prototype.remove = function (object) {
			/// <summary>
			/// Release the memory associated with an object that was previously added to this collection. Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
			/// </summary>
			/// <param name="object" type="OfficeExtension.ClientObject|OfficeExtension.ClientObject[]"></param>
		};
		TrackedObjects.prototype.__proto__ = null;
		return TrackedObjects;
	})();
	OfficeExtension.TrackedObjects = TrackedObjects;
})(OfficeExtension || (OfficeExtension = {__proto__: null}));

(function (OfficeExtension) {
	var EventHandlers = (function () {
		function EventHandlers() { }
		EventHandlers.prototype.add = function (handler) {
			return new EventHandlerResult(null, null, handler);
		};
		EventHandlers.prototype.remove = function (handler) { };
		EventHandlers.prototype.__proto__ = null;
		return EventHandlers;
	}());
	OfficeExtension.EventHandlers = EventHandlers;

	var EventHandlerResult = (function () {
		function EventHandlerResult() { }
		EventHandlerResult.prototype.remove = function () { };
		EventHandlerResult.prototype.__proto__ = null;
		return EventHandlerResult;
	}());
	OfficeExtension.EventHandlerResult = EventHandlerResult;
})(OfficeExtension || (OfficeExtension = {__proto__: null}));

OfficeExtension.__proto__ = null;
