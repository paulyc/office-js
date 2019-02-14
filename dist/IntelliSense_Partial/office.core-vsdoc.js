
var OfficeCore;
(function (OfficeCore) {
	var RequestContext = (function (_super) {
		__extends(RequestContext, _super);
		function RequestContext() {
			/// <summary>
			/// The RequestContext object facilitates requests to the OfficeCore application. Since the Office add-in and the OfficeCore application run in two different processes, the request context is required to get access to the OfficeCore object model from the add-in.
			/// </summary>
			_super.call(this, null);
		}
		return RequestContext;
	})(OfficeExtension.ClientRequestContext);
	OfficeCore.RequestContext = RequestContext;

	OfficeCore.run = function (batch) {
		/// <signature>
		/// <summary>
		/// Executes a batch script that performs actions on the OfficeCore object model, using a new RequestContext. When the promise is resolved, any tracked objects that were automatically allocated during execution will be released.
		/// </summary>
		/// <param name="batch" type="function(context) { ... }">
		/// A function that takes in a RequestContext and returns a promise (typically, just the result of "context.sync()").
		/// <br />
		/// The context parameter facilitates requests to the OfficeCore application. Since the Office add-in and the OfficeCore application run in two different processes, the RequestContext is required to get access to the OfficeCore object model from the add-in.
		/// </param>
		/// </signature>
		/// <signature>
		/// <summary>
		/// Executes a batch script that performs actions on the OfficeCore object model, using the RequestContext of a previously-created API object. When the promise is resolved, any tracked objects that were automatically allocated during execution will be released.
		/// </summary>
		/// <param name="object" type="OfficeExtension.ClientObject">
		/// A previously-created API object. The batch will use the same RequestContext as the passed-in object, which means that any changes applied to the object will be picked up by "context.sync()".
		/// </param>
		/// <param name="batch" type="function(context) { ... }">
		/// A function that takes in a RequestContext and returns a promise (typically, just the result of "context.sync()").
		/// <br />
		/// The context parameter facilitates requests to the OfficeCore application. Since the Office add-in and the OfficeCore application run in two different processes, the RequestContext is required to get access to the OfficeCore object model from the add-in.
		/// </param>
		/// </signature>
		/// <signature>
		/// <summary>
		/// Executes a batch script that performs actions on the OfficeCore object model, using the RequestContext of a previously-created API object. When the promise is resolved, any tracked objects that were automatically allocated during execution will be released.
		/// </summary>
		/// <param name="objects" type="Array&lt;OfficeExtension.ClientObject&gt;">
		/// An array of previously-created API objects. The array will be validated to make sure that all of the objects share the same context. The batch will use this shared RequestContext, which means that any changes applied to these objects will be picked up by "context.sync()".
		/// </param>
		/// <param name="batch" type="function(context) { ... }">
		/// A function that takes in a RequestContext and returns a promise (typically, just the result of "context.sync()").
		/// <br />
		/// The context parameter facilitates requests to the OfficeCore application. Since the Office add-in and the OfficeCore application run in two different processes, the RequestContext is required to get access to the OfficeCore object model from the add-in.
		/// </param>
		/// </signature>
		arguments[arguments.length - 1](new OfficeCore.RequestContext());
		return new OfficeExtension.Promise();
	}
})(OfficeCore || (OfficeCore = {__proto__: null}));
OfficeCore.__proto__ = null;

