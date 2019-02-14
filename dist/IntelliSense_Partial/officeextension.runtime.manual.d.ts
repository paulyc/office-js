declare namespace OfficeExtension {
	/** An abstract proxy object that represents an object in an Office document. You create proxy objects from the context (or from other proxy objects), add commands to a queue to act on the object, and then synchronize the proxy object state with the document by calling "context.sync()". */
	class ClientObject {
		/** The request context associated with the object */
		context: ClientRequestContext;
		/** Returns a boolean value for whether the corresponding object is a null object. You must call "context.sync()" before reading the isNullObject property. */
		isNullObject: boolean;
	}
}
declare namespace OfficeExtension {
	interface LoadOption {
		select?: string | string[];
		expand?: string | string[];
		top?: number;
		skip?: number;
	}
	export declare interface UpdateOptions {
		/**
		 * Throw an error if the passed-in property list includes read-only properties (default = true).
		 */
		throwOnReadOnly?: boolean
	}

	/** Contains debug information about the request context. */
	export declare interface RequestContextDebugInfo {
		/** The statements to be executed in the host. */
		pendingStatements: string[];
	}

	/** An abstract RequestContext object that facilitates requests to the host Office application. The "Excel.run" and "Word.run" methods provide a request context. */
	class ClientRequestContext {
		constructor(url?: string);

		/** Collection of objects that are tracked for automatic adjustments based on surrounding changes in the document. */
		trackedObjects: TrackedObjects;

		/** Request headers */
		requestHeaders: { [name: string]: string };

		/** Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties. */
		load(object: ClientObject, option?: string | string[]| LoadOption): void;

		/**
		* Queues up a command to recursively load the specified properties of the object and its navigation properties.
		* You must call "context.sync()" before reading the properties.
		* 
		* @param object The object to be loaded.
		* @param options The key-value pairing of load options for the types, such as { "Workbook": "worksheets,tables",  "Worksheet": "tables",  "Tables": "name" }
		* @param maxDepth The maximum recursive depth.
		*/
		loadRecursive(object: ClientObject, options: { [typeName: string]: string | string[] | LoadOption }, maxDepth?: number): void;

		/** Adds a trace message to the queue. If the promise returned by "context.sync()" is rejected due to an error, this adds a ".traceMessages" array to the OfficeExtension.Error object, containing all trace messages that were executed. These messages can help you monitor the program execution sequence and detect the cause of the error. */
		trace(message: string): void;

		/** Synchronizes the state between JavaScript proxy objects and the Office document, by executing instructions queued on the request context and retrieving properties of loaded Office objects for use in your code.�This method returns a promise, which is resolved when the synchronization is complete. */
		sync<T>(passThroughValue?: T): Promise<T>;

		/** Debug information */
		readonly debugInfo: RequestContextDebugInfo;
	}
}
declare namespace OfficeExtension {
	/** Contains the result for methods that return primitive types. The object's value property is retrieved from the document after "context.sync()" is invoked. */
	class ClientResult<T> {
		/** The value of the result that is retrieved from the document after "context.sync()" is invoked. */
		value: T;
	}

	type RetrieveResult<T extends ClientObject, TData> = { $proxy: T; $isNullObject: boolean; toJSON: () => TData; } & TData;
}
declare namespace OfficeExtension {
	/** Configuration */
	export declare var config: {
		/**
		 * Determines whether to have extended error logging on failure.
		 *
		 * When true, the error object will include a "debugInfo.fullStatements" property that lists out all the actions that were part of the batch request, both before and after the point of failure.
		 *
		 * Having this feature on will introduce a performance penalty, and will also log possibly-sensitive data (e.g., the contents of the commands being sent to the host).
		 * It is recommended that you only have it on during debugging.  Also, if you are logging the error.debugInfo to a database or analytics service,
		 * you should strip out the "debugInfo.fullStatements" property before sending it.
		 */
		extendedErrorLogging: boolean;
	};

	export interface DebugInfo {
		/** Error code string, such as "InvalidArgument". */
		code: string;
		/** The error message passed through from the host Office application. */
		message: string;
		/** Inner error, if applicable. */
		innerError?: DebugInfo | string;

		/** The object type and property or method name (or similar information), if available. */
		errorLocation?: string;
		/** The statement associated with error (or similar information), if available. */
		statement?: string;
		/** The surrounding statements associated with error, if available. */
		surroundingStatements?: string[];
		/** The full statements of the request, if available.*/
		fullStatements?: string[];
	}

	/** The error object returned by "context.sync()", if a promise is rejected due to an error while processing the request. */
	class Error {
		/** Error name: "OfficeExtension.Error".*/
		name: string;
		/** The error message passed through from the host Office application. */
		message: string;
		/** Stack trace, if applicable. */
		stack: string;
		/** Error code string, such as "InvalidArgument". */
		code: string;
		/** Trace messages (if any) that were added via a "context.trace()" invocation before calling "context.sync()". If there was an error, this contains all trace messages that were executed before the error occurred. These messages can help you monitor the program execution sequence and detect the case of the error. */
		traceMessages: Array<string>;
		/** Debug info (useful for detailed logging of the error, i.e., via JSON.stringify(...)). */
		debugInfo: DebugInfo;
		/** Inner error, if applicable. */
		innerError: Error;
	}
}
declare namespace OfficeExtension {
	class ErrorCodes {
		public static accessDenied: string;
		public static generalException: string;
		public static activityLimitReached: string;
		public static invalidObjectPath: string;
		public static propertyNotLoaded: string;
		public static valueNotLoaded: string;
		public static invalidRequestContext: string;
		public static invalidArgument: string;
		public static runMustReturnPromise: string;
		public static cannotRegisterEvent: string;
		public static apiNotFound: string;
		public static connectionFailure: string;
	}
}
declare namespace OfficeExtension {
	/** An Promise object that represents a deferred interaction with the host Office application. The publically-consumable OfficeExtension.Promise is available starting in ExcelApi 1.2 and WordApi 1.2. Promises can be chained via ".then", and errors can be caught via ".catch". Remember to always use a ".catch" on the outer promise, and to return intermediary promises so as not to break the promise chain. When a "native" Promise implementation is available, OfficeExtension.Promise will switch to use the native Promise instead. */
	export const Promise: PromiseConstructor;
}



declare namespace OfficeExtension {
	/** Collection of tracked objects, contained within a request context. See "context.trackedObjects" for more information. */
	class TrackedObjects {
		/** Track a new object for automatic adjustment based on surrounding changes in the document. Only some object types require this. If you are using an object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created. */
		add(object: ClientObject): void;
		/** Track a new object for automatic adjustment based on surrounding changes in the document. Only some object types require this. If you are using an object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created. */
		add(objects: ClientObject[]): void;
		/** Release the memory associated with an object that was previously added to this collection. Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect. */
		remove(object: ClientObject): void;
		/** Release the memory associated with an object that was previously added to this collection. Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect. */
		remove(objects: ClientObject[]): void;
	}
}

declare namespace OfficeExtension {
	export class EventHandlers<T> {
		constructor(context: ClientRequestContext, parentObject: ClientObject, name: string, eventInfo: EventInfo<T>);
		add(handler: (args: T) => Promise<any>): EventHandlerResult<T>;
		remove(handler: (args: T) => Promise<any>): void;
	}

	export class EventHandlerResult<T> {
		constructor(context: ClientRequestContext, handlers: EventHandlers<T>, handler: (args: T) => Promise<any>);
		/** The request context associated with the object */
		context: ClientRequestContext;
		remove(): void;
	}

	export interface EventInfo<T> {
		registerFunc: (callback: (args: any) => void) => Promise<any>;
		unregisterFunc: (callback: (args: any) => void) => Promise<any>;
		eventArgsTransformFunc: (args: any) => Promise<T>;
	}
}
declare namespace OfficeExtension {
	/**
	* Request URL and headers 
	*/
	interface RequestUrlAndHeaderInfo {
		/** Request URL */
		url: string;
		/** Request headers */
		headers?: {
			[name: string]: string;
		};
	}
}
