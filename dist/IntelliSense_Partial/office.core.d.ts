declare namespace OfficeCore {
    var BeginFirstPartyOnlyIntelliSenseBiShimClass: any;
    /**
     * [Api set: AgaveVisual 0.5]
     */
    class BiShim extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext; 
        initialize(capabilities: string): void;
        getData(): OfficeExtension.ClientResult<string>;
        setVisualObjects(visualObjects: string): void;
        setVisualObjectsToPersist(visualObjectsToPersist: string): void;
        /**
         * Create a new instance of OfficeCore.BiShim object
         */
        static newObject(context: OfficeExtension.ClientRequestContext): OfficeCore.BiShim;
        toJSON(): {
            [key: string]: string;
        };
    }
    var EndFirstPartyOnlyIntelliSenseBiShimClass: any;
    enum AgaveVisualErrorCodes {
        generalException = "GeneralException",
    }
    module Interfaces {
        interface CollectionLoadOptions {
            $top?: number;
            $skip?: number;
        }
    }
}
declare namespace OfficeCore {
    var BeginFirstPartyOnlyIntelliSenseFlightingServiceClass: any;
    /**
     * [Api set: Experiment 1.1]
     */
    class FlightingService extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext; 
        getClientSessionId(): OfficeExtension.ClientResult<string>;
        getDeferredFlights(): OfficeExtension.ClientResult<string>;
        getFeature(featureName: string, type: OfficeCore.FeatureType, defaultValue: number | boolean | string, possibleValues?: Array<number> | Array<string> | Array<boolean> | Array<ScopedValue>): OfficeCore.ABType;
        getFeature(featureName: string, type: "Boolean" | "Integer" | "String", defaultValue: number | boolean | string, possibleValues?: Array<number> | Array<string> | Array<boolean> | Array<ScopedValue>): OfficeCore.ABType;
        getFeatureGate(featureName: string, scope?: string): OfficeCore.ABType;
        resetOverride(featureName: string): void;
        setOverride(featureName: string, type: OfficeCore.FeatureType, value: number | boolean | string): void;
        setOverride(featureName: string, type: "Boolean" | "Integer" | "String", value: number | boolean | string): void;
        /**
         * Create a new instance of OfficeCore.FlightingService object
         */
        static newObject(context: OfficeExtension.ClientRequestContext): OfficeCore.FlightingService;
        toJSON(): {
            [key: string]: string;
        };
    }
    var EndFirstPartyOnlyIntelliSenseFlightingServiceClass: any;
    var BeginFirstPartyOnlyIntelliSenseScopedValueClass: any;
    /**
     *
     * Provides information about the scoped value.
     *
     * [Api set: Experiment 1.1]
     */
    interface ScopedValue {
        /**
         *
         * Gets the scope.
         *
         * [Api set: Experiment 1.1]
         */
        scope: string;
        /**
         *
         * Gets the value.
         *
         * [Api set: Experiment 1.1]
         */
        value: string | number | boolean;
    }
    var EndFirstPartyOnlyIntelliSenseScopedValueClass: any;
    var BeginFirstPartyOnlyIntelliSenseABTypeClass: any;
    /**
     * [Api set: Experiment 1.1]
     */
    class ABType extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext; 
        readonly value: string | number | boolean;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: OfficeCore.Interfaces.ABTypeLoadOptions): OfficeCore.ABType;
        load(option?: string | string[]): OfficeCore.ABType;
        load(option?: {
            select?: string;
            expand?: string;
        }): OfficeCore.ABType;
        toJSON(): OfficeCore.Interfaces.ABTypeData;
    }
    var EndFirstPartyOnlyIntelliSenseABTypeClass: any;
    var BeginFirstPartyOnlyIntelliSenseFeatureTypeClass: any;
    /**
     * [Api set: Experiment 1.1]
     */
    enum FeatureType {
        boolean = "Boolean",
        integer = "Integer",
        string = "String",
    }
    var EndFirstPartyOnlyIntelliSenseFeatureTypeClass: any;
    enum ExperimentErrorCodes {
        generalException = "GeneralException",
    }
    module Interfaces {
        interface CollectionLoadOptions {
            $top?: number;
            $skip?: number;
        }
        /** An interface describing the data returned by calling "abtype.toJSON()". */
        interface ABTypeData {
            value?: string | number | boolean;
        }
        /**
         * [Api set: Experiment 1.1]
         */
        interface ABTypeLoadOptions {
            $all?: boolean;
            value?: boolean;
        }
    }
}
declare namespace OfficeCore {
    const OfficeOnlineDomainList: string[];
    function isHostOriginTrusted(): boolean;
}
declare namespace OfficeCore {
    let BeginFirstPartyOnlyIntelliSenseFirstPartyApis: any;
    class FirstPartyApis {
        private context;
        constructor(context: RequestContext);
        readonly roamingSettings: RoamingSettingCollection;
        readonly tap: Tap;
        readonly skill: Skill;
    }
    let EndFirstPartyOnlyIntelliSenseFirstPartyApis: any;
    class RequestContext extends OfficeExtension.ClientRequestContext {
        constructor(url?: string | OfficeExtension.RequestUrlAndHeaderInfo | any);
        static BeginFirstPartyOnlyIntelliSense: any;
        readonly firstParty: FirstPartyApis;
        readonly flighting: FlightingService;
        readonly telemetry: TelemetryService;
        readonly bi: BiShim;
        static EndFirstPartyOnlyIntelliSense: any;
    }
    /**
     * Executes a batch script that performs actions on the Office object model, using a new RequestContext. When the promise is resolved, any tracked objects that were automatically allocated during execution will be released.
     * @param batch - A function that takes in a RequestContext and returns a promise (typically, just the result of "context.sync()"). The context parameter facilitates requests to the Office application. Since the Office add-in and the Office application run in two different processes, the RequestContext is required to get access to the Office object model from the add-in.
     */
    function run<T>(batch: (context: OfficeCore.RequestContext) => Promise<T>): Promise<T>;
    /**
     * Executes a batch script that performs actions on the Office object model, using the RequestContext of a previously-created object. When the promise is resolved, any tracked objects that were automatically allocated during execution will be released.
     * @param context - A previously-created context object. The batch will use the same RequestContext as the passed-in object, which means that any changes applied to the object will be picked up by "context.sync()".
     * @param batch - A function that takes in a RequestContext and returns a promise (typically, just the result of "context.sync()"). The context parameter facilitates requests to the Office application. Since the Office add-in and the Office application run in two different processes, the RequestContext is required to get access to the Office object model from the add-in.
     */
    function run<T>(context: OfficeCore.RequestContext, batch: (context: OfficeCore.RequestContext) => Promise<T>): Promise<T>;
    /**
     * Executes a batch script that performs actions on the Office object model, using the RequestContext of a previously-created API object. When the promise is resolved, any tracked objects that were automatically allocated during execution will be released.
     * @param object - A previously-created API object. The batch will use the same RequestContext as the passed-in object, which means that any changes applied to the object will be picked up by "context.sync()".
     * @param batch - A function that takes in a RequestContext and returns a promise (typically, just the result of "context.sync()"). The context parameter facilitates requests to the Office application. Since the Office add-in and the Office application run in two different processes, the RequestContext is required to get access to the Office object model from the add-in.
     */
    function run<T>(object: OfficeExtension.ClientObject, batch: (context: OfficeCore.RequestContext) => Promise<T>): Promise<T>;
    /**
     * Executes a batch script that performs actions on the Office object model, using the RequestContext of previously-created API objects.
     * @param objects - An array of previously-created API objects. The array will be validated to make sure that all of the objects share the same context. The batch will use this shared RequestContext, which means that any changes applied to these objects will be picked up by "context.sync()".
     * @param batch - A function that takes in a RequestContext and returns a promise (typically, just the result of "context.sync()"). The context parameter facilitates requests to the Office application. Since the Office add-in and the Office application run in two different processes, the RequestContext is required to get access to the Office object model from the add-in.
     */
    function run<T>(objects: OfficeExtension.ClientObject[], batch: (context: OfficeCore.RequestContext) => Promise<T>): Promise<T>;
}
declare namespace OfficeCore {
    var BeginFirstPartyOnlyIntelliSenseSkillEventArgsClass: any;
    /**
     *
     * Provides information about the new result.
     *
     * [Api set: Skill 1.1]
     */
    interface SkillEventArgs {
        /**
         *
         * The serialized JSON string of the event data object.
         *
         * [Api set: Skill 1.1]
         */
        data: string;
        /**
         *
         * The integer represented event type, e.g. NewSearch, NewSkillResultAvailable.
         *
         * [Api set: Skill 1.1]
         */
        type: number;
    }
    var EndFirstPartyOnlyIntelliSenseSkillEventArgsClass: any;
    var BeginFirstPartyOnlyIntelliSenseSkillClass: any;
    /**
     *
     * Represents a collection of Apis for Skill feature.
     *
     * [Api set: Skill 1.1]
     */
    class Skill extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext; 
        /**
         *
         * Perform an action from the Skill result pane
         *
         * [Api set: Skill 1.1]
         *
         * @param paneId Required. The id of the Skill result pane where the call is made.
         * @param actionId Required. The unique id of the action, which is usually a GUID
         * @param actionDescriptor Required. The serialized JSON string of the action descriptor object.
         * @returns The serialized JSON string of the return object for the action.
         */
        executeAction(paneId: string, actionId: string, actionDescriptor: string): OfficeExtension.ClientResult<string>;
        /**
         *
         * Notify host with any event fired from the Skill result pane
         *
         * [Api set: Skill 1.1]
         *
         * @param paneId Required. The id of the Skill result pane where the event is fired.
         * @param eventDescriptor Required. The serialized JSON string of the event descriptor object.
         */
        notifyPaneEvent(paneId: string, eventDescriptor: string): void;
        /**
         *
         * To register an handler to the HostSkillEvent
         *
         * [Api set: Skill 1.1]
         */
        registerHostSkillEvent(): void;
        /**
         *
         * TEST ONLY
         *
         * [Api set: Skill 1.1]
         */
        testFireEvent(): void;
        /**
         *
         * To unregister an handler to the HostSkillEvent
         *
         * [Api set: Skill 1.1]
         */
        unregisterHostSkillEvent(): void;
        /**
         * Create a new instance of OfficeCore.Skill object
         */
        static newObject(context: OfficeExtension.ClientRequestContext): OfficeCore.Skill;
        /**
         *
         * Fire whenever there is a new Skill event from host
         *
         * [Api set: Skill 1.1]
         *
         * @eventproperty
         */
        readonly onHostSkillEvent: OfficeExtension.EventHandlers<OfficeCore.SkillEventArgs>;
        toJSON(): {
            [key: string]: string;
        };
    }
    var EndFirstPartyOnlyIntelliSenseSkillClass: any;
    enum SkillErrorCodes {
        generalException = "GeneralException",
    }
    module Interfaces {
        /**
        * Provides ways to load properties of only a subset of members of a collection.
        */
        interface CollectionLoadOptions {
            /**
            * Specify the number of items in the queried collection to be included in the result.
            */
            $top?: number;
            /**
            * Specify the number of items in the collection that are to be skipped and not included in the result. If top is specified, the selection of result will start after skipping the specified number of items.
            */
            $skip?: number;
        }
    }
}
declare namespace OfficeCore {
    var BeginFirstPartyOnlyIntelliSenseTelemetryServiceClass: any;
    /**
     * [Api set: Telemetry 1.1]
     */
    class TelemetryService extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext; 
        sendTelemetryEvent(telemetryProperties: OfficeCore.TelemetryProperties, eventName: string, eventContract: string, eventFlags: OfficeCore.EventFlags, value: OfficeCore.DataField[]): void;
        /**
         * Create a new instance of OfficeCore.TelemetryService object
         */
        static newObject(context: OfficeExtension.ClientRequestContext): OfficeCore.TelemetryService;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original OfficeCore.TelemetryService object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `OfficeCore.Interfaces.TelemetryServiceData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): {
            [key: string]: string;
        };
    }
    var EndFirstPartyOnlyIntelliSenseTelemetryServiceClass: any;
    var BeginFirstPartyOnlyIntelliSenseEventFlagsClass: any;
    /**
     * [Api set: Telemetry 1.1]
     */
    interface EventFlags {
        costPriority: number;
        dataCategories: number;
        diagnosticLevel: number;
        persistencePriority: number;
        samplingPolicy: number;
    }
    var EndFirstPartyOnlyIntelliSenseEventFlagsClass: any;
    var BeginFirstPartyOnlyIntelliSenseDataFieldClass: any;
    /**
     * [Api set: Telemetry 1.1]
     */
    interface DataField {
        classification: number;
        name: string;
        type?: OfficeCore.DataFieldType | "Unset" | "String" | "Boolean" | "Int64" | "Double";
        value: any;
    }
    var EndFirstPartyOnlyIntelliSenseDataFieldClass: any;
    var BeginFirstPartyOnlyIntelliSenseTelemetryPropertiesClass: any;
    /**
     * [Api set: Telemetry 1.1]
     */
    interface TelemetryProperties {
        ariaTenantToken?: string;
        nexusTenantToken?: number;
    }
    var EndFirstPartyOnlyIntelliSenseTelemetryPropertiesClass: any;
    var BeginFirstPartyOnlyIntelliSenseDataFieldTypeClass: any;
    /**
     * [Api set: Telemetry 1.1]
     */
    enum DataFieldType {
        unset = "Unset",
        string = "String",
        boolean = "Boolean",
        int64 = "Int64",
        double = "Double",
    }
    var EndFirstPartyOnlyIntelliSenseDataFieldTypeClass: any;
    enum TelemetryErrorCodes {
        generalException = "GeneralException",
    }
    module Interfaces {
        /**
        * Provides ways to load properties of only a subset of members of a collection.
        */
        interface CollectionLoadOptions {
            /**
            * Specify the number of items in the queried collection to be included in the result.
            */
            $top?: number;
            /**
            * Specify the number of items in the collection that are to be skipped and not included in the result. If top is specified, the selection of result will start after skipping the specified number of items.
            */
            $skip?: number;
        }
    }
}
declare namespace OfficeFirstPartyAuth {
    function getAccessToken(options: OfficeCore.TokenParameters): Promise<OfficeCore.SingleSignOnToken>;
    function getPrimaryIdentityInfo(): Promise<OfficeCore.OfficeIdentityInfo>;
}
declare namespace OfficeCore {
    var BeginFirstPartyOnlyIntelliSenseIdentityTypeClass: any;
    /**
     * [Api set: FirstPartyAuthentication 1.1]
     */
    enum IdentityType {
        organizationAccount = "OrganizationAccount",
        microsoftAccount = "MicrosoftAccount",
    }
    var EndFirstPartyOnlyIntelliSenseIdentityTypeClass: any;
    var BeginFirstPartyOnlyIntelliSenseTokenReceivedEventArgsClass: any;
    /**
     *
     * Office identity object that holds the user information
     *
     * [Api set: FirstPartyAuthentication 1.2]
     */
    interface TokenReceivedEventArgs {
        /**
         *
         * Return code. 0 means success and otherwise it's the error code.
         *
         * [Api set: FirstPartyAuthentication 1.2]
         */
        code: number;
        /**
         *
         * Return error JSON string.
         *
         * [Api set: FirstPartyAuthentication 1.2]
         */
        errorInfo: string;
        /**
         *
         * The identity SingleSignOnToken, which contains the accesstoken string and identity account type.
         *
         * [Api set: FirstPartyAuthentication 1.2]
         */
        tokenValue: OfficeCore.SingleSignOnToken;
    }
    var EndFirstPartyOnlyIntelliSenseTokenReceivedEventArgsClass: any;
    var BeginFirstPartyOnlyIntelliSenseOfficeIdentityInfoClass: any;
    /**
     *
     * Office identity object that holds the user information
     *
     * [Api set: FirstPartyAuthentication 1.2]
     */
    interface OfficeIdentityInfo {
        /**
         *
         * A display name for the user
         *
         * [Api set: FirstPartyAuthentication 1.2]
         */
        displayName: string;
        /**
         *
         * The Email address associated with the identity
         *
         * [Api set: FirstPartyAuthentication 1.2]
         */
        email: string;
        /**
         *
         * The federation provider (such as Worldwide, BlackForest or Gallatin)
         *
         * [Api set: FirstPartyAuthentication 1.2]
         */
        federationProvider: string;
        /**
         *
         * The identity account type
         *
         * [Api set: FirstPartyAuthentication 1.2]
         */
        identityType: OfficeCore.IdentityType | "OrganizationAccount" | "MicrosoftAccount";
    }
    var EndFirstPartyOnlyIntelliSenseOfficeIdentityInfoClass: any;
    var BeginFirstPartyOnlyIntelliSenseAuthenticationServiceClass: any;
    /**
     * [Api set: FirstPartyAuthentication 1.1]
     */
    class AuthenticationService extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext; 
        /**
         *
         * Gets the collection of Roaming Settings
         *
         * [Api set: FirstPartyAuthentication 1.1]
         */
        readonly roamingSettings: OfficeCore.RoamingSettingCollection;
        /**
         *
         * Get the access token for the current primary identity.
         *
         * [Api set: FirstPartyAuthentication 1.1]
         *
         * @param tokenParameters The parameter for the required access token.
         * @param targetId The parameter for matching client event with original request.
         * @returns The access token object.
         */
        getAccessToken(tokenParameters: OfficeCore.TokenParameters, targetId: string): OfficeExtension.ClientResult<OfficeCore.SingleSignOnToken>;
        /**
         *
         * Get the information of the primary identity (in rich client, it's the active profile).
         *
         * [Api set: FirstPartyAuthentication 1.2]
         * @returns The primary identity type.
         */
        getPrimaryIdentityInfo(): OfficeExtension.ClientResult<OfficeCore.OfficeIdentityInfo>;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         *
         * @remarks
         *
         * In addition to this signature, this method has the following signatures:
         *
         * `load(option?: string | string[]): OfficeCore.AuthenticationService` - Where option is a comma-delimited string or an array of strings that specify the properties to load.
         *
         * `load(option?: { select?: string; expand?: string; }): OfficeCore.AuthenticationService` - Where option.select is a comma-delimited string that specifies the properties to load, and options.expand is a comma-delimited string that specifies the navigation properties to load.
         *
         * `load(option?: { select?: string; expand?: string; top?: number; skip?: number }): OfficeCore.AuthenticationService` - Only available on collection types. It is similar to the preceding signature. Option.top specifies the maximum number of collection items that can be included in the result. Option.skip specifies the number of items that are to be skipped and not included in the result. If option.top is specified, the result set will start after skipping the specified number of items.
         *
         * @param options Provides options for which properties of the object to load.
         */
        load(option?: string | string[]): OfficeCore.AuthenticationService;
        load(option?: {
            select?: string;
            expand?: string;
        }): OfficeCore.AuthenticationService;
        /**
         * Create a new instance of OfficeCore.AuthenticationService object
         */
        static newObject(context: OfficeExtension.ClientRequestContext): OfficeCore.AuthenticationService;
        /**
         *
         * Occurs when token data comes back.
         *
         * [Api set: FirstPartyAuthentication 1.2]
         *
         * @eventproperty
         */
        readonly onTokenReceived: OfficeExtension.EventHandlers<OfficeCore.TokenReceivedEventArgs>;
        toJSON(): OfficeCore.Interfaces.AuthenticationServiceData;
    }
    var EndFirstPartyOnlyIntelliSenseAuthenticationServiceClass: any;
    var BeginFirstPartyOnlyIntelliSenseTokenParametersClass: any;
    /**
     * [Api set: FirstPartyAuthentication 1.1]
     */
    interface TokenParameters {
        /**
         *
         * The auth challenge string.
         *
         * [Api set: FirstPartyAuthentication 1.1]
         */
        authChallenge?: string;
        /**
         *
         * The auth policy string.
         *
         * [Api set: FirstPartyAuthentication 1.1]
         */
        policy?: string;
        /**
         *
         * The resource URL (or target)
         *
         * [Api set: FirstPartyAuthentication 1.1]
         */
        resource?: string;
    }
    var EndFirstPartyOnlyIntelliSenseTokenParametersClass: any;
    var BeginFirstPartyOnlyIntelliSenseSingleSignOnTokenClass: any;
    /**
     * [Api set: FirstPartyAuthentication 1.1]
     */
    interface SingleSignOnToken {
        /**
         *
         * The access token for the primary identity.
         *
         * [Api set: FirstPartyAuthentication 1.1]
         */
        accessToken: string;
        /**
         *
         * The identity type associated with the access token
         *
         * [Api set: FirstPartyAuthentication 1.1]
         */
        tokenIdenityType: OfficeCore.IdentityType | "OrganizationAccount" | "MicrosoftAccount";
    }
    var EndFirstPartyOnlyIntelliSenseSingleSignOnTokenClass: any;
    var BeginFirstPartyOnlyIntelliSenseRoamingSettingClass: any;
    /**
     *
     * Represents a roaming setting object.
     *
     * [Api set: Roaming 1.1]
     */
    class RoamingSetting extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext; 
        /**
         *
         * Returns the Id that represents the id of the Roaming Setting. Read-only.
         *
         * [Api set: Roaming 1.1]
         */
        readonly id: string;
        /**
         *
         * Represents the value stored for this setting.
         *
         * [Api set: Roaming 1.1]
         */
        value: any;
        /** Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
         *
         * @remarks
         *
         * This method has the following additional signature:
         *
         * `set(properties: OfficeCore.RoamingSetting): void`
         *
         * @param properties A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
         * @param options Provides an option to suppress errors if the properties object tries to set any read-only properties.
         */
        set(properties: Interfaces.RoamingSettingUpdateData, options?: OfficeExtension.UpdateOptions): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: OfficeCore.RoamingSetting): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         *
         * @remarks
         *
         * In addition to this signature, this method has the following signatures:
         *
         * `load(option?: string | string[]): OfficeCore.RoamingSetting` - Where option is a comma-delimited string or an array of strings that specify the properties to load.
         *
         * `load(option?: { select?: string; expand?: string; }): OfficeCore.RoamingSetting` - Where option.select is a comma-delimited string that specifies the properties to load, and options.expand is a comma-delimited string that specifies the navigation properties to load.
         *
         * `load(option?: { select?: string; expand?: string; top?: number; skip?: number }): OfficeCore.RoamingSetting` - Only available on collection types. It is similar to the preceding signature. Option.top specifies the maximum number of collection items that can be included in the result. Option.skip specifies the number of items that are to be skipped and not included in the result. If option.top is specified, the result set will start after skipping the specified number of items.
         *
         * @param options Provides options for which properties of the object to load.
         */
        load(option?: OfficeCore.Interfaces.RoamingSettingLoadOptions): OfficeCore.RoamingSetting;
        load(option?: string | string[]): OfficeCore.RoamingSetting;
        load(option?: {
            select?: string;
            expand?: string;
        }): OfficeCore.RoamingSetting;
        toJSON(): OfficeCore.Interfaces.RoamingSettingData;
    }
    var EndFirstPartyOnlyIntelliSenseRoamingSettingClass: any;
    var BeginFirstPartyOnlyIntelliSenseRoamingSettingCollectionClass: any;
    /**
     *
     * Contains the collection of roaming setting objects.
     *
     * [Api set: Roaming 1.1]
     */
    class RoamingSettingCollection extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext; 
        /**
         *
         * Gets a roaming setting object by its key. Returns a null object if not found.
         *
         * [Api set: Roaming 1.1]
         *
         * @param id
         * @returns
         */
        getItem(id: string): OfficeCore.RoamingSetting;
        /**
         *
         * Gets a roaming setting object by its key. Returns a null object if not found.
         *
         * [Api set: Roaming 1.1]
         *
         * @param id
         * @returns
         */
        getItemOrNullObject(id: string): OfficeCore.RoamingSetting;
        toJSON(): {
            [key: string]: string;
        };
    }
    var EndFirstPartyOnlyIntelliSenseRoamingSettingCollectionClass: any;
    /**
     *
     * Represents a single comment in the document.
     *
     * [Api set: Comments 1.1]
     */
    class Comment extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext; 
        /**
         *
         * Gets this comment's parent. If this is a root comment, throws.
         *
         * [Api set: Comments 1.1]
         */
        readonly parent: OfficeCore.Comment;
        /**
         *
         * Gets this comment's parent. If this is a root comment, returns a null object.
         *
         * [Api set: Comments 1.1]
         */
        readonly parentOrNullObject: OfficeCore.Comment;
        /**
         *
         * Gets the replies to this comment. If this is not a root comment, returns an empty collection.
         *
         * [Api set: Comments 1.1]
         */
        readonly replies: OfficeCore.CommentCollection;
        /**
         *
         * Gets an object representing the comment's author. Read-only.
         *
         * [Api set: Comments 1.1]
         */
        readonly author: OfficeCore.CommentAuthor;
        /**
         *
         * Gets when the comment was created. Read-only.
         *
         * [Api set: Comments 1.1]
         */
        readonly created: Date;
        /**
         *
         * Returns a value that uniquely identifies the comment in a given document. Read-only.
         *
         * [Api set: Comments 1.1]
         */
        readonly id: string;
        /**
         *
         * Gets the level of the comment: 0 if it is a root comment, or 1 if it is a reply. Read-only.
         *
         * [Api set: Comments 1.1]
         */
        readonly level: number;
        /**
         *
         * Gets the comment's mentions.
         *
         * [Api set: Comments 1.1]
         */
        readonly mentions: OfficeCore.CommentMention[];
        /**
         *
         * Gets or sets whether this comment is resolved.
         *
         * [Api set: Comments 1.1]
         */
        resolved: boolean;
        /**
         *
         * Gets or sets the comment's plain text, without formatting.
         *
         * [Api set: Comments 1.1]
         */
        text: string;
        /** Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
         *
         * @remarks
         *
         * This method has the following additional signature:
         *
         * `set(properties: OfficeCore.Comment): void`
         *
         * @param properties A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
         * @param options Provides an option to suppress errors if the properties object tries to set any read-only properties.
         */
        set(properties: Interfaces.CommentUpdateData, options?: OfficeExtension.UpdateOptions): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: OfficeCore.Comment): void;
        /**
         *
         * Deletes this comment. If this is a root comment, deletes the entire comment thread.
         *
         * [Api set: Comments 1.1]
         */
        delete(): void;
        /**
         *
         * Gets this comment's parent. If this is a root comment, returns a new comment object representing itself.
            This method is useful for accessing thread-level properties from either a reply or the root comment.
            
            e.g. comment.getParentOrSelf().resolved = true;
         *
         * [Api set: Comments 1.1]
         */
        getParentOrSelf(): OfficeCore.Comment;
        /**
         *
         * Gets the comment's rich text in the specified markup format.
         *
         * [Api set: Comments 1.1]
         */
        getRichText(format: OfficeCore.CommentTextFormat): OfficeExtension.ClientResult<string>;
        /**
         *
         * Gets the comment's rich text in the specified markup format.
         *
         * [Api set: Comments 1.1]
         */
        getRichText(format: "Plain" | "Markdown" | "Delta"): OfficeExtension.ClientResult<string>;
        /**
         *
         * Appends a new reply to the comment's thread.
         *
         * [Api set: Comments 1.1]
         *
         * @param text The body of the reply.
         * @param format The markup format of the text parameter.
         * @returns
         */
        reply(text: string, format: OfficeCore.CommentTextFormat): OfficeCore.Comment;
        /**
         *
         * Appends a new reply to the comment's thread.
         *
         * [Api set: Comments 1.1]
         *
         * @param text The body of the reply.
         * @param format The markup format of the text parameter.
         * @returns
         */
        reply(text: string, format: "Plain" | "Markdown" | "Delta"): OfficeCore.Comment;
        /**
         *
         * Sets the comment's rich text.
         *
         * [Api set: Comments 1.1]
         *
         * @param text The text of the comment.
         * @param format The markup format of the 'text' parameter.
         */
        setRichText(text: string, format: OfficeCore.CommentTextFormat): OfficeExtension.ClientResult<string>;
        /**
         *
         * Sets the comment's rich text.
         *
         * [Api set: Comments 1.1]
         *
         * @param text The text of the comment.
         * @param format The markup format of the 'text' parameter.
         */
        setRichText(text: string, format: "Plain" | "Markdown" | "Delta"): OfficeExtension.ClientResult<string>;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         *
         * @remarks
         *
         * In addition to this signature, this method has the following signatures:
         *
         * `load(option?: string | string[]): OfficeCore.Comment` - Where option is a comma-delimited string or an array of strings that specify the properties to load.
         *
         * `load(option?: { select?: string; expand?: string; }): OfficeCore.Comment` - Where option.select is a comma-delimited string that specifies the properties to load, and options.expand is a comma-delimited string that specifies the navigation properties to load.
         *
         * `load(option?: { select?: string; expand?: string; top?: number; skip?: number }): OfficeCore.Comment` - Only available on collection types. It is similar to the preceding signature. Option.top specifies the maximum number of collection items that can be included in the result. Option.skip specifies the number of items that are to be skipped and not included in the result. If option.top is specified, the result set will start after skipping the specified number of items.
         *
         * @param options Provides options for which properties of the object to load.
         */
        load(option?: OfficeCore.Interfaces.CommentLoadOptions): OfficeCore.Comment;
        load(option?: string | string[]): OfficeCore.Comment;
        load(option?: {
            select?: string;
            expand?: string;
        }): OfficeCore.Comment;
        toJSON(): OfficeCore.Interfaces.CommentData;
    }
    /**
     *
     * Represents the author of a comment.
     *
     * [Api set: Comments 1.1]
     */
    interface CommentAuthor {
        /**
         *
         * The email address of the author.
         *
         * [Api set: Comments 1.1]
         */
        email: string;
        /**
         *
         * The name of the author.
         *
         * [Api set: Comments 1.1]
         */
        name: string;
    }
    /**
     *
     * Represents a mention within a comment.
     *
     * [Api set: Comments 1.1]
     */
    interface CommentMention {
        /**
         *
         * The email address of the person mentioned.
         *
         * [Api set: Comments 1.1]
         */
        email: string;
        /**
         *
         * The name of the person mentioned.
         *
         * [Api set: Comments 1.1]
         */
        name: string;
        /**
         *
         * The text displayed for the mention.
         *
         * [Api set: Comments 1.1]
         */
        text: string;
    }
    /**
     *
     * Represents a collection of comments, either replies to a specific comment thread, or all comments in the document or part of the document.
     *
     * [Api set: Comments 1.1]
     */
    class CommentCollection extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext; 
        /** Gets the loaded child items in this collection. */
        readonly items: OfficeCore.Comment[];
        /**
         *
         * Returns the number of comments in the collection. Read-only.
         *
         * [Api set: Comments 1.1]
         */
        getCount(): OfficeExtension.ClientResult<number>;
        /**
         *
         * Gets a comment object using its id.
         *
         * [Api set: Comments 1.1]
         */
        getItem(id: string): OfficeCore.Comment;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         *
         * @remarks
         *
         * In addition to this signature, this method has the following signatures:
         *
         * `load(option?: string | string[]): OfficeCore.CommentCollection` - Where option is a comma-delimited string or an array of strings that specify the properties to load.
         *
         * `load(option?: { select?: string; expand?: string; }): OfficeCore.CommentCollection` - Where option.select is a comma-delimited string that specifies the properties to load, and options.expand is a comma-delimited string that specifies the navigation properties to load.
         *
         * `load(option?: { select?: string; expand?: string; top?: number; skip?: number }): OfficeCore.CommentCollection` - Only available on collection types. It is similar to the preceding signature. Option.top specifies the maximum number of collection items that can be included in the result. Option.skip specifies the number of items that are to be skipped and not included in the result. If option.top is specified, the result set will start after skipping the specified number of items.
         *
         * @param options Provides options for which properties of the object to load.
         */
        load(option?: OfficeCore.Interfaces.CommentCollectionLoadOptions & OfficeCore.Interfaces.CollectionLoadOptions): OfficeCore.CommentCollection;
        load(option?: string | string[]): OfficeCore.CommentCollection;
        load(option?: OfficeExtension.LoadOption): OfficeCore.CommentCollection;
        toJSON(): OfficeCore.Interfaces.CommentCollectionData;
    }
    /**
     *
     * Represents a markup (rich) text format.
     *
     * [Api set: Comments 1.1]
     */
    enum CommentTextFormat {
        plain = "Plain",
        markdown = "Markdown",
        delta = "Delta",
    }
    var BeginFirstPartyOnlyIntelliSenseTapClass: any;
    /**
     *
     * Represents a collection of Apis for Tap feature.
     *
     * [Api set: Tap 1.1]
     */
    class Tap extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext; 
        /**
         *
         * Gets the enterprise user info - including whether Tap feature is enabled, tenant root, and "Go Local" information.
         *
         * [Api set: Tap 1.1]
         * @returns The enterprise user info in Json format.
         */
        getEnterpriseUserInfo(): OfficeExtension.ClientResult<string>;
        /**
         *
         * Returns the user friendly path for the given document URL as shown on MRU.
         *
         * [Api set: Tap 1.1]
         *
         * @param documentUrl Required. Document URL.
         * @returns The user friendly path.
         */
        getMruFriendlyPath(documentUrl: string): OfficeExtension.ClientResult<string>;
        /**
         *
         * Given a document URL, opens the file in the default application. Win32 Only.
         *
         * [Api set: Tap 1.1]
         *
         * @param documentUrl Required. Document URL.
         * @param useUniversalAsBackup Optional. Whether to launch in universal app if win32 desktop app is not available. The default value is false.
         * @returns Whether launch is successful.
         */
        launchFileUrlInOfficeApp(documentUrl: string, useUniversalAsBackup?: boolean): OfficeExtension.ClientResult<boolean>;
        /**
         *
         * Performs local window file index search. Win32 Only.
         *
         * [Api set: Tap 1.1]
         *
         * @param query Required. Search query.
         * @param numResultsRequested Required. Numers of files to return.
         * @param supportedFileExtensions Optional. All supported file extension (should split with ","). The default value is all file extensions for current host application.
         * @param documentUrlToExclude Optional. The document url to exclude from results. The default value is empty.
         * @returns The local document results in Json format.
         */
        performLocalSearch(query: string, numResultsRequested: number, supportedFileExtensions?: string, documentUrlToExclude?: string): OfficeExtension.ClientResult<string>;
        /**
         *
         * Read search cache file from local disk. Win32 Only.
         *
         * [Api set: Tap 1.1]
         *
         * @param keyword Required. Search keyword.
         * @param expiredHours Required. Numers of hours that cache gets expired. If expired, return empty string.
         * @param filterObjectType Required. Selected object type for filter.
         * @returns The search cache file content.
         */
        readSearchCache(keyword: string, expiredHours: number, filterObjectType: OfficeCore.ObjectType): OfficeExtension.ClientResult<string>;
        /**
         *
         * Read search cache file from local disk. Win32 Only.
         *
         * [Api set: Tap 1.1]
         *
         * @param keyword Required. Search keyword.
         * @param expiredHours Required. Numers of hours that cache gets expired. If expired, return empty string.
         * @param filterObjectType Required. Selected object type for filter.
         * @returns The search cache file content.
         */
        readSearchCache(keyword: string, expiredHours: number, filterObjectType: "Unknown" | "Chart" | "SmartArt" | "Table" | "Image" | "Slide" | "OLE" | "Text"): OfficeExtension.ClientResult<string>;
        /**
         *
         * Write search cache file to local disk. Win32 Only.
         *
         * [Api set: Tap 1.1]
         *
         * @param fileContent Required. File content string.
         * @param keyword Required. Search keyword.
         * @param filterObjectType Required. Selected object type for filter.
         * @returns Whether write is successful.
         */
        writeSearchCache(fileContent: string, keyword: string, filterObjectType: OfficeCore.ObjectType): OfficeExtension.ClientResult<boolean>;
        /**
         *
         * Write search cache file to local disk. Win32 Only.
         *
         * [Api set: Tap 1.1]
         *
         * @param fileContent Required. File content string.
         * @param keyword Required. Search keyword.
         * @param filterObjectType Required. Selected object type for filter.
         * @returns Whether write is successful.
         */
        writeSearchCache(fileContent: string, keyword: string, filterObjectType: "Unknown" | "Chart" | "SmartArt" | "Table" | "Image" | "Slide" | "OLE" | "Text"): OfficeExtension.ClientResult<boolean>;
        /**
         * Create a new instance of OfficeCore.Tap object
         */
        static newObject(context: OfficeExtension.ClientRequestContext): OfficeCore.Tap;
        toJSON(): {
            [key: string]: string;
        };
    }
    var EndFirstPartyOnlyIntelliSenseTapClass: any;
    var BeginFirstPartyOnlyIntelliSenseObjectTypeClass: any;
    /**
     *
     * Represents object type for Tap feature.
     *
     * [Api set: Tap 1.1]
     */
    enum ObjectType {
        unknown = "Unknown",
        chart = "Chart",
        smartArt = "SmartArt",
        table = "Table",
        image = "Image",
        slide = "Slide",
        ole = "OLE",
        text = "Text",
    }
    var EndFirstPartyOnlyIntelliSenseObjectTypeClass: any;
    enum ErrorCodes {
        apiNotAvailable = "ApiNotAvailable",
        clientError = "ClientError",
        generalException = "GeneralException",
        interactiveFlowAborted = "InteractiveFlowAborted",
        invalidArgument = "InvalidArgument",
        invalidGrant = "InvalidGrant",
        invalidResourceUrl = "InvalidResourceUrl",
        resourceNotSupported = "ResourceNotSupported",
        serverError = "ServerError",
        unsupportedUserIdentity = "UnsupportedUserIdentity",
        userNotSignedIn = "UserNotSignedIn",
    }
    module Interfaces {
        /**
        * Provides ways to load properties of only a subset of members of a collection.
        */
        interface CollectionLoadOptions {
            /**
            * Specify the number of items in the queried collection to be included in the result.
            */
            $top?: number;
            /**
            * Specify the number of items in the collection that are to be skipped and not included in the result. If top is specified, the selection of result will start after skipping the specified number of items.
            */
            $skip?: number;
        }
        var BeginFirstPartyOnlyIntelliSenseRoamingSettingUpdateData: any;
        /** An interface for updating data on the RoamingSetting object, for use in "roamingSetting.set({ ... })". */
        interface RoamingSettingUpdateData {
            /**
             *
             * Represents the value stored for this setting.
             *
             * [Api set: Roaming 1.1]
             */
            value?: any;
        }
        var EndFirstPartyOnlyIntelliSenseRoamingSettingUpdateData: any;
        /** An interface for updating data on the Comment object, for use in "comment.set({ ... })". */
        interface CommentUpdateData {
            /**
             *
             * Gets or sets whether this comment is resolved.
             *
             * [Api set: Comments 1.1]
             */
            resolved?: boolean;
            /**
             *
             * Gets or sets the comment's plain text, without formatting.
             *
             * [Api set: Comments 1.1]
             */
            text?: string;
        }
        /** An interface for updating data on the CommentCollection object, for use in "commentCollection.set({ ... })". */
        interface CommentCollectionUpdateData {
            items?: OfficeCore.Interfaces.CommentData[];
        }
        var BeginFirstPartyOnlyIntelliSenseAuthenticationServiceData: any;
        /** An interface describing the data returned by calling "authenticationService.toJSON()". */
        interface AuthenticationServiceData {
        }
        var EndFirstPartyOnlyIntelliSenseAuthenticationServiceData: any;
        var BeginFirstPartyOnlyIntelliSenseRoamingSettingData: any;
        /** An interface describing the data returned by calling "roamingSetting.toJSON()". */
        interface RoamingSettingData {
            /**
             *
             * Returns the Id that represents the id of the Roaming Setting. Read-only.
             *
             * [Api set: Roaming 1.1]
             */
            id?: string;
            /**
             *
             * Represents the value stored for this setting.
             *
             * [Api set: Roaming 1.1]
             */
            value?: any;
        }
        var EndFirstPartyOnlyIntelliSenseRoamingSettingData: any;
        /** An interface describing the data returned by calling "comment.toJSON()". */
        interface CommentData {
            /**
            *
            * Gets this comment's parent. If this is a root comment, throws.
            *
            * [Api set: Comments 1.1]
            */
            parent?: OfficeCore.Interfaces.CommentData;
            /**
            *
            * Gets this comment's parent. If this is a root comment, returns a null object.
            *
            * [Api set: Comments 1.1]
            */
            parentOrNullObject?: OfficeCore.Interfaces.CommentData;
            /**
            *
            * Gets the replies to this comment. If this is not a root comment, returns an empty collection.
            *
            * [Api set: Comments 1.1]
            */
            replies?: OfficeCore.Interfaces.CommentData[];
            /**
             *
             * Gets an object representing the comment's author. Read-only.
             *
             * [Api set: Comments 1.1]
             */
            author?: OfficeCore.CommentAuthor;
            /**
             *
             * Gets when the comment was created. Read-only.
             *
             * [Api set: Comments 1.1]
             */
            created?: Date;
            /**
             *
             * Returns a value that uniquely identifies the comment in a given document. Read-only.
             *
             * [Api set: Comments 1.1]
             */
            id?: string;
            /**
             *
             * Gets the level of the comment: 0 if it is a root comment, or 1 if it is a reply. Read-only.
             *
             * [Api set: Comments 1.1]
             */
            level?: number;
            /**
             *
             * Gets the comment's mentions.
             *
             * [Api set: Comments 1.1]
             */
            mentions?: OfficeCore.CommentMention[];
            /**
             *
             * Gets or sets whether this comment is resolved.
             *
             * [Api set: Comments 1.1]
             */
            resolved?: boolean;
            /**
             *
             * Gets or sets the comment's plain text, without formatting.
             *
             * [Api set: Comments 1.1]
             */
            text?: string;
        }
        /** An interface describing the data returned by calling "commentCollection.toJSON()". */
        interface CommentCollectionData {
            items?: OfficeCore.Interfaces.CommentData[];
        }
        var BeginFirstPartyOnlyIntelliSenseRoamingSettingLoadOptions: any;
        /**
         *
         * Represents a roaming setting object.
         *
         * [Api set: Roaming 1.1]
         */
        interface RoamingSettingLoadOptions {
            $all?: boolean;
            /**
             *
             * Returns the Id that represents the id of the Roaming Setting. Read-only.
             *
             * [Api set: Roaming 1.1]
             */
            id?: boolean;
            /**
             *
             * Represents the value stored for this setting.
             *
             * [Api set: Roaming 1.1]
             */
            value?: boolean;
        }
        var EndFirstPartyOnlyIntelliSenseRoamingSettingLoadOptions: any;
        /**
         *
         * Represents a single comment in the document.
         *
         * [Api set: Comments 1.1]
         */
        interface CommentLoadOptions {
            $all?: boolean;
            /**
            *
            * Gets this comment's parent. If this is a root comment, throws.
            *
            * [Api set: Comments 1.1]
            */
            parent?: OfficeCore.Interfaces.CommentLoadOptions;
            /**
            *
            * Gets this comment's parent. If this is a root comment, returns a null object.
            *
            * [Api set: Comments 1.1]
            */
            parentOrNullObject?: OfficeCore.Interfaces.CommentLoadOptions;
            /**
             *
             * Gets an object representing the comment's author. Read-only.
             *
             * [Api set: Comments 1.1]
             */
            author?: boolean;
            /**
             *
             * Gets when the comment was created. Read-only.
             *
             * [Api set: Comments 1.1]
             */
            created?: boolean;
            /**
             *
             * Returns a value that uniquely identifies the comment in a given document. Read-only.
             *
             * [Api set: Comments 1.1]
             */
            id?: boolean;
            /**
             *
             * Gets the level of the comment: 0 if it is a root comment, or 1 if it is a reply. Read-only.
             *
             * [Api set: Comments 1.1]
             */
            level?: boolean;
            /**
             *
             * Gets the comment's mentions.
             *
             * [Api set: Comments 1.1]
             */
            mentions?: boolean;
            /**
             *
             * Gets or sets whether this comment is resolved.
             *
             * [Api set: Comments 1.1]
             */
            resolved?: boolean;
            /**
             *
             * Gets or sets the comment's plain text, without formatting.
             *
             * [Api set: Comments 1.1]
             */
            text?: boolean;
        }
        /**
         *
         * Represents a collection of comments, either replies to a specific comment thread, or all comments in the document or part of the document.
         *
         * [Api set: Comments 1.1]
         */
        interface CommentCollectionLoadOptions {
            $all?: boolean;
            /**
            *
            * For EACH ITEM in the collection: Gets this comment's parent. If this is a root comment, throws.
            *
            * [Api set: Comments 1.1]
            */
            parent?: OfficeCore.Interfaces.CommentLoadOptions;
            /**
            *
            * For EACH ITEM in the collection: Gets this comment's parent. If this is a root comment, returns a null object.
            *
            * [Api set: Comments 1.1]
            */
            parentOrNullObject?: OfficeCore.Interfaces.CommentLoadOptions;
            /**
             *
             * For EACH ITEM in the collection: Gets an object representing the comment's author. Read-only.
             *
             * [Api set: Comments 1.1]
             */
            author?: boolean;
            /**
             *
             * For EACH ITEM in the collection: Gets when the comment was created. Read-only.
             *
             * [Api set: Comments 1.1]
             */
            created?: boolean;
            /**
             *
             * For EACH ITEM in the collection: Returns a value that uniquely identifies the comment in a given document. Read-only.
             *
             * [Api set: Comments 1.1]
             */
            id?: boolean;
            /**
             *
             * For EACH ITEM in the collection: Gets the level of the comment: 0 if it is a root comment, or 1 if it is a reply. Read-only.
             *
             * [Api set: Comments 1.1]
             */
            level?: boolean;
            /**
             *
             * For EACH ITEM in the collection: Gets the comment's mentions.
             *
             * [Api set: Comments 1.1]
             */
            mentions?: boolean;
            /**
             *
             * For EACH ITEM in the collection: Gets or sets whether this comment is resolved.
             *
             * [Api set: Comments 1.1]
             */
            resolved?: boolean;
            /**
             *
             * For EACH ITEM in the collection: Gets or sets the comment's plain text, without formatting.
             *
             * [Api set: Comments 1.1]
             */
            text?: boolean;
        }
    }
}
