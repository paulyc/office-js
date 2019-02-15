/*
	Copyright (c) Microsoft Corporation.  All rights reserved.
*/

/*
	Your use of this file is governed by the Microsoft Services Agreement http://go.microsoft.com/fwlink/?LinkId=266419.
*/

/*
* @overview es6-promise - a tiny implementation of Promises/A+.
* @copyright Copyright (c) 2014 Yehuda Katz, Tom Dale, Stefan Penner and contributors (Conversion to ES6 API by Jake Archibald)
* @license   Licensed under MIT license
*            See https://raw.githubusercontent.com/jakearchibald/es6-promise/master/LICENSE
* @version   2.3.0
*/


// Sources:
// osfweb: 16.0\11329.10000
// runtime: 16.0.11329.30001
// core: 16.0\11405.10000
// host: excel 16.0.11329.30001

var __extends=(this && this.__extends) || function (d, b) {
	for (var p in b)
		if (b.hasOwnProperty(p))
			d[p]=b[p];
	function __() { this.constructor=d; }
	d.prototype=b===null ? Object.create(b) : (__.prototype=b.prototype, new __());
};
var OfficeExt;
(function (OfficeExt) {
	var MicrosoftAjaxFactory=(function () {
		function MicrosoftAjaxFactory() {
		}
		MicrosoftAjaxFactory.prototype.isMsAjaxLoaded=function () {
			if (typeof (Sys) !=='undefined' && typeof (Type) !=='undefined' &&
				Sys.StringBuilder && typeof (Sys.StringBuilder)==="function" &&
				Type.registerNamespace && typeof (Type.registerNamespace)==="function" &&
				Type.registerClass && typeof (Type.registerClass)==="function" &&
				typeof (Function._validateParams)==="function" &&
				Sys.Serialization && Sys.Serialization.JavaScriptSerializer && typeof (Sys.Serialization.JavaScriptSerializer.serialize)==="function") {
				return true;
			}
			else {
				return false;
			}
		};
		MicrosoftAjaxFactory.prototype.loadMsAjaxFull=function (callback) {
			var msAjaxCDNPath=(window.location.protocol.toLowerCase()==='https:' ? 'https:' : 'http:')+'//ajax.aspnetcdn.com/ajax/3.5/MicrosoftAjax.js';
			OSF.OUtil.loadScript(msAjaxCDNPath, callback);
		};
		Object.defineProperty(MicrosoftAjaxFactory.prototype, "msAjaxError", {
			get: function () {
				if (this._msAjaxError==null && this.isMsAjaxLoaded()) {
					this._msAjaxError=Error;
				}
				return this._msAjaxError;
			},
			set: function (errorClass) {
				this._msAjaxError=errorClass;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(MicrosoftAjaxFactory.prototype, "msAjaxString", {
			get: function () {
				if (this._msAjaxString==null && this.isMsAjaxLoaded()) {
					this._msAjaxString=String;
				}
				return this._msAjaxString;
			},
			set: function (stringClass) {
				this._msAjaxString=stringClass;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(MicrosoftAjaxFactory.prototype, "msAjaxDebug", {
			get: function () {
				if (this._msAjaxDebug==null && this.isMsAjaxLoaded()) {
					this._msAjaxDebug=Sys.Debug;
				}
				return this._msAjaxDebug;
			},
			set: function (debugClass) {
				this._msAjaxDebug=debugClass;
			},
			enumerable: true,
			configurable: true
		});
		return MicrosoftAjaxFactory;
	})();
	OfficeExt.MicrosoftAjaxFactory=MicrosoftAjaxFactory;
})(OfficeExt || (OfficeExt={}));
var OsfMsAjaxFactory=new OfficeExt.MicrosoftAjaxFactory();
var OSF=OSF || {};
var OfficeExt;
(function (OfficeExt) {
	var SafeStorage=(function () {
		function SafeStorage(_internalStorage) {
			this._internalStorage=_internalStorage;
		}
		SafeStorage.prototype.getItem=function (key) {
			try {
				return this._internalStorage && this._internalStorage.getItem(key);
			}
			catch (e) {
				return null;
			}
		};
		SafeStorage.prototype.setItem=function (key, data) {
			try {
				this._internalStorage && this._internalStorage.setItem(key, data);
			}
			catch (e) {
			}
		};
		SafeStorage.prototype.clear=function () {
			try {
				this._internalStorage && this._internalStorage.clear();
			}
			catch (e) {
			}
		};
		SafeStorage.prototype.removeItem=function (key) {
			try {
				this._internalStorage && this._internalStorage.removeItem(key);
			}
			catch (e) {
			}
		};
		SafeStorage.prototype.getKeysWithPrefix=function (keyPrefix) {
			var keyList=[];
			try {
				var len=this._internalStorage && this._internalStorage.length || 0;
				for (var i=0; i < len; i++) {
					var key=this._internalStorage.key(i);
					if (key.indexOf(keyPrefix)===0) {
						keyList.push(key);
					}
				}
			}
			catch (e) {
			}
			return keyList;
		};
		return SafeStorage;
	})();
	OfficeExt.SafeStorage=SafeStorage;
})(OfficeExt || (OfficeExt={}));
OSF.XdmFieldName={
	ConversationUrl: "ConversationUrl",
	AppId: "AppId"
};
OSF.WindowNameItemKeys={
	BaseFrameName: "baseFrameName",
	HostInfo: "hostInfo",
	XdmInfo: "xdmInfo",
	SerializerVersion: "serializerVersion",
	AppContext: "appContext"
};
OSF.OUtil=(function () {
	var _uniqueId=-1;
	var _xdmInfoKey='&_xdm_Info=';
	var _serializerVersionKey='&_serializer_version=';
	var _xdmSessionKeyPrefix='_xdm_';
	var _serializerVersionKeyPrefix='_serializer_version=';
	var _fragmentSeparator='#';
	var _fragmentInfoDelimiter='&';
	var _classN="class";
	var _loadedScripts={};
	var _defaultScriptLoadingTimeout=30000;
	var _safeSessionStorage=null;
	var _safeLocalStorage=null;
	var _rndentropy=new Date().getTime();
	function _random() {
		var nextrand=0x7fffffff * (Math.random());
		nextrand ^=_rndentropy ^ ((new Date().getMilliseconds()) << Math.floor(Math.random() * (31 - 10)));
		return nextrand.toString(16);
	}
	;
	function _getSessionStorage() {
		if (!_safeSessionStorage) {
			try {
				var sessionStorage=window.sessionStorage;
			}
			catch (ex) {
				sessionStorage=null;
			}
			_safeSessionStorage=new OfficeExt.SafeStorage(sessionStorage);
		}
		return _safeSessionStorage;
	}
	;
	function _reOrderTabbableElements(elements) {
		var bucket0=[];
		var bucketPositive=[];
		var i;
		var len=elements.length;
		var ele;
		for (i=0; i < len; i++) {
			ele=elements[i];
			if (ele.tabIndex) {
				if (ele.tabIndex > 0) {
					bucketPositive.push(ele);
				}
				else if (ele.tabIndex===0) {
					bucket0.push(ele);
				}
			}
			else {
				bucket0.push(ele);
			}
		}
		bucketPositive=bucketPositive.sort(function (left, right) {
			var diff=left.tabIndex - right.tabIndex;
			if (diff===0) {
				diff=bucketPositive.indexOf(left) - bucketPositive.indexOf(right);
			}
			return diff;
		});
		return [].concat(bucketPositive, bucket0);
	}
	;
	return {
		set_entropy: function OSF_OUtil$set_entropy(entropy) {
			if (typeof entropy=="string") {
				for (var i=0; i < entropy.length; i+=4) {
					var temp=0;
					for (var j=0; j < 4 && i+j < entropy.length; j++) {
						temp=(temp << 8)+entropy.charCodeAt(i+j);
					}
					_rndentropy ^=temp;
				}
			}
			else if (typeof entropy=="number") {
				_rndentropy ^=entropy;
			}
			else {
				_rndentropy ^=0x7fffffff * Math.random();
			}
			_rndentropy &=0x7fffffff;
		},
		extend: function OSF_OUtil$extend(child, parent) {
			var F=function () { };
			F.prototype=parent.prototype;
			child.prototype=new F();
			child.prototype.constructor=child;
			child.uber=parent.prototype;
			if (parent.prototype.constructor===Object.prototype.constructor) {
				parent.prototype.constructor=parent;
			}
		},
		setNamespace: function OSF_OUtil$setNamespace(name, parent) {
			if (parent && name && !parent[name]) {
				parent[name]={};
			}
		},
		unsetNamespace: function OSF_OUtil$unsetNamespace(name, parent) {
			if (parent && name && parent[name]) {
				delete parent[name];
			}
		},
		serializeSettings: function OSF_OUtil$serializeSettings(settingsCollection) {
			var ret={};
			for (var key in settingsCollection) {
				var value=settingsCollection[key];
				try {
					if (JSON) {
						value=JSON.stringify(value, function dateReplacer(k, v) {
							return OSF.OUtil.isDate(this[k]) ? OSF.DDA.SettingsManager.DateJSONPrefix+this[k].getTime()+OSF.DDA.SettingsManager.DataJSONSuffix : v;
						});
					}
					else {
						value=Sys.Serialization.JavaScriptSerializer.serialize(value);
					}
					ret[key]=value;
				}
				catch (ex) {
				}
			}
			return ret;
		},
		deserializeSettings: function OSF_OUtil$deserializeSettings(serializedSettings) {
			var ret={};
			serializedSettings=serializedSettings || {};
			for (var key in serializedSettings) {
				var value=serializedSettings[key];
				try {
					if (JSON) {
						value=JSON.parse(value, function dateReviver(k, v) {
							var d;
							if (typeof v==='string' && v && v.length > 6 && v.slice(0, 5)===OSF.DDA.SettingsManager.DateJSONPrefix && v.slice(-1)===OSF.DDA.SettingsManager.DataJSONSuffix) {
								d=new Date(parseInt(v.slice(5, -1)));
								if (d) {
									return d;
								}
							}
							return v;
						});
					}
					else {
						value=Sys.Serialization.JavaScriptSerializer.deserialize(value, true);
					}
					ret[key]=value;
				}
				catch (ex) {
				}
			}
			return ret;
		},
		loadScript: function OSF_OUtil$loadScript(url, callback, timeoutInMs) {
			if (url && callback) {
				var doc=window.document;
				var _loadedScriptEntry=_loadedScripts[url];
				if (!_loadedScriptEntry) {
					var script=doc.createElement("script");
					script.type="text/javascript";
					_loadedScriptEntry={ loaded: false, pendingCallbacks: [callback], timer: null };
					_loadedScripts[url]=_loadedScriptEntry;
					var onLoadCallback=function OSF_OUtil_loadScript$onLoadCallback() {
						if (_loadedScriptEntry.timer !=null) {
							clearTimeout(_loadedScriptEntry.timer);
							delete _loadedScriptEntry.timer;
						}
						_loadedScriptEntry.loaded=true;
						var pendingCallbackCount=_loadedScriptEntry.pendingCallbacks.length;
						for (var i=0; i < pendingCallbackCount; i++) {
							var currentCallback=_loadedScriptEntry.pendingCallbacks.shift();
							currentCallback();
						}
					};
					var onLoadError=function OSF_OUtil_loadScript$onLoadError() {
						delete _loadedScripts[url];
						if (_loadedScriptEntry.timer !=null) {
							clearTimeout(_loadedScriptEntry.timer);
							delete _loadedScriptEntry.timer;
						}
						var pendingCallbackCount=_loadedScriptEntry.pendingCallbacks.length;
						for (var i=0; i < pendingCallbackCount; i++) {
							var currentCallback=_loadedScriptEntry.pendingCallbacks.shift();
							currentCallback();
						}
					};
					if (script.readyState) {
						script.onreadystatechange=function () {
							if (script.readyState=="loaded" || script.readyState=="complete") {
								script.onreadystatechange=null;
								onLoadCallback();
							}
						};
					}
					else {
						script.onload=onLoadCallback;
					}
					script.onerror=onLoadError;
					timeoutInMs=timeoutInMs || _defaultScriptLoadingTimeout;
					_loadedScriptEntry.timer=setTimeout(onLoadError, timeoutInMs);
					script.setAttribute("crossOrigin", "anonymous");
					script.src=url;
					doc.getElementsByTagName("head")[0].appendChild(script);
				}
				else if (_loadedScriptEntry.loaded) {
					callback();
				}
				else {
					_loadedScriptEntry.pendingCallbacks.push(callback);
				}
			}
		},
		loadCSS: function OSF_OUtil$loadCSS(url) {
			if (url) {
				var doc=window.document;
				var link=doc.createElement("link");
				link.type="text/css";
				link.rel="stylesheet";
				link.href=url;
				doc.getElementsByTagName("head")[0].appendChild(link);
			}
		},
		parseEnum: function OSF_OUtil$parseEnum(str, enumObject) {
			var parsed=enumObject[str.trim()];
			if (typeof (parsed)=='undefined') {
				OsfMsAjaxFactory.msAjaxDebug.trace("invalid enumeration string:"+str);
				throw OsfMsAjaxFactory.msAjaxError.argument("str");
			}
			return parsed;
		},
		delayExecutionAndCache: function OSF_OUtil$delayExecutionAndCache() {
			var obj={ calc: arguments[0] };
			return function () {
				if (obj.calc) {
					obj.val=obj.calc.apply(this, arguments);
					delete obj.calc;
				}
				return obj.val;
			};
		},
		getUniqueId: function OSF_OUtil$getUniqueId() {
			_uniqueId=_uniqueId+1;
			return _uniqueId.toString();
		},
		formatString: function OSF_OUtil$formatString() {
			var args=arguments;
			var source=args[0];
			return source.replace(/{(\d+)}/gm, function (match, number) {
				var index=parseInt(number, 10)+1;
				return args[index]===undefined ? '{'+number+'}' : args[index];
			});
		},
		generateConversationId: function OSF_OUtil$generateConversationId() {
			return [_random(), _random(), (new Date()).getTime().toString()].join('_');
		},
		getFrameName: function OSF_OUtil$getFrameName(cacheKey) {
			return _xdmSessionKeyPrefix+cacheKey+this.generateConversationId();
		},
		addXdmInfoAsHash: function OSF_OUtil$addXdmInfoAsHash(url, xdmInfoValue) {
			return OSF.OUtil.addInfoAsHash(url, _xdmInfoKey, xdmInfoValue, false);
		},
		addSerializerVersionAsHash: function OSF_OUtil$addSerializerVersionAsHash(url, serializerVersion) {
			return OSF.OUtil.addInfoAsHash(url, _serializerVersionKey, serializerVersion, true);
		},
		addInfoAsHash: function OSF_OUtil$addInfoAsHash(url, keyName, infoValue, encodeInfo) {
			url=url.trim() || '';
			var urlParts=url.split(_fragmentSeparator);
			var urlWithoutFragment=urlParts.shift();
			var fragment=urlParts.join(_fragmentSeparator);
			var newFragment;
			if (encodeInfo) {
				newFragment=[keyName, encodeURIComponent(infoValue), fragment].join('');
			}
			else {
				newFragment=[fragment, keyName, infoValue].join('');
			}
			return [urlWithoutFragment, _fragmentSeparator, newFragment].join('');
		},
		parseHostInfoFromWindowName: function OSF_OUtil$parseHostInfoFromWindowName(skipSessionStorage, windowName) {
			return OSF.OUtil.parseInfoFromWindowName(skipSessionStorage, windowName, OSF.WindowNameItemKeys.HostInfo);
		},
		parseXdmInfo: function OSF_OUtil$parseXdmInfo(skipSessionStorage) {
			var xdmInfoValue=OSF.OUtil.parseXdmInfoWithGivenFragment(skipSessionStorage, window.location.hash);
			if (!xdmInfoValue) {
				xdmInfoValue=OSF.OUtil.parseXdmInfoFromWindowName(skipSessionStorage, window.name);
			}
			return xdmInfoValue;
		},
		parseXdmInfoFromWindowName: function OSF_OUtil$parseXdmInfoFromWindowName(skipSessionStorage, windowName) {
			return OSF.OUtil.parseInfoFromWindowName(skipSessionStorage, windowName, OSF.WindowNameItemKeys.XdmInfo);
		},
		parseXdmInfoWithGivenFragment: function OSF_OUtil$parseXdmInfoWithGivenFragment(skipSessionStorage, fragment) {
			return OSF.OUtil.parseInfoWithGivenFragment(_xdmInfoKey, _xdmSessionKeyPrefix, false, skipSessionStorage, fragment);
		},
		parseSerializerVersion: function OSF_OUtil$parseSerializerVersion(skipSessionStorage) {
			var serializerVersion=OSF.OUtil.parseSerializerVersionWithGivenFragment(skipSessionStorage, window.location.hash);
			if (isNaN(serializerVersion)) {
				serializerVersion=OSF.OUtil.parseSerializerVersionFromWindowName(skipSessionStorage, window.name);
			}
			return serializerVersion;
		},
		parseSerializerVersionFromWindowName: function OSF_OUtil$parseSerializerVersionFromWindowName(skipSessionStorage, windowName) {
			return parseInt(OSF.OUtil.parseInfoFromWindowName(skipSessionStorage, windowName, OSF.WindowNameItemKeys.SerializerVersion));
		},
		parseSerializerVersionWithGivenFragment: function OSF_OUtil$parseSerializerVersionWithGivenFragment(skipSessionStorage, fragment) {
			return parseInt(OSF.OUtil.parseInfoWithGivenFragment(_serializerVersionKey, _serializerVersionKeyPrefix, true, skipSessionStorage, fragment));
		},
		parseInfoFromWindowName: function OSF_OUtil$parseInfoFromWindowName(skipSessionStorage, windowName, infoKey) {
			try {
				var windowNameObj=JSON.parse(windowName);
				var infoValue=windowNameObj !=null ? windowNameObj[infoKey] : null;
				var osfSessionStorage=_getSessionStorage();
				if (!skipSessionStorage && osfSessionStorage && windowNameObj !=null) {
					var sessionKey=windowNameObj[OSF.WindowNameItemKeys.BaseFrameName]+infoKey;
					if (infoValue) {
						osfSessionStorage.setItem(sessionKey, infoValue);
					}
					else {
						infoValue=osfSessionStorage.getItem(sessionKey);
					}
				}
				return infoValue;
			}
			catch (Exception) {
				return null;
			}
		},
		parseInfoWithGivenFragment: function OSF_OUtil$parseInfoWithGivenFragment(infoKey, infoKeyPrefix, decodeInfo, skipSessionStorage, fragment) {
			var fragmentParts=fragment.split(infoKey);
			var infoValue=fragmentParts.length > 1 ? fragmentParts[fragmentParts.length - 1] : null;
			if (decodeInfo && infoValue !=null) {
				if (infoValue.indexOf(_fragmentInfoDelimiter) >=0) {
					infoValue=infoValue.split(_fragmentInfoDelimiter)[0];
				}
				infoValue=decodeURIComponent(infoValue);
			}
			var osfSessionStorage=_getSessionStorage();
			if (!skipSessionStorage && osfSessionStorage) {
				var sessionKeyStart=window.name.indexOf(infoKeyPrefix);
				if (sessionKeyStart > -1) {
					var sessionKeyEnd=window.name.indexOf(";", sessionKeyStart);
					if (sessionKeyEnd==-1) {
						sessionKeyEnd=window.name.length;
					}
					var sessionKey=window.name.substring(sessionKeyStart, sessionKeyEnd);
					if (infoValue) {
						osfSessionStorage.setItem(sessionKey, infoValue);
					}
					else {
						infoValue=osfSessionStorage.getItem(sessionKey);
					}
				}
			}
			return infoValue;
		},
		getConversationId: function OSF_OUtil$getConversationId() {
			var searchString=window.location.search;
			var conversationId=null;
			if (searchString) {
				var index=searchString.indexOf("&");
				conversationId=index > 0 ? searchString.substring(1, index) : searchString.substr(1);
				if (conversationId && conversationId.charAt(conversationId.length - 1)==='=') {
					conversationId=conversationId.substring(0, conversationId.length - 1);
					if (conversationId) {
						conversationId=decodeURIComponent(conversationId);
					}
				}
			}
			return conversationId;
		},
		getInfoItems: function OSF_OUtil$getInfoItems(strInfo) {
			var items=strInfo.split("$");
			if (typeof items[1]=="undefined") {
				items=strInfo.split("|");
			}
			if (typeof items[1]=="undefined") {
				items=strInfo.split("%7C");
			}
			return items;
		},
		getXdmFieldValue: function OSF_OUtil$getXdmFieldValue(xdmFieldName, skipSessionStorage) {
			var fieldValue='';
			var xdmInfoValue=OSF.OUtil.parseXdmInfo(skipSessionStorage);
			if (xdmInfoValue) {
				var items=OSF.OUtil.getInfoItems(xdmInfoValue);
				if (items !=undefined && items.length >=3) {
					switch (xdmFieldName) {
						case OSF.XdmFieldName.ConversationUrl:
							fieldValue=items[2];
							break;
						case OSF.XdmFieldName.AppId:
							fieldValue=items[1];
							break;
					}
				}
			}
			return fieldValue;
		},
		validateParamObject: function OSF_OUtil$validateParamObject(params, expectedProperties, callback) {
			var e=Function._validateParams(arguments, [{ name: "params", type: Object, mayBeNull: false },
				{ name: "expectedProperties", type: Object, mayBeNull: false },
				{ name: "callback", type: Function, mayBeNull: true }
			]);
			if (e)
				throw e;
			for (var p in expectedProperties) {
				e=Function._validateParameter(params[p], expectedProperties[p], p);
				if (e)
					throw e;
			}
		},
		writeProfilerMark: function OSF_OUtil$writeProfilerMark(text) {
			if (window.msWriteProfilerMark) {
				window.msWriteProfilerMark(text);
				OsfMsAjaxFactory.msAjaxDebug.trace(text);
			}
		},
		outputDebug: function OSF_OUtil$outputDebug(text) {
			if (typeof (OsfMsAjaxFactory) !=='undefined' && OsfMsAjaxFactory.msAjaxDebug && OsfMsAjaxFactory.msAjaxDebug.trace) {
				OsfMsAjaxFactory.msAjaxDebug.trace(text);
			}
		},
		defineNondefaultProperty: function OSF_OUtil$defineNondefaultProperty(obj, prop, descriptor, attributes) {
			descriptor=descriptor || {};
			for (var nd in attributes) {
				var attribute=attributes[nd];
				if (descriptor[attribute]==undefined) {
					descriptor[attribute]=true;
				}
			}
			Object.defineProperty(obj, prop, descriptor);
			return obj;
		},
		defineNondefaultProperties: function OSF_OUtil$defineNondefaultProperties(obj, descriptors, attributes) {
			descriptors=descriptors || {};
			for (var prop in descriptors) {
				OSF.OUtil.defineNondefaultProperty(obj, prop, descriptors[prop], attributes);
			}
			return obj;
		},
		defineEnumerableProperty: function OSF_OUtil$defineEnumerableProperty(obj, prop, descriptor) {
			return OSF.OUtil.defineNondefaultProperty(obj, prop, descriptor, ["enumerable"]);
		},
		defineEnumerableProperties: function OSF_OUtil$defineEnumerableProperties(obj, descriptors) {
			return OSF.OUtil.defineNondefaultProperties(obj, descriptors, ["enumerable"]);
		},
		defineMutableProperty: function OSF_OUtil$defineMutableProperty(obj, prop, descriptor) {
			return OSF.OUtil.defineNondefaultProperty(obj, prop, descriptor, ["writable", "enumerable", "configurable"]);
		},
		defineMutableProperties: function OSF_OUtil$defineMutableProperties(obj, descriptors) {
			return OSF.OUtil.defineNondefaultProperties(obj, descriptors, ["writable", "enumerable", "configurable"]);
		},
		finalizeProperties: function OSF_OUtil$finalizeProperties(obj, descriptor) {
			descriptor=descriptor || {};
			var props=Object.getOwnPropertyNames(obj);
			var propsLength=props.length;
			for (var i=0; i < propsLength; i++) {
				var prop=props[i];
				var desc=Object.getOwnPropertyDescriptor(obj, prop);
				if (!desc.get && !desc.set) {
					desc.writable=descriptor.writable || false;
				}
				desc.configurable=descriptor.configurable || false;
				desc.enumerable=descriptor.enumerable || true;
				Object.defineProperty(obj, prop, desc);
			}
			return obj;
		},
		mapList: function OSF_OUtil$MapList(list, mapFunction) {
			var ret=[];
			if (list) {
				for (var item in list) {
					ret.push(mapFunction(list[item]));
				}
			}
			return ret;
		},
		listContainsKey: function OSF_OUtil$listContainsKey(list, key) {
			for (var item in list) {
				if (key==item) {
					return true;
				}
			}
			return false;
		},
		listContainsValue: function OSF_OUtil$listContainsElement(list, value) {
			for (var item in list) {
				if (value==list[item]) {
					return true;
				}
			}
			return false;
		},
		augmentList: function OSF_OUtil$augmentList(list, addenda) {
			var add=list.push ? function (key, value) { list.push(value); } : function (key, value) { list[key]=value; };
			for (var key in addenda) {
				add(key, addenda[key]);
			}
		},
		redefineList: function OSF_Outil$redefineList(oldList, newList) {
			for (var key1 in oldList) {
				delete oldList[key1];
			}
			for (var key2 in newList) {
				oldList[key2]=newList[key2];
			}
		},
		isArray: function OSF_OUtil$isArray(obj) {
			return Object.prototype.toString.apply(obj)==="[object Array]";
		},
		isFunction: function OSF_OUtil$isFunction(obj) {
			return Object.prototype.toString.apply(obj)==="[object Function]";
		},
		isDate: function OSF_OUtil$isDate(obj) {
			return Object.prototype.toString.apply(obj)==="[object Date]";
		},
		addEventListener: function OSF_OUtil$addEventListener(element, eventName, listener) {
			if (element.addEventListener) {
				element.addEventListener(eventName, listener, false);
			}
			else if ((Sys.Browser.agent===Sys.Browser.InternetExplorer) && element.attachEvent) {
				element.attachEvent("on"+eventName, listener);
			}
			else {
				element["on"+eventName]=listener;
			}
		},
		removeEventListener: function OSF_OUtil$removeEventListener(element, eventName, listener) {
			if (element.removeEventListener) {
				element.removeEventListener(eventName, listener, false);
			}
			else if ((Sys.Browser.agent===Sys.Browser.InternetExplorer) && element.detachEvent) {
				element.detachEvent("on"+eventName, listener);
			}
			else {
				element["on"+eventName]=null;
			}
		},
		getCookieValue: function OSF_OUtil$getCookieValue(cookieName) {
			var tmpCookieString=RegExp(cookieName+"[^;]+").exec(document.cookie);
			return tmpCookieString.toString().replace(/^[^=]+./, "");
		},
		xhrGet: function OSF_OUtil$xhrGet(url, onSuccess, onError) {
			var xmlhttp;
			try {
				xmlhttp=new XMLHttpRequest();
				xmlhttp.onreadystatechange=function () {
					if (xmlhttp.readyState==4) {
						if (xmlhttp.status==200) {
							onSuccess(xmlhttp.responseText);
						}
						else {
							onError(xmlhttp.status);
						}
					}
				};
				xmlhttp.open("GET", url, true);
				xmlhttp.send();
			}
			catch (ex) {
				onError(ex);
			}
		},
		xhrGetFull: function OSF_OUtil$xhrGetFull(url, oneDriveFileName, onSuccess, onError) {
			var xmlhttp;
			var requestedFileName=oneDriveFileName;
			try {
				xmlhttp=new XMLHttpRequest();
				xmlhttp.onreadystatechange=function () {
					if (xmlhttp.readyState==4) {
						if (xmlhttp.status==200) {
							onSuccess(xmlhttp, requestedFileName);
						}
						else {
							onError(xmlhttp.status);
						}
					}
				};
				xmlhttp.open("GET", url, true);
				xmlhttp.send();
			}
			catch (ex) {
				onError(ex);
			}
		},
		encodeBase64: function OSF_Outil$encodeBase64(input) {
			if (!input)
				return input;
			var codex="ABCDEFGHIJKLMNOP"+"QRSTUVWXYZabcdef"+"ghijklmnopqrstuv"+"wxyz0123456789+/=";
			var output=[];
			var temp=[];
			var index=0;
			var c1, c2, c3, a, b, c;
			var i;
			var length=input.length;
			do {
				c1=input.charCodeAt(index++);
				c2=input.charCodeAt(index++);
				c3=input.charCodeAt(index++);
				i=0;
				a=c1 & 255;
				b=c1 >> 8;
				c=c2 & 255;
				temp[i++]=a >> 2;
				temp[i++]=((a & 3) << 4) | (b >> 4);
				temp[i++]=((b & 15) << 2) | (c >> 6);
				temp[i++]=c & 63;
				if (!isNaN(c2)) {
					a=c2 >> 8;
					b=c3 & 255;
					c=c3 >> 8;
					temp[i++]=a >> 2;
					temp[i++]=((a & 3) << 4) | (b >> 4);
					temp[i++]=((b & 15) << 2) | (c >> 6);
					temp[i++]=c & 63;
				}
				if (isNaN(c2)) {
					temp[i - 1]=64;
				}
				else if (isNaN(c3)) {
					temp[i - 2]=64;
					temp[i - 1]=64;
				}
				for (var t=0; t < i; t++) {
					output.push(codex.charAt(temp[t]));
				}
			} while (index < length);
			return output.join("");
		},
		getSessionStorage: function OSF_Outil$getSessionStorage() {
			return _getSessionStorage();
		},
		getLocalStorage: function OSF_Outil$getLocalStorage() {
			if (!_safeLocalStorage) {
				try {
					var localStorage=window.localStorage;
				}
				catch (ex) {
					localStorage=null;
				}
				_safeLocalStorage=new OfficeExt.SafeStorage(localStorage);
			}
			return _safeLocalStorage;
		},
		convertIntToCssHexColor: function OSF_Outil$convertIntToCssHexColor(val) {
			var hex="#"+(Number(val)+0x1000000).toString(16).slice(-6);
			return hex;
		},
		attachClickHandler: function OSF_Outil$attachClickHandler(element, handler) {
			element.onclick=function (e) {
				handler();
			};
			element.ontouchend=function (e) {
				handler();
				e.preventDefault();
			};
		},
		getQueryStringParamValue: function OSF_Outil$getQueryStringParamValue(queryString, paramName) {
			var e=Function._validateParams(arguments, [{ name: "queryString", type: String, mayBeNull: false },
				{ name: "paramName", type: String, mayBeNull: false }
			]);
			if (e) {
				OsfMsAjaxFactory.msAjaxDebug.trace("OSF_Outil_getQueryStringParamValue: Parameters cannot be null.");
				return "";
			}
			var queryExp=new RegExp("[\\?&]"+paramName+"=([^&#]*)", "i");
			if (!queryExp.test(queryString)) {
				OsfMsAjaxFactory.msAjaxDebug.trace("OSF_Outil_getQueryStringParamValue: The parameter is not found.");
				return "";
			}
			return queryExp.exec(queryString)[1];
		},
		isiOS: function OSF_Outil$isiOS() {
			return (window.navigator.userAgent.match(/(iPad|iPhone|iPod)/g) ? true : false);
		},
		isChrome: function OSF_Outil$isChrome() {
			return (window.navigator.userAgent.indexOf("Chrome") > 0) && !OSF.OUtil.isEdge();
		},
		isEdge: function OSF_Outil$isEdge() {
			return window.navigator.userAgent.indexOf("Edge") > 0;
		},
		isIE: function OSF_Outil$isIE() {
			return window.navigator.userAgent.indexOf("Trident") > 0;
		},
		isFirefox: function OSF_Outil$isFirefox() {
			return window.navigator.userAgent.indexOf("Firefox") > 0;
		},
		shallowCopy: function OSF_Outil$shallowCopy(sourceObj) {
			if (sourceObj==null) {
				return null;
			}
			else if (!(sourceObj instanceof Object)) {
				return sourceObj;
			}
			else if (Array.isArray(sourceObj)) {
				var copyArr=[];
				for (var i=0; i < sourceObj.length; i++) {
					copyArr.push(sourceObj[i]);
				}
				return copyArr;
			}
			else {
				var copyObj=sourceObj.constructor();
				for (var property in sourceObj) {
					if (sourceObj.hasOwnProperty(property)) {
						copyObj[property]=sourceObj[property];
					}
				}
				return copyObj;
			}
		},
		createObject: function OSF_Outil$createObject(properties) {
			var obj=null;
			if (properties) {
				obj={};
				var len=properties.length;
				for (var i=0; i < len; i++) {
					obj[properties[i].name]=properties[i].value;
				}
			}
			return obj;
		},
		addClass: function OSF_OUtil$addClass(elmt, val) {
			if (!OSF.OUtil.hasClass(elmt, val)) {
				var className=elmt.getAttribute(_classN);
				if (className) {
					elmt.setAttribute(_classN, className+" "+val);
				}
				else {
					elmt.setAttribute(_classN, val);
				}
			}
		},
		removeClass: function OSF_OUtil$removeClass(elmt, val) {
			if (OSF.OUtil.hasClass(elmt, val)) {
				var className=elmt.getAttribute(_classN);
				var reg=new RegExp('(\\s|^)'+val+'(\\s|$)');
				className=className.replace(reg, '');
				elmt.setAttribute(_classN, className);
			}
		},
		hasClass: function OSF_OUtil$hasClass(elmt, clsName) {
			var className=elmt.getAttribute(_classN);
			return className && className.match(new RegExp('(\\s|^)'+clsName+'(\\s|$)'));
		},
		focusToFirstTabbable: function OSF_OUtil$focusToFirstTabbable(all, backward) {
			var next;
			var focused=false;
			var candidate;
			var setFlag=function (e) {
				focused=true;
			};
			var findNextPos=function (allLen, currPos, backward) {
				if (currPos < 0 || currPos > allLen) {
					return -1;
				}
				else if (currPos===0 && backward) {
					return -1;
				}
				else if (currPos===allLen - 1 && !backward) {
					return -1;
				}
				if (backward) {
					return currPos - 1;
				}
				else {
					return currPos+1;
				}
			};
			all=_reOrderTabbableElements(all);
			next=backward ? all.length - 1 : 0;
			if (all.length===0) {
				return null;
			}
			while (!focused && next >=0 && next < all.length) {
				candidate=all[next];
				window.focus();
				candidate.addEventListener('focus', setFlag);
				candidate.focus();
				candidate.removeEventListener('focus', setFlag);
				next=findNextPos(all.length, next, backward);
				if (!focused && candidate===document.activeElement) {
					focused=true;
				}
			}
			if (focused) {
				return candidate;
			}
			else {
				return null;
			}
		},
		focusToNextTabbable: function OSF_OUtil$focusToNextTabbable(all, curr, shift) {
			var currPos;
			var next;
			var focused=false;
			var candidate;
			var setFlag=function (e) {
				focused=true;
			};
			var findCurrPos=function (all, curr) {
				var i=0;
				for (; i < all.length; i++) {
					if (all[i]===curr) {
						return i;
					}
				}
				return -1;
			};
			var findNextPos=function (allLen, currPos, shift) {
				if (currPos < 0 || currPos > allLen) {
					return -1;
				}
				else if (currPos===0 && shift) {
					return -1;
				}
				else if (currPos===allLen - 1 && !shift) {
					return -1;
				}
				if (shift) {
					return currPos - 1;
				}
				else {
					return currPos+1;
				}
			};
			all=_reOrderTabbableElements(all);
			currPos=findCurrPos(all, curr);
			next=findNextPos(all.length, currPos, shift);
			if (next < 0) {
				return null;
			}
			while (!focused && next >=0 && next < all.length) {
				candidate=all[next];
				candidate.addEventListener('focus', setFlag);
				candidate.focus();
				candidate.removeEventListener('focus', setFlag);
				next=findNextPos(all.length, next, shift);
				if (!focused && candidate===document.activeElement) {
					focused=true;
				}
			}
			if (focused) {
				return candidate;
			}
			else {
				return null;
			}
		}
	};
})();
OSF.OUtil.Guid=(function () {
	var hexCode=["0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "a", "b", "c", "d", "e", "f"];
	return {
		generateNewGuid: function OSF_Outil_Guid$generateNewGuid() {
			var result="";
			var tick=(new Date()).getTime();
			var index=0;
			for (; index < 32 && tick > 0; index++) {
				if (index==8 || index==12 || index==16 || index==20) {
					result+="-";
				}
				result+=hexCode[tick % 16];
				tick=Math.floor(tick / 16);
			}
			for (; index < 32; index++) {
				if (index==8 || index==12 || index==16 || index==20) {
					result+="-";
				}
				result+=hexCode[Math.floor(Math.random() * 16)];
			}
			return result;
		}
	};
})();
window.OSF=OSF;
OSF.OUtil.setNamespace("OSF", window);
OSF.MessageIDs={
	"FetchBundleUrl": 0,
	"LoadReactBundle": 1,
	"LoadBundleSuccess": 2,
	"LoadBundleError": 3
};
OSF.AppName={
	Unsupported: 0,
	Excel: 1,
	Word: 2,
	PowerPoint: 4,
	Outlook: 8,
	ExcelWebApp: 16,
	WordWebApp: 32,
	OutlookWebApp: 64,
	Project: 128,
	AccessWebApp: 256,
	PowerpointWebApp: 512,
	ExcelIOS: 1024,
	Sway: 2048,
	WordIOS: 4096,
	PowerPointIOS: 8192,
	Access: 16384,
	Lync: 32768,
	OutlookIOS: 65536,
	OneNoteWebApp: 131072,
	OneNote: 262144,
	ExcelWinRT: 524288,
	WordWinRT: 1048576,
	PowerpointWinRT: 2097152,
	OutlookAndroid: 4194304,
	OneNoteWinRT: 8388608,
	ExcelAndroid: 8388609,
	VisioWebApp: 8388610,
	OneNoteIOS: 8388611,
	WordAndroid: 8388613,
	PowerpointAndroid: 8388614,
	Visio: 8388615,
	OneNoteAndroid: 4194305
};
OSF.InternalPerfMarker={
	DataCoercionBegin: "Agave.HostCall.CoerceDataStart",
	DataCoercionEnd: "Agave.HostCall.CoerceDataEnd"
};
OSF.HostCallPerfMarker={
	IssueCall: "Agave.HostCall.IssueCall",
	ReceiveResponse: "Agave.HostCall.ReceiveResponse",
	RuntimeExceptionRaised: "Agave.HostCall.RuntimeExecptionRaised"
};
OSF.AgaveHostAction={
	"Select": 0,
	"UnSelect": 1,
	"CancelDialog": 2,
	"InsertAgave": 3,
	"CtrlF6In": 4,
	"CtrlF6Exit": 5,
	"CtrlF6ExitShift": 6,
	"SelectWithError": 7,
	"NotifyHostError": 8,
	"RefreshAddinCommands": 9,
	"PageIsReady": 10,
	"TabIn": 11,
	"TabInShift": 12,
	"TabExit": 13,
	"TabExitShift": 14,
	"EscExit": 15,
	"F2Exit": 16,
	"ExitNoFocusable": 17,
	"ExitNoFocusableShift": 18,
	"MouseEnter": 19,
	"MouseLeave": 20,
	"UpdateTargetUrl": 21,
	"InstallCustomFunctions": 22,
	"SendTelemetryEvent": 23,
	"UninstallCustomFunctions": 24
};
OSF.SharedConstants={
	"NotificationConversationIdSuffix": '_ntf'
};
OSF.DialogMessageType={
	DialogMessageReceived: 0,
	DialogParentMessageReceived: 1,
	DialogClosed: 12006
};
OSF.OfficeAppContext=function OSF_OfficeAppContext(id, appName, appVersion, appUILocale, dataLocale, docUrl, clientMode, settings, reason, osfControlType, eToken, correlationId, appInstanceId, touchEnabled, commerceAllowed, appMinorVersion, requirementMatrix, hostCustomMessage, hostFullVersion, clientWindowHeight, clientWindowWidth, addinName, appDomains, dialogRequirementMatrix) {
	this._id=id;
	this._appName=appName;
	this._appVersion=appVersion;
	this._appUILocale=appUILocale;
	this._dataLocale=dataLocale;
	this._docUrl=docUrl;
	this._clientMode=clientMode;
	this._settings=settings;
	this._reason=reason;
	this._osfControlType=osfControlType;
	this._eToken=eToken;
	this._correlationId=correlationId;
	this._appInstanceId=appInstanceId;
	this._touchEnabled=touchEnabled;
	this._commerceAllowed=commerceAllowed;
	this._appMinorVersion=appMinorVersion;
	this._requirementMatrix=requirementMatrix;
	this._hostCustomMessage=hostCustomMessage;
	this._hostFullVersion=hostFullVersion;
	this._isDialog=false;
	this._clientWindowHeight=clientWindowHeight;
	this._clientWindowWidth=clientWindowWidth;
	this._addinName=addinName;
	this._appDomains=appDomains;
	this._dialogRequirementMatrix=dialogRequirementMatrix;
	this.get_id=function get_id() { return this._id; };
	this.get_appName=function get_appName() { return this._appName; };
	this.get_appVersion=function get_appVersion() { return this._appVersion; };
	this.get_appUILocale=function get_appUILocale() { return this._appUILocale; };
	this.get_dataLocale=function get_dataLocale() { return this._dataLocale; };
	this.get_docUrl=function get_docUrl() { return this._docUrl; };
	this.get_clientMode=function get_clientMode() { return this._clientMode; };
	this.get_bindings=function get_bindings() { return this._bindings; };
	this.get_settings=function get_settings() { return this._settings; };
	this.get_reason=function get_reason() { return this._reason; };
	this.get_osfControlType=function get_osfControlType() { return this._osfControlType; };
	this.get_eToken=function get_eToken() { return this._eToken; };
	this.get_correlationId=function get_correlationId() { return this._correlationId; };
	this.get_appInstanceId=function get_appInstanceId() { return this._appInstanceId; };
	this.get_touchEnabled=function get_touchEnabled() { return this._touchEnabled; };
	this.get_commerceAllowed=function get_commerceAllowed() { return this._commerceAllowed; };
	this.get_appMinorVersion=function get_appMinorVersion() { return this._appMinorVersion; };
	this.get_requirementMatrix=function get_requirementMatrix() { return this._requirementMatrix; };
	this.get_dialogRequirementMatrix=function get_dialogRequirementMatrix() { return this._dialogRequirementMatrix; };
	this.get_hostCustomMessage=function get_hostCustomMessage() { return this._hostCustomMessage; };
	this.get_hostFullVersion=function get_hostFullVersion() { return this._hostFullVersion; };
	this.get_isDialog=function get_isDialog() { return this._isDialog; };
	this.get_clientWindowHeight=function get_clientWindowHeight() { return this._clientWindowHeight; };
	this.get_clientWindowWidth=function get_clientWindowWidth() { return this._clientWindowWidth; };
	this.get_addinName=function get_addinName() { return this._addinName; };
	this.get_appDomains=function get_appDomains() { return this._appDomains; };
};
OSF.OsfControlType={
	DocumentLevel: 0,
	ContainerLevel: 1
};
OSF.ClientMode={
	ReadOnly: 0,
	ReadWrite: 1
};
OSF.OUtil.setNamespace("Microsoft", window);
OSF.OUtil.setNamespace("Office", Microsoft);
OSF.OUtil.setNamespace("Client", Microsoft.Office);
OSF.OUtil.setNamespace("WebExtension", Microsoft.Office);
Microsoft.Office.WebExtension.InitializationReason={
	Inserted: "inserted",
	DocumentOpened: "documentOpened"
};
Microsoft.Office.WebExtension.ValueFormat={
	Unformatted: "unformatted",
	Formatted: "formatted"
};
Microsoft.Office.WebExtension.FilterType={
	All: "all"
};
Microsoft.Office.WebExtension.Parameters={
	BindingType: "bindingType",
	CoercionType: "coercionType",
	ValueFormat: "valueFormat",
	FilterType: "filterType",
	Columns: "columns",
	SampleData: "sampleData",
	GoToType: "goToType",
	SelectionMode: "selectionMode",
	Id: "id",
	PromptText: "promptText",
	ItemName: "itemName",
	FailOnCollision: "failOnCollision",
	StartRow: "startRow",
	StartColumn: "startColumn",
	RowCount: "rowCount",
	ColumnCount: "columnCount",
	Callback: "callback",
	AsyncContext: "asyncContext",
	Data: "data",
	Rows: "rows",
	OverwriteIfStale: "overwriteIfStale",
	FileType: "fileType",
	EventType: "eventType",
	Handler: "handler",
	SliceSize: "sliceSize",
	SliceIndex: "sliceIndex",
	ActiveView: "activeView",
	Status: "status",
	PlatformType: "platformType",
	HostType: "hostType",
	ForceConsent: "forceConsent",
	ForceAddAccount: "forceAddAccount",
	AuthChallenge: "authChallenge",
	Reserved: "reserved",
	Tcid: "tcid",
	Xml: "xml",
	Namespace: "namespace",
	Prefix: "prefix",
	XPath: "xPath",
	Text: "text",
	ImageLeft: "imageLeft",
	ImageTop: "imageTop",
	ImageWidth: "imageWidth",
	ImageHeight: "imageHeight",
	TaskId: "taskId",
	FieldId: "fieldId",
	FieldValue: "fieldValue",
	ServerUrl: "serverUrl",
	ListName: "listName",
	ResourceId: "resourceId",
	ViewType: "viewType",
	ViewName: "viewName",
	GetRawValue: "getRawValue",
	CellFormat: "cellFormat",
	TableOptions: "tableOptions",
	TaskIndex: "taskIndex",
	ResourceIndex: "resourceIndex",
	CustomFieldId: "customFieldId",
	Url: "url",
	MessageHandler: "messageHandler",
	Width: "width",
	Height: "height",
	RequireHTTPs: "requireHTTPS",
	MessageToParent: "messageToParent",
	DisplayInIframe: "displayInIframe",
	MessageContent: "messageContent",
	HideTitle: "hideTitle",
	UseDeviceIndependentPixels: "useDeviceIndependentPixels",
	PromptBeforeOpen: "promptBeforeOpen",
	AppCommandInvocationCompletedData: "appCommandInvocationCompletedData",
	Base64: "base64",
	FormId: "formId"
};
OSF.OUtil.setNamespace("DDA", OSF);
OSF.DDA.DocumentMode={
	ReadOnly: 1,
	ReadWrite: 0
};
OSF.DDA.PropertyDescriptors={
	AsyncResultStatus: "AsyncResultStatus"
};
OSF.DDA.EventDescriptors={};
OSF.DDA.ListDescriptors={};
OSF.DDA.UI={};
OSF.DDA.getXdmEventName=function OSF_DDA$GetXdmEventName(id, eventType) {
	if (eventType==Microsoft.Office.WebExtension.EventType.BindingSelectionChanged ||
		eventType==Microsoft.Office.WebExtension.EventType.BindingDataChanged ||
		eventType==Microsoft.Office.WebExtension.EventType.DataNodeDeleted ||
		eventType==Microsoft.Office.WebExtension.EventType.DataNodeInserted ||
		eventType==Microsoft.Office.WebExtension.EventType.DataNodeReplaced) {
		return id+"_"+eventType;
	}
	else {
		return eventType;
	}
};
OSF.DDA.MethodDispId={
	dispidMethodMin: 64,
	dispidGetSelectedDataMethod: 64,
	dispidSetSelectedDataMethod: 65,
	dispidAddBindingFromSelectionMethod: 66,
	dispidAddBindingFromPromptMethod: 67,
	dispidGetBindingMethod: 68,
	dispidReleaseBindingMethod: 69,
	dispidGetBindingDataMethod: 70,
	dispidSetBindingDataMethod: 71,
	dispidAddRowsMethod: 72,
	dispidClearAllRowsMethod: 73,
	dispidGetAllBindingsMethod: 74,
	dispidLoadSettingsMethod: 75,
	dispidSaveSettingsMethod: 76,
	dispidGetDocumentCopyMethod: 77,
	dispidAddBindingFromNamedItemMethod: 78,
	dispidAddColumnsMethod: 79,
	dispidGetDocumentCopyChunkMethod: 80,
	dispidReleaseDocumentCopyMethod: 81,
	dispidNavigateToMethod: 82,
	dispidGetActiveViewMethod: 83,
	dispidGetDocumentThemeMethod: 84,
	dispidGetOfficeThemeMethod: 85,
	dispidGetFilePropertiesMethod: 86,
	dispidClearFormatsMethod: 87,
	dispidSetTableOptionsMethod: 88,
	dispidSetFormatsMethod: 89,
	dispidExecuteRichApiRequestMethod: 93,
	dispidAppCommandInvocationCompletedMethod: 94,
	dispidCloseContainerMethod: 97,
	dispidGetAccessTokenMethod: 98,
	dispidOpenBrowserWindow: 102,
	dispidCreateDocumentMethod: 105,
	dispidInsertFormMethod: 106,
	dispidDisplayRibbonCalloutAsyncMethod: 109,
	dispidGetSelectedTaskMethod: 110,
	dispidGetSelectedResourceMethod: 111,
	dispidGetTaskMethod: 112,
	dispidGetResourceFieldMethod: 113,
	dispidGetWSSUrlMethod: 114,
	dispidGetTaskFieldMethod: 115,
	dispidGetProjectFieldMethod: 116,
	dispidGetSelectedViewMethod: 117,
	dispidGetTaskByIndexMethod: 118,
	dispidGetResourceByIndexMethod: 119,
	dispidSetTaskFieldMethod: 120,
	dispidSetResourceFieldMethod: 121,
	dispidGetMaxTaskIndexMethod: 122,
	dispidGetMaxResourceIndexMethod: 123,
	dispidCreateTaskMethod: 124,
	dispidAddDataPartMethod: 128,
	dispidGetDataPartByIdMethod: 129,
	dispidGetDataPartsByNamespaceMethod: 130,
	dispidGetDataPartXmlMethod: 131,
	dispidGetDataPartNodesMethod: 132,
	dispidDeleteDataPartMethod: 133,
	dispidGetDataNodeValueMethod: 134,
	dispidGetDataNodeXmlMethod: 135,
	dispidGetDataNodesMethod: 136,
	dispidSetDataNodeValueMethod: 137,
	dispidSetDataNodeXmlMethod: 138,
	dispidAddDataNamespaceMethod: 139,
	dispidGetDataUriByPrefixMethod: 140,
	dispidGetDataPrefixByUriMethod: 141,
	dispidGetDataNodeTextMethod: 142,
	dispidSetDataNodeTextMethod: 143,
	dispidMessageParentMethod: 144,
	dispidSendMessageMethod: 145,
	dispidExecuteFeature: 146,
	dispidQueryFeature: 147,
	dispidMethodMax: 147
};
OSF.DDA.EventDispId={
	dispidEventMin: 0,
	dispidInitializeEvent: 0,
	dispidSettingsChangedEvent: 1,
	dispidDocumentSelectionChangedEvent: 2,
	dispidBindingSelectionChangedEvent: 3,
	dispidBindingDataChangedEvent: 4,
	dispidDocumentOpenEvent: 5,
	dispidDocumentCloseEvent: 6,
	dispidActiveViewChangedEvent: 7,
	dispidDocumentThemeChangedEvent: 8,
	dispidOfficeThemeChangedEvent: 9,
	dispidDialogMessageReceivedEvent: 10,
	dispidDialogNotificationShownInAddinEvent: 11,
	dispidDialogParentMessageReceivedEvent: 12,
	dispidObjectDeletedEvent: 13,
	dispidObjectSelectionChangedEvent: 14,
	dispidObjectDataChangedEvent: 15,
	dispidContentControlAddedEvent: 16,
	dispidActivationStatusChangedEvent: 32,
	dispidRichApiMessageEvent: 33,
	dispidAppCommandInvokedEvent: 39,
	dispidOlkItemSelectedChangedEvent: 46,
	dispidOlkRecipientsChangedEvent: 47,
	dispidOlkAppointmentTimeChangedEvent: 48,
	dispidOlkRecurrenceChangedEvent: 49,
	dispidOlkAttachmentsChangedEvent: 50,
	dispidOlkEnhancedLocationsChangedEvent: 51,
	dispidOlkInfobarClickedEvent: 52,
	dispidTaskSelectionChangedEvent: 56,
	dispidResourceSelectionChangedEvent: 57,
	dispidViewSelectionChangedEvent: 58,
	dispidDataNodeAddedEvent: 60,
	dispidDataNodeReplacedEvent: 61,
	dispidDataNodeDeletedEvent: 62,
	dispidEventMax: 63
};
OSF.DDA.ErrorCodeManager=(function () {
	var _errorMappings={};
	return {
		getErrorArgs: function OSF_DDA_ErrorCodeManager$getErrorArgs(errorCode) {
			var errorArgs=_errorMappings[errorCode];
			if (!errorArgs) {
				errorArgs=_errorMappings[this.errorCodes.ooeInternalError];
			}
			else {
				if (!errorArgs.name) {
					errorArgs.name=_errorMappings[this.errorCodes.ooeInternalError].name;
				}
				if (!errorArgs.message) {
					errorArgs.message=_errorMappings[this.errorCodes.ooeInternalError].message;
				}
			}
			return errorArgs;
		},
		addErrorMessage: function OSF_DDA_ErrorCodeManager$addErrorMessage(errorCode, errorNameMessage) {
			_errorMappings[errorCode]=errorNameMessage;
		},
		errorCodes: {
			ooeSuccess: 0,
			ooeChunkResult: 1,
			ooeCoercionTypeNotSupported: 1000,
			ooeGetSelectionNotMatchDataType: 1001,
			ooeCoercionTypeNotMatchBinding: 1002,
			ooeInvalidGetRowColumnCounts: 1003,
			ooeSelectionNotSupportCoercionType: 1004,
			ooeInvalidGetStartRowColumn: 1005,
			ooeNonUniformPartialGetNotSupported: 1006,
			ooeGetDataIsTooLarge: 1008,
			ooeFileTypeNotSupported: 1009,
			ooeGetDataParametersConflict: 1010,
			ooeInvalidGetColumns: 1011,
			ooeInvalidGetRows: 1012,
			ooeInvalidReadForBlankRow: 1013,
			ooeUnsupportedDataObject: 2000,
			ooeCannotWriteToSelection: 2001,
			ooeDataNotMatchSelection: 2002,
			ooeOverwriteWorksheetData: 2003,
			ooeDataNotMatchBindingSize: 2004,
			ooeInvalidSetStartRowColumn: 2005,
			ooeInvalidDataFormat: 2006,
			ooeDataNotMatchCoercionType: 2007,
			ooeDataNotMatchBindingType: 2008,
			ooeSetDataIsTooLarge: 2009,
			ooeNonUniformPartialSetNotSupported: 2010,
			ooeInvalidSetColumns: 2011,
			ooeInvalidSetRows: 2012,
			ooeSetDataParametersConflict: 2013,
			ooeCellDataAmountBeyondLimits: 2014,
			ooeSelectionCannotBound: 3000,
			ooeBindingNotExist: 3002,
			ooeBindingToMultipleSelection: 3003,
			ooeInvalidSelectionForBindingType: 3004,
			ooeOperationNotSupportedOnThisBindingType: 3005,
			ooeNamedItemNotFound: 3006,
			ooeMultipleNamedItemFound: 3007,
			ooeInvalidNamedItemForBindingType: 3008,
			ooeUnknownBindingType: 3009,
			ooeOperationNotSupportedOnMatrixData: 3010,
			ooeInvalidColumnsForBinding: 3011,
			ooeSettingNameNotExist: 4000,
			ooeSettingsCannotSave: 4001,
			ooeSettingsAreStale: 4002,
			ooeOperationNotSupported: 5000,
			ooeInternalError: 5001,
			ooeDocumentReadOnly: 5002,
			ooeEventHandlerNotExist: 5003,
			ooeInvalidApiCallInContext: 5004,
			ooeShuttingDown: 5005,
			ooeUnsupportedEnumeration: 5007,
			ooeIndexOutOfRange: 5008,
			ooeBrowserAPINotSupported: 5009,
			ooeInvalidParam: 5010,
			ooeRequestTimeout: 5011,
			ooeInvalidOrTimedOutSession: 5012,
			ooeInvalidApiArguments: 5013,
			ooeOperationCancelled: 5014,
			ooeWorkbookHidden: 5015,
			ooeTooManyIncompleteRequests: 5100,
			ooeRequestTokenUnavailable: 5101,
			ooeActivityLimitReached: 5102,
			ooeCustomXmlNodeNotFound: 6000,
			ooeCustomXmlError: 6100,
			ooeCustomXmlExceedQuota: 6101,
			ooeCustomXmlOutOfDate: 6102,
			ooeNoCapability: 7000,
			ooeCannotNavTo: 7001,
			ooeSpecifiedIdNotExist: 7002,
			ooeNavOutOfBound: 7004,
			ooeElementMissing: 8000,
			ooeProtectedError: 8001,
			ooeInvalidCellsValue: 8010,
			ooeInvalidTableOptionValue: 8011,
			ooeInvalidFormatValue: 8012,
			ooeRowIndexOutOfRange: 8020,
			ooeColIndexOutOfRange: 8021,
			ooeFormatValueOutOfRange: 8022,
			ooeCellFormatAmountBeyondLimits: 8023,
			ooeMemoryFileLimit: 11000,
			ooeNetworkProblemRetrieveFile: 11001,
			ooeInvalidSliceSize: 11002,
			ooeInvalidCallback: 11101,
			ooeInvalidWidth: 12000,
			ooeInvalidHeight: 12001,
			ooeNavigationError: 12002,
			ooeInvalidScheme: 12003,
			ooeAppDomains: 12004,
			ooeRequireHTTPS: 12005,
			ooeWebDialogClosed: 12006,
			ooeDialogAlreadyOpened: 12007,
			ooeEndUserAllow: 12008,
			ooeEndUserIgnore: 12009,
			ooeNotUILessDialog: 12010,
			ooeCrossZone: 12011,
			ooeNotSSOAgave: 13000,
			ooeSSOUserNotSignedIn: 13001,
			ooeSSOUserAborted: 13002,
			ooeSSOUnsupportedUserIdentity: 13003,
			ooeSSOInvalidResourceUrl: 13004,
			ooeSSOInvalidGrant: 13005,
			ooeSSOClientError: 13006,
			ooeSSOServerError: 13007,
			ooeAddinIsAlreadyRequestingToken: 13008,
			ooeSSOUserConsentNotSupportedByCurrentAddinCategory: 13009,
			ooeSSOConnectionLost: 13010,
			ooeResourceNotAllowed: 13011,
			ooeSSOUnsupportedPlatform: 13012,
			ooeAccessDenied: 13990,
			ooeGeneralException: 13991
		},
		initializeErrorMessages: function OSF_DDA_ErrorCodeManager$initializeErrorMessages(stringNS) {
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeCoercionTypeNotSupported]={ name: stringNS.L_InvalidCoercion, message: stringNS.L_CoercionTypeNotSupported };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeGetSelectionNotMatchDataType]={ name: stringNS.L_DataReadError, message: stringNS.L_GetSelectionNotSupported };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeCoercionTypeNotMatchBinding]={ name: stringNS.L_InvalidCoercion, message: stringNS.L_CoercionTypeNotMatchBinding };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidGetRowColumnCounts]={ name: stringNS.L_DataReadError, message: stringNS.L_InvalidGetRowColumnCounts };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSelectionNotSupportCoercionType]={ name: stringNS.L_DataReadError, message: stringNS.L_SelectionNotSupportCoercionType };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidGetStartRowColumn]={ name: stringNS.L_DataReadError, message: stringNS.L_InvalidGetStartRowColumn };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeNonUniformPartialGetNotSupported]={ name: stringNS.L_DataReadError, message: stringNS.L_NonUniformPartialGetNotSupported };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeGetDataIsTooLarge]={ name: stringNS.L_DataReadError, message: stringNS.L_GetDataIsTooLarge };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeFileTypeNotSupported]={ name: stringNS.L_DataReadError, message: stringNS.L_FileTypeNotSupported };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeGetDataParametersConflict]={ name: stringNS.L_DataReadError, message: stringNS.L_GetDataParametersConflict };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidGetColumns]={ name: stringNS.L_DataReadError, message: stringNS.L_InvalidGetColumns };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidGetRows]={ name: stringNS.L_DataReadError, message: stringNS.L_InvalidGetRows };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidReadForBlankRow]={ name: stringNS.L_DataReadError, message: stringNS.L_InvalidReadForBlankRow };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeUnsupportedDataObject]={ name: stringNS.L_DataWriteError, message: stringNS.L_UnsupportedDataObject };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeCannotWriteToSelection]={ name: stringNS.L_DataWriteError, message: stringNS.L_CannotWriteToSelection };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeDataNotMatchSelection]={ name: stringNS.L_DataWriteError, message: stringNS.L_DataNotMatchSelection };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeOverwriteWorksheetData]={ name: stringNS.L_DataWriteError, message: stringNS.L_OverwriteWorksheetData };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeDataNotMatchBindingSize]={ name: stringNS.L_DataWriteError, message: stringNS.L_DataNotMatchBindingSize };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidSetStartRowColumn]={ name: stringNS.L_DataWriteError, message: stringNS.L_InvalidSetStartRowColumn };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidDataFormat]={ name: stringNS.L_InvalidFormat, message: stringNS.L_InvalidDataFormat };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeDataNotMatchCoercionType]={ name: stringNS.L_InvalidDataObject, message: stringNS.L_DataNotMatchCoercionType };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeDataNotMatchBindingType]={ name: stringNS.L_InvalidDataObject, message: stringNS.L_DataNotMatchBindingType };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSetDataIsTooLarge]={ name: stringNS.L_DataWriteError, message: stringNS.L_SetDataIsTooLarge };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeNonUniformPartialSetNotSupported]={ name: stringNS.L_DataWriteError, message: stringNS.L_NonUniformPartialSetNotSupported };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidSetColumns]={ name: stringNS.L_DataWriteError, message: stringNS.L_InvalidSetColumns };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidSetRows]={ name: stringNS.L_DataWriteError, message: stringNS.L_InvalidSetRows };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSetDataParametersConflict]={ name: stringNS.L_DataWriteError, message: stringNS.L_SetDataParametersConflict };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSelectionCannotBound]={ name: stringNS.L_BindingCreationError, message: stringNS.L_SelectionCannotBound };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeBindingNotExist]={ name: stringNS.L_InvalidBindingError, message: stringNS.L_BindingNotExist };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeBindingToMultipleSelection]={ name: stringNS.L_BindingCreationError, message: stringNS.L_BindingToMultipleSelection };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidSelectionForBindingType]={ name: stringNS.L_BindingCreationError, message: stringNS.L_InvalidSelectionForBindingType };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeOperationNotSupportedOnThisBindingType]={ name: stringNS.L_InvalidBindingOperation, message: stringNS.L_OperationNotSupportedOnThisBindingType };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeNamedItemNotFound]={ name: stringNS.L_BindingCreationError, message: stringNS.L_NamedItemNotFound };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeMultipleNamedItemFound]={ name: stringNS.L_BindingCreationError, message: stringNS.L_MultipleNamedItemFound };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidNamedItemForBindingType]={ name: stringNS.L_BindingCreationError, message: stringNS.L_InvalidNamedItemForBindingType };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeUnknownBindingType]={ name: stringNS.L_InvalidBinding, message: stringNS.L_UnknownBindingType };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeOperationNotSupportedOnMatrixData]={ name: stringNS.L_InvalidBindingOperation, message: stringNS.L_OperationNotSupportedOnMatrixData };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidColumnsForBinding]={ name: stringNS.L_InvalidBinding, message: stringNS.L_InvalidColumnsForBinding };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSettingNameNotExist]={ name: stringNS.L_ReadSettingsError, message: stringNS.L_SettingNameNotExist };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSettingsCannotSave]={ name: stringNS.L_SaveSettingsError, message: stringNS.L_SettingsCannotSave };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSettingsAreStale]={ name: stringNS.L_SettingsStaleError, message: stringNS.L_SettingsAreStale };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeOperationNotSupported]={ name: stringNS.L_HostError, message: stringNS.L_OperationNotSupported };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError]={ name: stringNS.L_InternalError, message: stringNS.L_InternalErrorDescription };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeDocumentReadOnly]={ name: stringNS.L_PermissionDenied, message: stringNS.L_DocumentReadOnly };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeEventHandlerNotExist]={ name: stringNS.L_EventRegistrationError, message: stringNS.L_EventHandlerNotExist };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidApiCallInContext]={ name: stringNS.L_InvalidAPICall, message: stringNS.L_InvalidApiCallInContext };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeShuttingDown]={ name: stringNS.L_ShuttingDown, message: stringNS.L_ShuttingDown };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeUnsupportedEnumeration]={ name: stringNS.L_UnsupportedEnumeration, message: stringNS.L_UnsupportedEnumerationMessage };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeIndexOutOfRange]={ name: stringNS.L_IndexOutOfRange, message: stringNS.L_IndexOutOfRange };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeBrowserAPINotSupported]={ name: stringNS.L_APINotSupported, message: stringNS.L_BrowserAPINotSupported };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeRequestTimeout]={ name: stringNS.L_APICallFailed, message: stringNS.L_RequestTimeout };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidOrTimedOutSession]={ name: stringNS.L_InvalidOrTimedOutSession, message: stringNS.L_InvalidOrTimedOutSessionMessage };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeTooManyIncompleteRequests]={ name: stringNS.L_APICallFailed, message: stringNS.L_TooManyIncompleteRequests };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeRequestTokenUnavailable]={ name: stringNS.L_APICallFailed, message: stringNS.L_RequestTokenUnavailable };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeActivityLimitReached]={ name: stringNS.L_APICallFailed, message: stringNS.L_ActivityLimitReached };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidApiArguments]={ name: stringNS.L_APICallFailed, message: stringNS.L_InvalidApiArgumentsMessage };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeWorkbookHidden]={ name: stringNS.L_APICallFailed, message: stringNS.L_WorkbookHiddenMessage };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeCustomXmlNodeNotFound]={ name: stringNS.L_InvalidNode, message: stringNS.L_CustomXmlNodeNotFound };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeCustomXmlError]={ name: stringNS.L_CustomXmlError, message: stringNS.L_CustomXmlError };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeCustomXmlExceedQuota]={ name: stringNS.L_CustomXmlExceedQuotaName, message: stringNS.L_CustomXmlExceedQuotaMessage };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeCustomXmlOutOfDate]={ name: stringNS.L_CustomXmlOutOfDateName, message: stringNS.L_CustomXmlOutOfDateMessage };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeNoCapability]={ name: stringNS.L_PermissionDenied, message: stringNS.L_NoCapability };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeCannotNavTo]={ name: stringNS.L_CannotNavigateTo, message: stringNS.L_CannotNavigateTo };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSpecifiedIdNotExist]={ name: stringNS.L_SpecifiedIdNotExist, message: stringNS.L_SpecifiedIdNotExist };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeNavOutOfBound]={ name: stringNS.L_NavOutOfBound, message: stringNS.L_NavOutOfBound };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeCellDataAmountBeyondLimits]={ name: stringNS.L_DataWriteReminder, message: stringNS.L_CellDataAmountBeyondLimits };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeElementMissing]={ name: stringNS.L_MissingParameter, message: stringNS.L_ElementMissing };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeProtectedError]={ name: stringNS.L_PermissionDenied, message: stringNS.L_NoCapability };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidCellsValue]={ name: stringNS.L_InvalidValue, message: stringNS.L_InvalidCellsValue };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidTableOptionValue]={ name: stringNS.L_InvalidValue, message: stringNS.L_InvalidTableOptionValue };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidFormatValue]={ name: stringNS.L_InvalidValue, message: stringNS.L_InvalidFormatValue };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeRowIndexOutOfRange]={ name: stringNS.L_OutOfRange, message: stringNS.L_RowIndexOutOfRange };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeColIndexOutOfRange]={ name: stringNS.L_OutOfRange, message: stringNS.L_ColIndexOutOfRange };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeFormatValueOutOfRange]={ name: stringNS.L_OutOfRange, message: stringNS.L_FormatValueOutOfRange };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeCellFormatAmountBeyondLimits]={ name: stringNS.L_FormattingReminder, message: stringNS.L_CellFormatAmountBeyondLimits };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeMemoryFileLimit]={ name: stringNS.L_MemoryLimit, message: stringNS.L_CloseFileBeforeRetrieve };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeNetworkProblemRetrieveFile]={ name: stringNS.L_NetworkProblem, message: stringNS.L_NetworkProblemRetrieveFile };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidSliceSize]={ name: stringNS.L_InvalidValue, message: stringNS.L_SliceSizeNotSupported };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeDialogAlreadyOpened]={ name: stringNS.L_DisplayDialogError, message: stringNS.L_DialogAlreadyOpened };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidWidth]={ name: stringNS.L_IndexOutOfRange, message: stringNS.L_IndexOutOfRange };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidHeight]={ name: stringNS.L_IndexOutOfRange, message: stringNS.L_IndexOutOfRange };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeNavigationError]={ name: stringNS.L_DisplayDialogError, message: stringNS.L_NetworkProblem };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidScheme]={ name: stringNS.L_DialogNavigateError, message: stringNS.L_DialogInvalidScheme };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeAppDomains]={ name: stringNS.L_DisplayDialogError, message: stringNS.L_DialogAddressNotTrusted };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeRequireHTTPS]={ name: stringNS.L_DisplayDialogError, message: stringNS.L_DialogRequireHTTPS };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeEndUserIgnore]={ name: stringNS.L_DisplayDialogError, message: stringNS.L_UserClickIgnore };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeCrossZone]={ name: stringNS.L_DisplayDialogError, message: stringNS.L_NewWindowCrossZoneErrorString };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeNotSSOAgave]={ name: stringNS.L_APINotSupported, message: stringNS.L_InvalidSSOAddinMessage };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSSOUserNotSignedIn]={ name: stringNS.L_UserNotSignedIn, message: stringNS.L_UserNotSignedIn };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSSOUserAborted]={ name: stringNS.L_UserAborted, message: stringNS.L_UserAbortedMessage };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSSOUnsupportedUserIdentity]={ name: stringNS.L_UnsupportedUserIdentity, message: stringNS.L_UnsupportedUserIdentityMessage };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSSOInvalidResourceUrl]={ name: stringNS.L_InvalidResourceUrl, message: stringNS.L_InvalidResourceUrlMessage };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSSOInvalidGrant]={ name: stringNS.L_InvalidGrant, message: stringNS.L_InvalidGrantMessage };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSSOClientError]={ name: stringNS.L_SSOClientError, message: stringNS.L_SSOClientErrorMessage };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSSOServerError]={ name: stringNS.L_SSOServerError, message: stringNS.L_SSOServerErrorMessage };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeAddinIsAlreadyRequestingToken]={ name: stringNS.L_AddinIsAlreadyRequestingToken, message: stringNS.L_AddinIsAlreadyRequestingTokenMessage };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSSOUserConsentNotSupportedByCurrentAddinCategory]={ name: stringNS.L_SSOUserConsentNotSupportedByCurrentAddinCategory, message: stringNS.L_SSOUserConsentNotSupportedByCurrentAddinCategoryMessage };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSSOConnectionLost]={ name: stringNS.L_SSOConnectionLostError, message: stringNS.L_SSOConnectionLostErrorMessage };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSSOUnsupportedPlatform]={ name: stringNS.L_SSOConnectionLostError, message: stringNS.L_SSOUnsupportedPlatform };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeOperationCancelled]={ name: stringNS.L_OperationCancelledError, message: stringNS.L_OperationCancelledErrorMessage };
		}
	};
})();
var OfficeExt;
(function (OfficeExt) {
	var Requirement;
	(function (Requirement) {
		var RequirementVersion=(function () {
			function RequirementVersion() {
			}
			return RequirementVersion;
		})();
		Requirement.RequirementVersion=RequirementVersion;
		var RequirementMatrix=(function () {
			function RequirementMatrix(_setMap) {
				this.isSetSupported=function _isSetSupported(name, minVersion) {
					if (name==undefined) {
						return false;
					}
					if (minVersion==undefined) {
						minVersion=0;
					}
					var setSupportArray=this._setMap;
					var sets=setSupportArray._sets;
					if (sets.hasOwnProperty(name.toLowerCase())) {
						var setMaxVersion=sets[name.toLowerCase()];
						try {
							var setMaxVersionNum=this._getVersion(setMaxVersion);
							minVersion=minVersion+"";
							var minVersionNum=this._getVersion(minVersion);
							if (setMaxVersionNum.major > 0 && setMaxVersionNum.major > minVersionNum.major) {
								return true;
							}
							if (setMaxVersionNum.minor > 0 &&
								setMaxVersionNum.minor > 0 &&
								setMaxVersionNum.major==minVersionNum.major &&
								setMaxVersionNum.minor >=minVersionNum.minor) {
								return true;
							}
						}
						catch (e) {
							return false;
						}
					}
					return false;
				};
				this._getVersion=function (version) {
					version=version+"";
					var temp=version.split(".");
					var major=0;
					var minor=0;
					if (temp.length < 2 && isNaN(Number(version))) {
						throw "version format incorrect";
					}
					else {
						major=Number(temp[0]);
						if (temp.length >=2) {
							minor=Number(temp[1]);
						}
						if (isNaN(major) || isNaN(minor)) {
							throw "version format incorrect";
						}
					}
					var result={ "minor": minor, "major": major };
					return result;
				};
				this._setMap=_setMap;
				this.isSetSupported=this.isSetSupported.bind(this);
			}
			return RequirementMatrix;
		})();
		Requirement.RequirementMatrix=RequirementMatrix;
		var DefaultSetRequirement=(function () {
			function DefaultSetRequirement(setMap) {
				this._addSetMap=function DefaultSetRequirement_addSetMap(addedSet) {
					for (var name in addedSet) {
						this._sets[name]=addedSet[name];
					}
				};
				this._sets=setMap;
			}
			return DefaultSetRequirement;
		})();
		Requirement.DefaultSetRequirement=DefaultSetRequirement;
		var DefaultDialogSetRequirement=(function (_super) {
			__extends(DefaultDialogSetRequirement, _super);
			function DefaultDialogSetRequirement() {
				_super.call(this, {
					"dialogapi": 1.1
				});
			}
			return DefaultDialogSetRequirement;
		})(DefaultSetRequirement);
		Requirement.DefaultDialogSetRequirement=DefaultDialogSetRequirement;
		var ExcelClientDefaultSetRequirement=(function (_super) {
			__extends(ExcelClientDefaultSetRequirement, _super);
			function ExcelClientDefaultSetRequirement() {
				_super.call(this, {
					"bindingevents": 1.1,
					"documentevents": 1.1,
					"excelapi": 1.1,
					"matrixbindings": 1.1,
					"matrixcoercion": 1.1,
					"selection": 1.1,
					"settings": 1.1,
					"tablebindings": 1.1,
					"tablecoercion": 1.1,
					"textbindings": 1.1,
					"textcoercion": 1.1
				});
			}
			return ExcelClientDefaultSetRequirement;
		})(DefaultSetRequirement);
		Requirement.ExcelClientDefaultSetRequirement=ExcelClientDefaultSetRequirement;
		var ExcelClientV1DefaultSetRequirement=(function (_super) {
			__extends(ExcelClientV1DefaultSetRequirement, _super);
			function ExcelClientV1DefaultSetRequirement() {
				_super.call(this);
				this._addSetMap({
					"imagecoercion": 1.1
				});
			}
			return ExcelClientV1DefaultSetRequirement;
		})(ExcelClientDefaultSetRequirement);
		Requirement.ExcelClientV1DefaultSetRequirement=ExcelClientV1DefaultSetRequirement;
		var OutlookClientDefaultSetRequirement=(function (_super) {
			__extends(OutlookClientDefaultSetRequirement, _super);
			function OutlookClientDefaultSetRequirement() {
				_super.call(this, {
					"mailbox": 1.3
				});
			}
			return OutlookClientDefaultSetRequirement;
		})(DefaultSetRequirement);
		Requirement.OutlookClientDefaultSetRequirement=OutlookClientDefaultSetRequirement;
		var WordClientDefaultSetRequirement=(function (_super) {
			__extends(WordClientDefaultSetRequirement, _super);
			function WordClientDefaultSetRequirement() {
				_super.call(this, {
					"bindingevents": 1.1,
					"compressedfile": 1.1,
					"customxmlparts": 1.1,
					"documentevents": 1.1,
					"file": 1.1,
					"htmlcoercion": 1.1,
					"matrixbindings": 1.1,
					"matrixcoercion": 1.1,
					"ooxmlcoercion": 1.1,
					"pdffile": 1.1,
					"selection": 1.1,
					"settings": 1.1,
					"tablebindings": 1.1,
					"tablecoercion": 1.1,
					"textbindings": 1.1,
					"textcoercion": 1.1,
					"textfile": 1.1,
					"wordapi": 1.1
				});
			}
			return WordClientDefaultSetRequirement;
		})(DefaultSetRequirement);
		Requirement.WordClientDefaultSetRequirement=WordClientDefaultSetRequirement;
		var WordClientV1DefaultSetRequirement=(function (_super) {
			__extends(WordClientV1DefaultSetRequirement, _super);
			function WordClientV1DefaultSetRequirement() {
				_super.call(this);
				this._addSetMap({
					"customxmlparts": 1.2,
					"wordapi": 1.2,
					"imagecoercion": 1.1
				});
			}
			return WordClientV1DefaultSetRequirement;
		})(WordClientDefaultSetRequirement);
		Requirement.WordClientV1DefaultSetRequirement=WordClientV1DefaultSetRequirement;
		var PowerpointClientDefaultSetRequirement=(function (_super) {
			__extends(PowerpointClientDefaultSetRequirement, _super);
			function PowerpointClientDefaultSetRequirement() {
				_super.call(this, {
					"activeview": 1.1,
					"compressedfile": 1.1,
					"documentevents": 1.1,
					"file": 1.1,
					"pdffile": 1.1,
					"selection": 1.1,
					"settings": 1.1,
					"textcoercion": 1.1
				});
			}
			return PowerpointClientDefaultSetRequirement;
		})(DefaultSetRequirement);
		Requirement.PowerpointClientDefaultSetRequirement=PowerpointClientDefaultSetRequirement;
		var PowerpointClientV1DefaultSetRequirement=(function (_super) {
			__extends(PowerpointClientV1DefaultSetRequirement, _super);
			function PowerpointClientV1DefaultSetRequirement() {
				_super.call(this);
				this._addSetMap({
					"imagecoercion": 1.1
				});
			}
			return PowerpointClientV1DefaultSetRequirement;
		})(PowerpointClientDefaultSetRequirement);
		Requirement.PowerpointClientV1DefaultSetRequirement=PowerpointClientV1DefaultSetRequirement;
		var ProjectClientDefaultSetRequirement=(function (_super) {
			__extends(ProjectClientDefaultSetRequirement, _super);
			function ProjectClientDefaultSetRequirement() {
				_super.call(this, {
					"selection": 1.1,
					"textcoercion": 1.1
				});
			}
			return ProjectClientDefaultSetRequirement;
		})(DefaultSetRequirement);
		Requirement.ProjectClientDefaultSetRequirement=ProjectClientDefaultSetRequirement;
		var ExcelWebDefaultSetRequirement=(function (_super) {
			__extends(ExcelWebDefaultSetRequirement, _super);
			function ExcelWebDefaultSetRequirement() {
				_super.call(this, {
					"bindingevents": 1.1,
					"documentevents": 1.1,
					"matrixbindings": 1.1,
					"matrixcoercion": 1.1,
					"selection": 1.1,
					"settings": 1.1,
					"tablebindings": 1.1,
					"tablecoercion": 1.1,
					"textbindings": 1.1,
					"textcoercion": 1.1,
					"file": 1.1
				});
			}
			return ExcelWebDefaultSetRequirement;
		})(DefaultSetRequirement);
		Requirement.ExcelWebDefaultSetRequirement=ExcelWebDefaultSetRequirement;
		var WordWebDefaultSetRequirement=(function (_super) {
			__extends(WordWebDefaultSetRequirement, _super);
			function WordWebDefaultSetRequirement() {
				_super.call(this, {
					"compressedfile": 1.1,
					"documentevents": 1.1,
					"file": 1.1,
					"imagecoercion": 1.1,
					"matrixcoercion": 1.1,
					"ooxmlcoercion": 1.1,
					"pdffile": 1.1,
					"selection": 1.1,
					"settings": 1.1,
					"tablecoercion": 1.1,
					"textcoercion": 1.1,
					"textfile": 1.1
				});
			}
			return WordWebDefaultSetRequirement;
		})(DefaultSetRequirement);
		Requirement.WordWebDefaultSetRequirement=WordWebDefaultSetRequirement;
		var PowerpointWebDefaultSetRequirement=(function (_super) {
			__extends(PowerpointWebDefaultSetRequirement, _super);
			function PowerpointWebDefaultSetRequirement() {
				_super.call(this, {
					"activeview": 1.1,
					"settings": 1.1
				});
			}
			return PowerpointWebDefaultSetRequirement;
		})(DefaultSetRequirement);
		Requirement.PowerpointWebDefaultSetRequirement=PowerpointWebDefaultSetRequirement;
		var OutlookWebDefaultSetRequirement=(function (_super) {
			__extends(OutlookWebDefaultSetRequirement, _super);
			function OutlookWebDefaultSetRequirement() {
				_super.call(this, {
					"mailbox": 1.3
				});
			}
			return OutlookWebDefaultSetRequirement;
		})(DefaultSetRequirement);
		Requirement.OutlookWebDefaultSetRequirement=OutlookWebDefaultSetRequirement;
		var SwayWebDefaultSetRequirement=(function (_super) {
			__extends(SwayWebDefaultSetRequirement, _super);
			function SwayWebDefaultSetRequirement() {
				_super.call(this, {
					"activeview": 1.1,
					"documentevents": 1.1,
					"selection": 1.1,
					"settings": 1.1,
					"textcoercion": 1.1
				});
			}
			return SwayWebDefaultSetRequirement;
		})(DefaultSetRequirement);
		Requirement.SwayWebDefaultSetRequirement=SwayWebDefaultSetRequirement;
		var AccessWebDefaultSetRequirement=(function (_super) {
			__extends(AccessWebDefaultSetRequirement, _super);
			function AccessWebDefaultSetRequirement() {
				_super.call(this, {
					"bindingevents": 1.1,
					"partialtablebindings": 1.1,
					"settings": 1.1,
					"tablebindings": 1.1,
					"tablecoercion": 1.1
				});
			}
			return AccessWebDefaultSetRequirement;
		})(DefaultSetRequirement);
		Requirement.AccessWebDefaultSetRequirement=AccessWebDefaultSetRequirement;
		var ExcelIOSDefaultSetRequirement=(function (_super) {
			__extends(ExcelIOSDefaultSetRequirement, _super);
			function ExcelIOSDefaultSetRequirement() {
				_super.call(this, {
					"bindingevents": 1.1,
					"documentevents": 1.1,
					"matrixbindings": 1.1,
					"matrixcoercion": 1.1,
					"selection": 1.1,
					"settings": 1.1,
					"tablebindings": 1.1,
					"tablecoercion": 1.1,
					"textbindings": 1.1,
					"textcoercion": 1.1
				});
			}
			return ExcelIOSDefaultSetRequirement;
		})(DefaultSetRequirement);
		Requirement.ExcelIOSDefaultSetRequirement=ExcelIOSDefaultSetRequirement;
		var WordIOSDefaultSetRequirement=(function (_super) {
			__extends(WordIOSDefaultSetRequirement, _super);
			function WordIOSDefaultSetRequirement() {
				_super.call(this, {
					"bindingevents": 1.1,
					"compressedfile": 1.1,
					"customxmlparts": 1.1,
					"documentevents": 1.1,
					"file": 1.1,
					"htmlcoercion": 1.1,
					"matrixbindings": 1.1,
					"matrixcoercion": 1.1,
					"ooxmlcoercion": 1.1,
					"pdffile": 1.1,
					"selection": 1.1,
					"settings": 1.1,
					"tablebindings": 1.1,
					"tablecoercion": 1.1,
					"textbindings": 1.1,
					"textcoercion": 1.1,
					"textfile": 1.1
				});
			}
			return WordIOSDefaultSetRequirement;
		})(DefaultSetRequirement);
		Requirement.WordIOSDefaultSetRequirement=WordIOSDefaultSetRequirement;
		var WordIOSV1DefaultSetRequirement=(function (_super) {
			__extends(WordIOSV1DefaultSetRequirement, _super);
			function WordIOSV1DefaultSetRequirement() {
				_super.call(this);
				this._addSetMap({
					"customxmlparts": 1.2,
					"wordapi": 1.2
				});
			}
			return WordIOSV1DefaultSetRequirement;
		})(WordIOSDefaultSetRequirement);
		Requirement.WordIOSV1DefaultSetRequirement=WordIOSV1DefaultSetRequirement;
		var PowerpointIOSDefaultSetRequirement=(function (_super) {
			__extends(PowerpointIOSDefaultSetRequirement, _super);
			function PowerpointIOSDefaultSetRequirement() {
				_super.call(this, {
					"activeview": 1.1,
					"compressedfile": 1.1,
					"documentevents": 1.1,
					"file": 1.1,
					"pdffile": 1.1,
					"selection": 1.1,
					"settings": 1.1,
					"textcoercion": 1.1
				});
			}
			return PowerpointIOSDefaultSetRequirement;
		})(DefaultSetRequirement);
		Requirement.PowerpointIOSDefaultSetRequirement=PowerpointIOSDefaultSetRequirement;
		var OutlookIOSDefaultSetRequirement=(function (_super) {
			__extends(OutlookIOSDefaultSetRequirement, _super);
			function OutlookIOSDefaultSetRequirement() {
				_super.call(this, {
					"mailbox": 1.1
				});
			}
			return OutlookIOSDefaultSetRequirement;
		})(DefaultSetRequirement);
		Requirement.OutlookIOSDefaultSetRequirement=OutlookIOSDefaultSetRequirement;
		var RequirementsMatrixFactory=(function () {
			function RequirementsMatrixFactory() {
			}
			RequirementsMatrixFactory.initializeOsfDda=function () {
				OSF.OUtil.setNamespace("Requirement", OSF.DDA);
			};
			RequirementsMatrixFactory.getDefaultRequirementMatrix=function (appContext) {
				this.initializeDefaultSetMatrix();
				var defaultRequirementMatrix=undefined;
				var clientRequirement=appContext.get_requirementMatrix();
				if (clientRequirement !=undefined && clientRequirement.length > 0 && typeof (JSON) !=="undefined") {
					var matrixItem=JSON.parse(appContext.get_requirementMatrix().toLowerCase());
					defaultRequirementMatrix=new RequirementMatrix(new DefaultSetRequirement(matrixItem));
				}
				else {
					var appLocator=RequirementsMatrixFactory.getClientFullVersionString(appContext);
					if (RequirementsMatrixFactory.DefaultSetArrayMatrix !=undefined && RequirementsMatrixFactory.DefaultSetArrayMatrix[appLocator] !=undefined) {
						defaultRequirementMatrix=new RequirementMatrix(RequirementsMatrixFactory.DefaultSetArrayMatrix[appLocator]);
					}
					else {
						defaultRequirementMatrix=new RequirementMatrix(new DefaultSetRequirement({}));
					}
				}
				return defaultRequirementMatrix;
			};
			RequirementsMatrixFactory.getDefaultDialogRequirementMatrix=function (appContext) {
				var defaultRequirementMatrix=undefined;
				var clientRequirement=appContext.get_dialogRequirementMatrix();
				if (clientRequirement !=undefined && clientRequirement.length > 0 && typeof (JSON) !=="undefined") {
					var matrixItem=JSON.parse(appContext.get_requirementMatrix().toLowerCase());
					defaultRequirementMatrix=new RequirementMatrix(new DefaultSetRequirement(matrixItem));
				}
				else {
					defaultRequirementMatrix=new RequirementMatrix(new DefaultDialogSetRequirement());
				}
				return defaultRequirementMatrix;
			};
			RequirementsMatrixFactory.getClientFullVersionString=function (appContext) {
				var appMinorVersion=appContext.get_appMinorVersion();
				var appMinorVersionString="";
				var appFullVersion="";
				var appName=appContext.get_appName();
				var isIOSClient=appName==1024 ||
					appName==4096 ||
					appName==8192 ||
					appName==65536;
				if (isIOSClient && appContext.get_appVersion()==1) {
					if (appName==4096 && appMinorVersion >=15) {
						appFullVersion="16.00.01";
					}
					else {
						appFullVersion="16.00";
					}
				}
				else if (appContext.get_appName()==64) {
					appFullVersion=appContext.get_appVersion();
				}
				else {
					if (appMinorVersion < 10) {
						appMinorVersionString="0"+appMinorVersion;
					}
					else {
						appMinorVersionString=""+appMinorVersion;
					}
					appFullVersion=appContext.get_appVersion()+"."+appMinorVersionString;
				}
				return appContext.get_appName()+"-"+appFullVersion;
			};
			RequirementsMatrixFactory.initializeDefaultSetMatrix=function () {
				RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.Excel_RCLIENT_1600]=new ExcelClientDefaultSetRequirement();
				RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.Word_RCLIENT_1600]=new WordClientDefaultSetRequirement();
				RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.PowerPoint_RCLIENT_1600]=new PowerpointClientDefaultSetRequirement();
				RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.Excel_RCLIENT_1601]=new ExcelClientV1DefaultSetRequirement();
				RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.Word_RCLIENT_1601]=new WordClientV1DefaultSetRequirement();
				RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.PowerPoint_RCLIENT_1601]=new PowerpointClientV1DefaultSetRequirement();
				RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.Outlook_RCLIENT_1600]=new OutlookClientDefaultSetRequirement();
				RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.Excel_WAC_1600]=new ExcelWebDefaultSetRequirement();
				RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.Word_WAC_1600]=new WordWebDefaultSetRequirement();
				RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.Outlook_WAC_1600]=new OutlookWebDefaultSetRequirement();
				RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.Outlook_WAC_1601]=new OutlookWebDefaultSetRequirement();
				RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.Project_RCLIENT_1600]=new ProjectClientDefaultSetRequirement();
				RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.Access_WAC_1600]=new AccessWebDefaultSetRequirement();
				RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.PowerPoint_WAC_1600]=new PowerpointWebDefaultSetRequirement();
				RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.Excel_IOS_1600]=new ExcelIOSDefaultSetRequirement();
				RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.SWAY_WAC_1600]=new SwayWebDefaultSetRequirement();
				RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.Word_IOS_1600]=new WordIOSDefaultSetRequirement();
				RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.Word_IOS_16001]=new WordIOSV1DefaultSetRequirement();
				RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.PowerPoint_IOS_1600]=new PowerpointIOSDefaultSetRequirement();
				RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.Outlook_IOS_1600]=new OutlookIOSDefaultSetRequirement();
			};
			RequirementsMatrixFactory.Excel_RCLIENT_1600="1-16.00";
			RequirementsMatrixFactory.Excel_RCLIENT_1601="1-16.01";
			RequirementsMatrixFactory.Word_RCLIENT_1600="2-16.00";
			RequirementsMatrixFactory.Word_RCLIENT_1601="2-16.01";
			RequirementsMatrixFactory.PowerPoint_RCLIENT_1600="4-16.00";
			RequirementsMatrixFactory.PowerPoint_RCLIENT_1601="4-16.01";
			RequirementsMatrixFactory.Outlook_RCLIENT_1600="8-16.00";
			RequirementsMatrixFactory.Excel_WAC_1600="16-16.00";
			RequirementsMatrixFactory.Word_WAC_1600="32-16.00";
			RequirementsMatrixFactory.Outlook_WAC_1600="64-16.00";
			RequirementsMatrixFactory.Outlook_WAC_1601="64-16.01";
			RequirementsMatrixFactory.Project_RCLIENT_1600="128-16.00";
			RequirementsMatrixFactory.Access_WAC_1600="256-16.00";
			RequirementsMatrixFactory.PowerPoint_WAC_1600="512-16.00";
			RequirementsMatrixFactory.Excel_IOS_1600="1024-16.00";
			RequirementsMatrixFactory.SWAY_WAC_1600="2048-16.00";
			RequirementsMatrixFactory.Word_IOS_1600="4096-16.00";
			RequirementsMatrixFactory.Word_IOS_16001="4096-16.00.01";
			RequirementsMatrixFactory.PowerPoint_IOS_1600="8192-16.00";
			RequirementsMatrixFactory.Outlook_IOS_1600="65536-16.00";
			RequirementsMatrixFactory.DefaultSetArrayMatrix={};
			return RequirementsMatrixFactory;
		})();
		Requirement.RequirementsMatrixFactory=RequirementsMatrixFactory;
	})(Requirement=OfficeExt.Requirement || (OfficeExt.Requirement={}));
})(OfficeExt || (OfficeExt={}));
OfficeExt.Requirement.RequirementsMatrixFactory.initializeOsfDda();
Microsoft.Office.WebExtension.ApplicationMode={
	WebEditor: "webEditor",
	WebViewer: "webViewer",
	Client: "client"
};
Microsoft.Office.WebExtension.DocumentMode={
	ReadOnly: "readOnly",
	ReadWrite: "readWrite"
};
OSF.NamespaceManager=(function OSF_NamespaceManager() {
	var _userOffice;
	var _useShortcut=false;
	return {
		enableShortcut: function OSF_NamespaceManager$enableShortcut() {
			if (!_useShortcut) {
				if (window.Office) {
					_userOffice=window.Office;
				}
				else {
					OSF.OUtil.setNamespace("Office", window);
				}
				window.Office=Microsoft.Office.WebExtension;
				_useShortcut=true;
			}
		},
		disableShortcut: function OSF_NamespaceManager$disableShortcut() {
			if (_useShortcut) {
				if (_userOffice) {
					window.Office=_userOffice;
				}
				else {
					OSF.OUtil.unsetNamespace("Office", window);
				}
				_useShortcut=false;
			}
		}
	};
})();
OSF.NamespaceManager.enableShortcut();
Microsoft.Office.WebExtension.useShortNamespace=function Microsoft_Office_WebExtension_useShortNamespace(useShortcut) {
	if (useShortcut) {
		OSF.NamespaceManager.enableShortcut();
	}
	else {
		OSF.NamespaceManager.disableShortcut();
	}
};
Microsoft.Office.WebExtension.select=function Microsoft_Office_WebExtension_select(str, errorCallback) {
	var promise;
	if (str && typeof str=="string") {
		var index=str.indexOf("#");
		if (index !=-1) {
			var op=str.substring(0, index);
			var target=str.substring(index+1);
			switch (op) {
				case "binding":
				case "bindings":
					if (target) {
						promise=new OSF.DDA.BindingPromise(target);
					}
					break;
			}
		}
	}
	if (!promise) {
		if (errorCallback) {
			var callbackType=typeof errorCallback;
			if (callbackType=="function") {
				var callArgs={};
				callArgs[Microsoft.Office.WebExtension.Parameters.Callback]=errorCallback;
				OSF.DDA.issueAsyncResult(callArgs, OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidApiCallInContext, OSF.DDA.ErrorCodeManager.getErrorArgs(OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidApiCallInContext));
			}
			else {
				throw OSF.OUtil.formatString(Strings.OfficeOM.L_CallbackNotAFunction, callbackType);
			}
		}
	}
	else {
		promise.onFail=errorCallback;
		return promise;
	}
};
OSF.DDA.Context=function OSF_DDA_Context(officeAppContext, document, license, appOM, getOfficeTheme) {
	OSF.OUtil.defineEnumerableProperties(this, {
		"contentLanguage": {
			value: officeAppContext.get_dataLocale()
		},
		"displayLanguage": {
			value: officeAppContext.get_appUILocale()
		},
		"touchEnabled": {
			value: officeAppContext.get_touchEnabled()
		},
		"commerceAllowed": {
			value: officeAppContext.get_commerceAllowed()
		},
		"host": {
			value: OfficeExt.HostName.Host.getInstance().getHost()
		},
		"platform": {
			value: OfficeExt.HostName.Host.getInstance().getPlatform()
		},
		"isDialog": {
			value: OSF._OfficeAppFactory.getHostInfo().isDialog
		},
		"diagnostics": {
			value: OfficeExt.HostName.Host.getInstance().getDiagnostics(officeAppContext.get_hostFullVersion())
		}
	});
	if (license) {
		OSF.OUtil.defineEnumerableProperty(this, "license", {
			value: license
		});
	}
	if (officeAppContext.ui) {
		OSF.OUtil.defineEnumerableProperty(this, "ui", {
			value: officeAppContext.ui
		});
	}
	if (officeAppContext.auth) {
		OSF.OUtil.defineEnumerableProperty(this, "auth", {
			value: officeAppContext.auth
		});
	}
	if (officeAppContext.application) {
		OSF.OUtil.defineEnumerableProperty(this, "application", {
			value: officeAppContext.application
		});
	}
	if (officeAppContext.get_isDialog()) {
		var requirements=OfficeExt.Requirement.RequirementsMatrixFactory.getDefaultDialogRequirementMatrix(officeAppContext);
		OSF.OUtil.defineEnumerableProperty(this, "requirements", {
			value: requirements
		});
	}
	else {
		if (document) {
			OSF.OUtil.defineEnumerableProperty(this, "document", {
				value: document
			});
		}
		if (appOM) {
			var displayName=appOM.displayName || "appOM";
			delete appOM.displayName;
			OSF.OUtil.defineEnumerableProperty(this, displayName, {
				value: appOM
			});
		}
		if (getOfficeTheme) {
			OSF.OUtil.defineEnumerableProperty(this, "officeTheme", {
				get: function () {
					return getOfficeTheme();
				}
			});
		}
		var requirements=OfficeExt.Requirement.RequirementsMatrixFactory.getDefaultRequirementMatrix(officeAppContext);
		OSF.OUtil.defineEnumerableProperty(this, "requirements", {
			value: requirements
		});
	}
};
OSF.DDA.OutlookContext=function OSF_DDA_OutlookContext(appContext, settings, license, appOM, getOfficeTheme) {
	OSF.DDA.OutlookContext.uber.constructor.call(this, appContext, null, license, appOM, getOfficeTheme);
	if (settings) {
		OSF.OUtil.defineEnumerableProperty(this, "roamingSettings", {
			value: settings
		});
	}
};
OSF.OUtil.extend(OSF.DDA.OutlookContext, OSF.DDA.Context);
OSF.DDA.OutlookAppOm=function OSF_DDA_OutlookAppOm(appContext, window, appReady) { };
OSF.DDA.Application=function OSF_DDA_Application(officeAppContext) {
};
OSF.DDA.Document=function OSF_DDA_Document(officeAppContext, settings) {
	var mode;
	switch (officeAppContext.get_clientMode()) {
		case OSF.ClientMode.ReadOnly:
			mode=Microsoft.Office.WebExtension.DocumentMode.ReadOnly;
			break;
		case OSF.ClientMode.ReadWrite:
			mode=Microsoft.Office.WebExtension.DocumentMode.ReadWrite;
			break;
	}
	;
	if (settings) {
		OSF.OUtil.defineEnumerableProperty(this, "settings", {
			value: settings
		});
	}
	;
	OSF.OUtil.defineMutableProperties(this, {
		"mode": {
			value: mode
		},
		"url": {
			value: officeAppContext.get_docUrl()
		}
	});
};
OSF.DDA.JsomDocument=function OSF_DDA_JsomDocument(officeAppContext, bindingFacade, settings) {
	OSF.DDA.JsomDocument.uber.constructor.call(this, officeAppContext, settings);
	if (bindingFacade) {
		OSF.OUtil.defineEnumerableProperty(this, "bindings", {
			get: function OSF_DDA_Document$GetBindings() { return bindingFacade; }
		});
	}
	var am=OSF.DDA.AsyncMethodNames;
	OSF.DDA.DispIdHost.addAsyncMethods(this, [
		am.GetSelectedDataAsync,
		am.SetSelectedDataAsync
	]);
	OSF.DDA.DispIdHost.addEventSupport(this, new OSF.EventDispatch([Microsoft.Office.WebExtension.EventType.DocumentSelectionChanged]));
};
OSF.OUtil.extend(OSF.DDA.JsomDocument, OSF.DDA.Document);
OSF.OUtil.defineEnumerableProperty(Microsoft.Office.WebExtension, "context", {
	get: function Microsoft_Office_WebExtension$GetContext() {
		var context;
		if (OSF && OSF._OfficeAppFactory) {
			context=OSF._OfficeAppFactory.getContext();
		}
		return context;
	}
});
OSF.DDA.License=function OSF_DDA_License(eToken) {
	OSF.OUtil.defineEnumerableProperty(this, "value", {
		value: eToken
	});
};
OSF.DDA.ApiMethodCall=function OSF_DDA_ApiMethodCall(requiredParameters, supportedOptions, privateStateCallbacks, checkCallArgs, displayName) {
	var requiredCount=requiredParameters.length;
	var getInvalidParameterString=OSF.OUtil.delayExecutionAndCache(function () {
		return OSF.OUtil.formatString(Strings.OfficeOM.L_InvalidParameters, displayName);
	});
	this.verifyArguments=function OSF_DDA_ApiMethodCall$VerifyArguments(params, args) {
		for (var name in params) {
			var param=params[name];
			var arg=args[name];
			if (param["enum"]) {
				switch (typeof arg) {
					case "string":
						if (OSF.OUtil.listContainsValue(param["enum"], arg)) {
							break;
						}
					case "undefined":
						throw OSF.DDA.ErrorCodeManager.errorCodes.ooeUnsupportedEnumeration;
					default:
						throw getInvalidParameterString();
				}
			}
			if (param["types"]) {
				if (!OSF.OUtil.listContainsValue(param["types"], typeof arg)) {
					throw getInvalidParameterString();
				}
			}
		}
	};
	this.extractRequiredArguments=function OSF_DDA_ApiMethodCall$ExtractRequiredArguments(userArgs, caller, stateInfo) {
		if (userArgs.length < requiredCount) {
			throw OsfMsAjaxFactory.msAjaxError.parameterCount(Strings.OfficeOM.L_MissingRequiredArguments);
		}
		var requiredArgs=[];
		var index;
		for (index=0; index < requiredCount; index++) {
			requiredArgs.push(userArgs[index]);
		}
		this.verifyArguments(requiredParameters, requiredArgs);
		var ret={};
		for (index=0; index < requiredCount; index++) {
			var param=requiredParameters[index];
			var arg=requiredArgs[index];
			if (param.verify) {
				var isValid=param.verify(arg, caller, stateInfo);
				if (!isValid) {
					throw getInvalidParameterString();
				}
			}
			ret[param.name]=arg;
		}
		return ret;
	},
		this.fillOptions=function OSF_DDA_ApiMethodCall$FillOptions(options, requiredArgs, caller, stateInfo) {
			options=options || {};
			for (var optionName in supportedOptions) {
				if (!OSF.OUtil.listContainsKey(options, optionName)) {
					var value=undefined;
					var option=supportedOptions[optionName];
					if (option.calculate && requiredArgs) {
						value=option.calculate(requiredArgs, caller, stateInfo);
					}
					if (!value && option.defaultValue !==undefined) {
						value=option.defaultValue;
					}
					options[optionName]=value;
				}
			}
			return options;
		};
	this.constructCallArgs=function OSF_DAA_ApiMethodCall$ConstructCallArgs(required, options, caller, stateInfo) {
		var callArgs={};
		for (var r in required) {
			callArgs[r]=required[r];
		}
		for (var o in options) {
			callArgs[o]=options[o];
		}
		for (var s in privateStateCallbacks) {
			callArgs[s]=privateStateCallbacks[s](caller, stateInfo);
		}
		if (checkCallArgs) {
			callArgs=checkCallArgs(callArgs, caller, stateInfo);
		}
		return callArgs;
	};
};
OSF.OUtil.setNamespace("AsyncResultEnum", OSF.DDA);
OSF.DDA.AsyncResultEnum.Properties={
	Context: "Context",
	Value: "Value",
	Status: "Status",
	Error: "Error"
};
Microsoft.Office.WebExtension.AsyncResultStatus={
	Succeeded: "succeeded",
	Failed: "failed"
};
OSF.DDA.AsyncResultEnum.ErrorCode={
	Success: 0,
	Failed: 1
};
OSF.DDA.AsyncResultEnum.ErrorProperties={
	Name: "Name",
	Message: "Message",
	Code: "Code"
};
OSF.DDA.AsyncMethodNames={};
OSF.DDA.AsyncMethodNames.addNames=function (methodNames) {
	for (var entry in methodNames) {
		var am={};
		OSF.OUtil.defineEnumerableProperties(am, {
			"id": {
				value: entry
			},
			"displayName": {
				value: methodNames[entry]
			}
		});
		OSF.DDA.AsyncMethodNames[entry]=am;
	}
};
OSF.DDA.AsyncMethodCall=function OSF_DDA_AsyncMethodCall(requiredParameters, supportedOptions, privateStateCallbacks, onSucceeded, onFailed, checkCallArgs, displayName) {
	var requiredCount=requiredParameters.length;
	var apiMethods=new OSF.DDA.ApiMethodCall(requiredParameters, supportedOptions, privateStateCallbacks, checkCallArgs, displayName);
	function OSF_DAA_AsyncMethodCall$ExtractOptions(userArgs, requiredArgs, caller, stateInfo) {
		if (userArgs.length > requiredCount+2) {
			throw OsfMsAjaxFactory.msAjaxError.parameterCount(Strings.OfficeOM.L_TooManyArguments);
		}
		var options, parameterCallback;
		for (var i=userArgs.length - 1; i >=requiredCount; i--) {
			var argument=userArgs[i];
			switch (typeof argument) {
				case "object":
					if (options) {
						throw OsfMsAjaxFactory.msAjaxError.parameterCount(Strings.OfficeOM.L_TooManyOptionalObjects);
					}
					else {
						options=argument;
					}
					break;
				case "function":
					if (parameterCallback) {
						throw OsfMsAjaxFactory.msAjaxError.parameterCount(Strings.OfficeOM.L_TooManyOptionalFunction);
					}
					else {
						parameterCallback=argument;
					}
					break;
				default:
					throw OsfMsAjaxFactory.msAjaxError.argument(Strings.OfficeOM.L_InValidOptionalArgument);
					break;
			}
		}
		options=apiMethods.fillOptions(options, requiredArgs, caller, stateInfo);
		if (parameterCallback) {
			if (options[Microsoft.Office.WebExtension.Parameters.Callback]) {
				throw Strings.OfficeOM.L_RedundantCallbackSpecification;
			}
			else {
				options[Microsoft.Office.WebExtension.Parameters.Callback]=parameterCallback;
			}
		}
		apiMethods.verifyArguments(supportedOptions, options);
		return options;
	}
	;
	this.verifyAndExtractCall=function OSF_DAA_AsyncMethodCall$VerifyAndExtractCall(userArgs, caller, stateInfo) {
		var required=apiMethods.extractRequiredArguments(userArgs, caller, stateInfo);
		var options=OSF_DAA_AsyncMethodCall$ExtractOptions(userArgs, required, caller, stateInfo);
		var callArgs=apiMethods.constructCallArgs(required, options, caller, stateInfo);
		return callArgs;
	};
	this.processResponse=function OSF_DAA_AsyncMethodCall$ProcessResponse(status, response, caller, callArgs) {
		var payload;
		if (status==OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess) {
			if (onSucceeded) {
				payload=onSucceeded(response, caller, callArgs);
			}
			else {
				payload=response;
			}
		}
		else {
			if (onFailed) {
				payload=onFailed(status, response);
			}
			else {
				payload=OSF.DDA.ErrorCodeManager.getErrorArgs(status);
			}
		}
		return payload;
	};
	this.getCallArgs=function (suppliedArgs) {
		var options, parameterCallback;
		for (var i=suppliedArgs.length - 1; i >=requiredCount; i--) {
			var argument=suppliedArgs[i];
			switch (typeof argument) {
				case "object":
					options=argument;
					break;
				case "function":
					parameterCallback=argument;
					break;
			}
		}
		options=options || {};
		if (parameterCallback) {
			options[Microsoft.Office.WebExtension.Parameters.Callback]=parameterCallback;
		}
		return options;
	};
};
OSF.DDA.AsyncMethodCallFactory=(function () {
	return {
		manufacture: function (params) {
			var supportedOptions=params.supportedOptions ? OSF.OUtil.createObject(params.supportedOptions) : [];
			var privateStateCallbacks=params.privateStateCallbacks ? OSF.OUtil.createObject(params.privateStateCallbacks) : [];
			return new OSF.DDA.AsyncMethodCall(params.requiredArguments || [], supportedOptions, privateStateCallbacks, params.onSucceeded, params.onFailed, params.checkCallArgs, params.method.displayName);
		}
	};
})();
OSF.DDA.AsyncMethodCalls={};
OSF.DDA.AsyncMethodCalls.define=function (callDefinition) {
	OSF.DDA.AsyncMethodCalls[callDefinition.method.id]=OSF.DDA.AsyncMethodCallFactory.manufacture(callDefinition);
};
OSF.DDA.Error=function OSF_DDA_Error(name, message, code) {
	OSF.OUtil.defineEnumerableProperties(this, {
		"name": {
			value: name
		},
		"message": {
			value: message
		},
		"code": {
			value: code
		}
	});
};
OSF.DDA.AsyncResult=function OSF_DDA_AsyncResult(initArgs, errorArgs) {
	OSF.OUtil.defineEnumerableProperties(this, {
		"value": {
			value: initArgs[OSF.DDA.AsyncResultEnum.Properties.Value]
		},
		"status": {
			value: errorArgs ? Microsoft.Office.WebExtension.AsyncResultStatus.Failed : Microsoft.Office.WebExtension.AsyncResultStatus.Succeeded
		}
	});
	if (initArgs[OSF.DDA.AsyncResultEnum.Properties.Context]) {
		OSF.OUtil.defineEnumerableProperty(this, "asyncContext", {
			value: initArgs[OSF.DDA.AsyncResultEnum.Properties.Context]
		});
	}
	if (errorArgs) {
		OSF.OUtil.defineEnumerableProperty(this, "error", {
			value: new OSF.DDA.Error(errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Name], errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Message], errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Code])
		});
	}
};
OSF.DDA.issueAsyncResult=function OSF_DDA$IssueAsyncResult(callArgs, status, payload) {
	var callback=callArgs[Microsoft.Office.WebExtension.Parameters.Callback];
	if (callback) {
		var asyncInitArgs={};
		asyncInitArgs[OSF.DDA.AsyncResultEnum.Properties.Context]=callArgs[Microsoft.Office.WebExtension.Parameters.AsyncContext];
		var errorArgs;
		if (status==OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess) {
			asyncInitArgs[OSF.DDA.AsyncResultEnum.Properties.Value]=payload;
		}
		else {
			errorArgs={};
			payload=payload || OSF.DDA.ErrorCodeManager.getErrorArgs(OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError);
			errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Code]=status || OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError;
			errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Name]=payload.name || payload;
			errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Message]=payload.message || payload;
		}
		callback(new OSF.DDA.AsyncResult(asyncInitArgs, errorArgs));
	}
};
OSF.DDA.SyncMethodNames={};
OSF.DDA.SyncMethodNames.addNames=function (methodNames) {
	for (var entry in methodNames) {
		var am={};
		OSF.OUtil.defineEnumerableProperties(am, {
			"id": {
				value: entry
			},
			"displayName": {
				value: methodNames[entry]
			}
		});
		OSF.DDA.SyncMethodNames[entry]=am;
	}
};
OSF.DDA.SyncMethodCall=function OSF_DDA_SyncMethodCall(requiredParameters, supportedOptions, privateStateCallbacks, checkCallArgs, displayName) {
	var requiredCount=requiredParameters.length;
	var apiMethods=new OSF.DDA.ApiMethodCall(requiredParameters, supportedOptions, privateStateCallbacks, checkCallArgs, displayName);
	function OSF_DAA_SyncMethodCall$ExtractOptions(userArgs, requiredArgs, caller, stateInfo) {
		if (userArgs.length > requiredCount+1) {
			throw OsfMsAjaxFactory.msAjaxError.parameterCount(Strings.OfficeOM.L_TooManyArguments);
		}
		var options, parameterCallback;
		for (var i=userArgs.length - 1; i >=requiredCount; i--) {
			var argument=userArgs[i];
			switch (typeof argument) {
				case "object":
					if (options) {
						throw OsfMsAjaxFactory.msAjaxError.parameterCount(Strings.OfficeOM.L_TooManyOptionalObjects);
					}
					else {
						options=argument;
					}
					break;
				default:
					throw OsfMsAjaxFactory.msAjaxError.argument(Strings.OfficeOM.L_InValidOptionalArgument);
					break;
			}
		}
		options=apiMethods.fillOptions(options, requiredArgs, caller, stateInfo);
		apiMethods.verifyArguments(supportedOptions, options);
		return options;
	}
	;
	this.verifyAndExtractCall=function OSF_DAA_AsyncMethodCall$VerifyAndExtractCall(userArgs, caller, stateInfo) {
		var required=apiMethods.extractRequiredArguments(userArgs, caller, stateInfo);
		var options=OSF_DAA_SyncMethodCall$ExtractOptions(userArgs, required, caller, stateInfo);
		var callArgs=apiMethods.constructCallArgs(required, options, caller, stateInfo);
		return callArgs;
	};
};
OSF.DDA.SyncMethodCallFactory=(function () {
	return {
		manufacture: function (params) {
			var supportedOptions=params.supportedOptions ? OSF.OUtil.createObject(params.supportedOptions) : [];
			return new OSF.DDA.SyncMethodCall(params.requiredArguments || [], supportedOptions, params.privateStateCallbacks, params.checkCallArgs, params.method.displayName);
		}
	};
})();
OSF.DDA.SyncMethodCalls={};
OSF.DDA.SyncMethodCalls.define=function (callDefinition) {
	OSF.DDA.SyncMethodCalls[callDefinition.method.id]=OSF.DDA.SyncMethodCallFactory.manufacture(callDefinition);
};
OSF.DDA.ListType=(function () {
	var listTypes={};
	return {
		setListType: function OSF_DDA_ListType$AddListType(t, prop) { listTypes[t]=prop; },
		isListType: function OSF_DDA_ListType$IsListType(t) { return OSF.OUtil.listContainsKey(listTypes, t); },
		getDescriptor: function OSF_DDA_ListType$getDescriptor(t) { return listTypes[t]; }
	};
})();
OSF.DDA.HostParameterMap=function (specialProcessor, mappings) {
	var toHostMap="toHost";
	var fromHostMap="fromHost";
	var sourceData="sourceData";
	var self="self";
	var dynamicTypes={};
	dynamicTypes[Microsoft.Office.WebExtension.Parameters.Data]={
		toHost: function (data) {
			if (data !=null && data.rows !==undefined) {
				var tableData={};
				tableData[OSF.DDA.TableDataProperties.TableRows]=data.rows;
				tableData[OSF.DDA.TableDataProperties.TableHeaders]=data.headers;
				data=tableData;
			}
			return data;
		},
		fromHost: function (args) {
			return args;
		}
	};
	dynamicTypes[Microsoft.Office.WebExtension.Parameters.SampleData]=dynamicTypes[Microsoft.Office.WebExtension.Parameters.Data];
	function mapValues(preimageSet, mapping) {
		var ret=preimageSet ? {} : undefined;
		for (var entry in preimageSet) {
			var preimage=preimageSet[entry];
			var image;
			if (OSF.DDA.ListType.isListType(entry)) {
				image=[];
				for (var subEntry in preimage) {
					image.push(mapValues(preimage[subEntry], mapping));
				}
			}
			else if (OSF.OUtil.listContainsKey(dynamicTypes, entry)) {
				image=dynamicTypes[entry][mapping](preimage);
			}
			else if (mapping==fromHostMap && specialProcessor.preserveNesting(entry)) {
				image=mapValues(preimage, mapping);
			}
			else {
				var maps=mappings[entry];
				if (maps) {
					var map=maps[mapping];
					if (map) {
						image=map[preimage];
						if (image===undefined) {
							image=preimage;
						}
					}
				}
				else {
					image=preimage;
				}
			}
			ret[entry]=image;
		}
		return ret;
	}
	;
	function generateArguments(imageSet, parameters) {
		var ret;
		for (var param in parameters) {
			var arg;
			if (specialProcessor.isComplexType(param)) {
				arg=generateArguments(imageSet, mappings[param][toHostMap]);
			}
			else {
				arg=imageSet[param];
			}
			if (arg !=undefined) {
				if (!ret) {
					ret={};
				}
				var index=parameters[param];
				if (index==self) {
					index=param;
				}
				ret[index]=specialProcessor.pack(param, arg);
			}
		}
		return ret;
	}
	;
	function extractArguments(source, parameters, extracted) {
		if (!extracted) {
			extracted={};
		}
		for (var param in parameters) {
			var index=parameters[param];
			var value;
			if (index==self) {
				value=source;
			}
			else if (index==sourceData) {
				extracted[param]=source.toArray();
				continue;
			}
			else {
				value=source[index];
			}
			if (value===null || value===undefined) {
				extracted[param]=undefined;
			}
			else {
				value=specialProcessor.unpack(param, value);
				var map;
				if (specialProcessor.isComplexType(param)) {
					map=mappings[param][fromHostMap];
					if (specialProcessor.preserveNesting(param)) {
						extracted[param]=extractArguments(value, map);
					}
					else {
						extractArguments(value, map, extracted);
					}
				}
				else {
					if (OSF.DDA.ListType.isListType(param)) {
						map={};
						var entryDescriptor=OSF.DDA.ListType.getDescriptor(param);
						map[entryDescriptor]=self;
						var extractedValues=new Array(value.length);
						for (var item in value) {
							extractedValues[item]=extractArguments(value[item], map);
						}
						extracted[param]=extractedValues;
					}
					else {
						extracted[param]=value;
					}
				}
			}
		}
		return extracted;
	}
	;
	function applyMap(mapName, preimage, mapping) {
		var parameters=mappings[mapName][mapping];
		var image;
		if (mapping=="toHost") {
			var imageSet=mapValues(preimage, mapping);
			image=generateArguments(imageSet, parameters);
		}
		else if (mapping=="fromHost") {
			var argumentSet=extractArguments(preimage, parameters);
			image=mapValues(argumentSet, mapping);
		}
		return image;
	}
	;
	if (!mappings) {
		mappings={};
	}
	this.addMapping=function (mapName, description) {
		var toHost, fromHost;
		if (description.map) {
			toHost=description.map;
			fromHost={};
			for (var preimage in toHost) {
				var image=toHost[preimage];
				if (image==self) {
					image=preimage;
				}
				fromHost[image]=preimage;
			}
		}
		else {
			toHost=description.toHost;
			fromHost=description.fromHost;
		}
		var pair=mappings[mapName];
		if (pair) {
			var currMap=pair[toHostMap];
			for (var th in currMap)
				toHost[th]=currMap[th];
			currMap=pair[fromHostMap];
			for (var fh in currMap)
				fromHost[fh]=currMap[fh];
		}
		else {
			pair=mappings[mapName]={};
		}
		pair[toHostMap]=toHost;
		pair[fromHostMap]=fromHost;
	};
	this.toHost=function (mapName, preimage) { return applyMap(mapName, preimage, toHostMap); };
	this.fromHost=function (mapName, image) { return applyMap(mapName, image, fromHostMap); };
	this.self=self;
	this.sourceData=sourceData;
	this.addComplexType=function (ct) { specialProcessor.addComplexType(ct); };
	this.getDynamicType=function (dt) { return specialProcessor.getDynamicType(dt); };
	this.setDynamicType=function (dt, handler) { specialProcessor.setDynamicType(dt, handler); };
	this.dynamicTypes=dynamicTypes;
	this.doMapValues=function (preimageSet, mapping) { return mapValues(preimageSet, mapping); };
};
OSF.DDA.SpecialProcessor=function (complexTypes, dynamicTypes) {
	this.addComplexType=function OSF_DDA_SpecialProcessor$addComplexType(ct) {
		complexTypes.push(ct);
	};
	this.getDynamicType=function OSF_DDA_SpecialProcessor$getDynamicType(dt) {
		return dynamicTypes[dt];
	};
	this.setDynamicType=function OSF_DDA_SpecialProcessor$setDynamicType(dt, handler) {
		dynamicTypes[dt]=handler;
	};
	this.isComplexType=function OSF_DDA_SpecialProcessor$isComplexType(t) {
		return OSF.OUtil.listContainsValue(complexTypes, t);
	};
	this.isDynamicType=function OSF_DDA_SpecialProcessor$isDynamicType(p) {
		return OSF.OUtil.listContainsKey(dynamicTypes, p);
	};
	this.preserveNesting=function OSF_DDA_SpecialProcessor$preserveNesting(p) {
		var pn=[];
		if (OSF.DDA.PropertyDescriptors)
			pn.push(OSF.DDA.PropertyDescriptors.Subset);
		if (OSF.DDA.DataNodeEventProperties) {
			pn=pn.concat([
				OSF.DDA.DataNodeEventProperties.OldNode,
				OSF.DDA.DataNodeEventProperties.NewNode,
				OSF.DDA.DataNodeEventProperties.NextSiblingNode
			]);
		}
		return OSF.OUtil.listContainsValue(pn, p);
	};
	this.pack=function OSF_DDA_SpecialProcessor$pack(param, arg) {
		var value;
		if (this.isDynamicType(param)) {
			value=dynamicTypes[param].toHost(arg);
		}
		else {
			value=arg;
		}
		return value;
	};
	this.unpack=function OSF_DDA_SpecialProcessor$unpack(param, arg) {
		var value;
		if (this.isDynamicType(param)) {
			value=dynamicTypes[param].fromHost(arg);
		}
		else {
			value=arg;
		}
		return value;
	};
};
OSF.DDA.getDecoratedParameterMap=function (specialProcessor, initialDefs) {
	var parameterMap=new OSF.DDA.HostParameterMap(specialProcessor);
	var self=parameterMap.self;
	function createObject(properties) {
		var obj=null;
		if (properties) {
			obj={};
			var len=properties.length;
			for (var i=0; i < len; i++) {
				obj[properties[i].name]=properties[i].value;
			}
		}
		return obj;
	}
	parameterMap.define=function define(definition) {
		var args={};
		var toHost=createObject(definition.toHost);
		if (definition.invertible) {
			args.map=toHost;
		}
		else if (definition.canonical) {
			args.toHost=args.fromHost=toHost;
		}
		else {
			args.toHost=toHost;
			args.fromHost=createObject(definition.fromHost);
		}
		parameterMap.addMapping(definition.type, args);
		if (definition.isComplexType)
			parameterMap.addComplexType(definition.type);
	};
	for (var id in initialDefs)
		parameterMap.define(initialDefs[id]);
	return parameterMap;
};
OSF.OUtil.setNamespace("DispIdHost", OSF.DDA);
OSF.DDA.DispIdHost.Methods={
	InvokeMethod: "invokeMethod",
	AddEventHandler: "addEventHandler",
	RemoveEventHandler: "removeEventHandler",
	OpenDialog: "openDialog",
	CloseDialog: "closeDialog",
	MessageParent: "messageParent",
	SendMessage: "sendMessage"
};
OSF.DDA.DispIdHost.Delegates={
	ExecuteAsync: "executeAsync",
	RegisterEventAsync: "registerEventAsync",
	UnregisterEventAsync: "unregisterEventAsync",
	ParameterMap: "parameterMap",
	OpenDialog: "openDialog",
	CloseDialog: "closeDialog",
	MessageParent: "messageParent",
	SendMessage: "sendMessage"
};
OSF.DDA.DispIdHost.Facade=function OSF_DDA_DispIdHost_Facade(getDelegateMethods, parameterMap) {
	var dispIdMap={};
	var jsom=OSF.DDA.AsyncMethodNames;
	var did=OSF.DDA.MethodDispId;
	var methodMap={
		"GoToByIdAsync": did.dispidNavigateToMethod,
		"GetSelectedDataAsync": did.dispidGetSelectedDataMethod,
		"SetSelectedDataAsync": did.dispidSetSelectedDataMethod,
		"GetDocumentCopyChunkAsync": did.dispidGetDocumentCopyChunkMethod,
		"ReleaseDocumentCopyAsync": did.dispidReleaseDocumentCopyMethod,
		"GetDocumentCopyAsync": did.dispidGetDocumentCopyMethod,
		"AddFromSelectionAsync": did.dispidAddBindingFromSelectionMethod,
		"AddFromPromptAsync": did.dispidAddBindingFromPromptMethod,
		"AddFromNamedItemAsync": did.dispidAddBindingFromNamedItemMethod,
		"GetAllAsync": did.dispidGetAllBindingsMethod,
		"GetByIdAsync": did.dispidGetBindingMethod,
		"ReleaseByIdAsync": did.dispidReleaseBindingMethod,
		"GetDataAsync": did.dispidGetBindingDataMethod,
		"SetDataAsync": did.dispidSetBindingDataMethod,
		"AddRowsAsync": did.dispidAddRowsMethod,
		"AddColumnsAsync": did.dispidAddColumnsMethod,
		"DeleteAllDataValuesAsync": did.dispidClearAllRowsMethod,
		"RefreshAsync": did.dispidLoadSettingsMethod,
		"SaveAsync": did.dispidSaveSettingsMethod,
		"GetActiveViewAsync": did.dispidGetActiveViewMethod,
		"GetFilePropertiesAsync": did.dispidGetFilePropertiesMethod,
		"GetOfficeThemeAsync": did.dispidGetOfficeThemeMethod,
		"GetDocumentThemeAsync": did.dispidGetDocumentThemeMethod,
		"ClearFormatsAsync": did.dispidClearFormatsMethod,
		"SetTableOptionsAsync": did.dispidSetTableOptionsMethod,
		"SetFormatsAsync": did.dispidSetFormatsMethod,
		"GetAccessTokenAsync": did.dispidGetAccessTokenMethod,
		"ExecuteRichApiRequestAsync": did.dispidExecuteRichApiRequestMethod,
		"AppCommandInvocationCompletedAsync": did.dispidAppCommandInvocationCompletedMethod,
		"CloseContainerAsync": did.dispidCloseContainerMethod,
		"OpenBrowserWindow": did.dispidOpenBrowserWindow,
		"CreateDocumentAsync": did.dispidCreateDocumentMethod,
		"InsertFormAsync": did.dispidInsertFormMethod,
		"ExecuteFeature": did.dispidExecuteFeature,
		"QueryFeature": did.dispidQueryFeature,
		"AddDataPartAsync": did.dispidAddDataPartMethod,
		"GetDataPartByIdAsync": did.dispidGetDataPartByIdMethod,
		"GetDataPartsByNameSpaceAsync": did.dispidGetDataPartsByNamespaceMethod,
		"GetPartXmlAsync": did.dispidGetDataPartXmlMethod,
		"GetPartNodesAsync": did.dispidGetDataPartNodesMethod,
		"DeleteDataPartAsync": did.dispidDeleteDataPartMethod,
		"GetNodeValueAsync": did.dispidGetDataNodeValueMethod,
		"GetNodeXmlAsync": did.dispidGetDataNodeXmlMethod,
		"GetRelativeNodesAsync": did.dispidGetDataNodesMethod,
		"SetNodeValueAsync": did.dispidSetDataNodeValueMethod,
		"SetNodeXmlAsync": did.dispidSetDataNodeXmlMethod,
		"AddDataPartNamespaceAsync": did.dispidAddDataNamespaceMethod,
		"GetDataPartNamespaceAsync": did.dispidGetDataUriByPrefixMethod,
		"GetDataPartPrefixAsync": did.dispidGetDataPrefixByUriMethod,
		"GetNodeTextAsync": did.dispidGetDataNodeTextMethod,
		"SetNodeTextAsync": did.dispidSetDataNodeTextMethod,
		"GetSelectedTask": did.dispidGetSelectedTaskMethod,
		"GetTask": did.dispidGetTaskMethod,
		"GetWSSUrl": did.dispidGetWSSUrlMethod,
		"GetTaskField": did.dispidGetTaskFieldMethod,
		"GetSelectedResource": did.dispidGetSelectedResourceMethod,
		"GetResourceField": did.dispidGetResourceFieldMethod,
		"GetProjectField": did.dispidGetProjectFieldMethod,
		"GetSelectedView": did.dispidGetSelectedViewMethod,
		"GetTaskByIndex": did.dispidGetTaskByIndexMethod,
		"GetResourceByIndex": did.dispidGetResourceByIndexMethod,
		"SetTaskField": did.dispidSetTaskFieldMethod,
		"SetResourceField": did.dispidSetResourceFieldMethod,
		"GetMaxTaskIndex": did.dispidGetMaxTaskIndexMethod,
		"GetMaxResourceIndex": did.dispidGetMaxResourceIndexMethod,
		"CreateTask": did.dispidCreateTaskMethod
	};
	for (var method in methodMap) {
		if (jsom[method]) {
			dispIdMap[jsom[method].id]=methodMap[method];
		}
	}
	jsom=OSF.DDA.SyncMethodNames;
	did=OSF.DDA.MethodDispId;
	var syncMethodMap={
		"MessageParent": did.dispidMessageParentMethod,
		"SendMessage": did.dispidSendMessageMethod
	};
	for (var method in syncMethodMap) {
		if (jsom[method]) {
			dispIdMap[jsom[method].id]=syncMethodMap[method];
		}
	}
	jsom=Microsoft.Office.WebExtension.EventType;
	did=OSF.DDA.EventDispId;
	var eventMap={
		"SettingsChanged": did.dispidSettingsChangedEvent,
		"DocumentSelectionChanged": did.dispidDocumentSelectionChangedEvent,
		"BindingSelectionChanged": did.dispidBindingSelectionChangedEvent,
		"BindingDataChanged": did.dispidBindingDataChangedEvent,
		"ActiveViewChanged": did.dispidActiveViewChangedEvent,
		"OfficeThemeChanged": did.dispidOfficeThemeChangedEvent,
		"DocumentThemeChanged": did.dispidDocumentThemeChangedEvent,
		"AppCommandInvoked": did.dispidAppCommandInvokedEvent,
		"DialogMessageReceived": did.dispidDialogMessageReceivedEvent,
		"DialogParentMessageReceived": did.dispidDialogParentMessageReceivedEvent,
		"ObjectDeleted": did.dispidObjectDeletedEvent,
		"ObjectSelectionChanged": did.dispidObjectSelectionChangedEvent,
		"ObjectDataChanged": did.dispidObjectDataChangedEvent,
		"ContentControlAdded": did.dispidContentControlAddedEvent,
		"RichApiMessage": did.dispidRichApiMessageEvent,
		"ItemChanged": did.dispidOlkItemSelectedChangedEvent,
		"RecipientsChanged": did.dispidOlkRecipientsChangedEvent,
		"AppointmentTimeChanged": did.dispidOlkAppointmentTimeChangedEvent,
		"RecurrenceChanged": did.dispidOlkRecurrenceChangedEvent,
		"AttachmentsChanged": did.dispidOlkAttachmentsChangedEvent,
		"EnhancedLocationsChanged": did.dispidOlkEnhancedLocationsChangedEvent,
		"InfobarClicked": did.dispidOlkInfobarClickedEvent,
		"TaskSelectionChanged": did.dispidTaskSelectionChangedEvent,
		"ResourceSelectionChanged": did.dispidResourceSelectionChangedEvent,
		"ViewSelectionChanged": did.dispidViewSelectionChangedEvent,
		"DataNodeInserted": did.dispidDataNodeAddedEvent,
		"DataNodeReplaced": did.dispidDataNodeReplacedEvent,
		"DataNodeDeleted": did.dispidDataNodeDeletedEvent
	};
	for (var event in eventMap) {
		if (jsom[event]) {
			dispIdMap[jsom[event]]=eventMap[event];
		}
	}
	function IsObjectEvent(dispId) {
		return (dispId==OSF.DDA.EventDispId.dispidObjectDeletedEvent ||
			dispId==OSF.DDA.EventDispId.dispidObjectSelectionChangedEvent ||
			dispId==OSF.DDA.EventDispId.dispidObjectDataChangedEvent ||
			dispId==OSF.DDA.EventDispId.dispidContentControlAddedEvent);
	}
	function onException(ex, asyncMethodCall, suppliedArgs, callArgs) {
		if (typeof ex=="number") {
			if (!callArgs) {
				callArgs=asyncMethodCall.getCallArgs(suppliedArgs);
			}
			OSF.DDA.issueAsyncResult(callArgs, ex, OSF.DDA.ErrorCodeManager.getErrorArgs(ex));
		}
		else {
			throw ex;
		}
	}
	;
	this[OSF.DDA.DispIdHost.Methods.InvokeMethod]=function OSF_DDA_DispIdHost_Facade$InvokeMethod(method, suppliedArguments, caller, privateState) {
		var callArgs;
		try {
			var methodName=method.id;
			var asyncMethodCall=OSF.DDA.AsyncMethodCalls[methodName];
			callArgs=asyncMethodCall.verifyAndExtractCall(suppliedArguments, caller, privateState);
			var dispId=dispIdMap[methodName];
			var delegate=getDelegateMethods(methodName);
			var richApiInExcelMethodSubstitution=null;
			if (window.Excel && window.Office.context.requirements.isSetSupported("RedirectV1Api")) {
				window.Excel._RedirectV1APIs=true;
			}
			if (window.Excel && window.Excel._RedirectV1APIs && (richApiInExcelMethodSubstitution=window.Excel._V1APIMap[methodName])) {
				var preprocessedCallArgs=OSF.OUtil.shallowCopy(callArgs);
				delete preprocessedCallArgs[Microsoft.Office.WebExtension.Parameters.AsyncContext];
				if (richApiInExcelMethodSubstitution.preprocess) {
					preprocessedCallArgs=richApiInExcelMethodSubstitution.preprocess(preprocessedCallArgs);
				}
				var ctx=new window.Excel.RequestContext();
				var result=richApiInExcelMethodSubstitution.call(ctx, preprocessedCallArgs);
				ctx.sync()
					.then(function () {
					var response=result.value;
					var status=response.status;
					delete response["status"];
					delete response["@odata.type"];
					if (richApiInExcelMethodSubstitution.postprocess) {
						response=richApiInExcelMethodSubstitution.postprocess(response, preprocessedCallArgs);
					}
					if (status !=0) {
						response=OSF.DDA.ErrorCodeManager.getErrorArgs(status);
					}
					OSF.DDA.issueAsyncResult(callArgs, status, response);
				})["catch"](function (error) {
					OSF.DDA.issueAsyncResult(callArgs, OSF.DDA.ErrorCodeManager.errorCodes.ooeFailure, null);
				});
			}
			else {
				var hostCallArgs;
				if (parameterMap.toHost) {
					hostCallArgs=parameterMap.toHost(dispId, callArgs);
				}
				else {
					hostCallArgs=callArgs;
				}
				var startTime=(new Date()).getTime();
				delegate[OSF.DDA.DispIdHost.Delegates.ExecuteAsync]({
					"dispId": dispId,
					"hostCallArgs": hostCallArgs,
					"onCalling": function OSF_DDA_DispIdFacade$Execute_onCalling() { },
					"onReceiving": function OSF_DDA_DispIdFacade$Execute_onReceiving() { },
					"onComplete": function (status, hostResponseArgs) {
						var responseArgs;
						if (status==OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess) {
							if (parameterMap.fromHost) {
								responseArgs=parameterMap.fromHost(dispId, hostResponseArgs);
							}
							else {
								responseArgs=hostResponseArgs;
							}
						}
						else {
							responseArgs=hostResponseArgs;
						}
						var payload=asyncMethodCall.processResponse(status, responseArgs, caller, callArgs);
						OSF.DDA.issueAsyncResult(callArgs, status, payload);
						if (OSF.AppTelemetry) {
							OSF.AppTelemetry.onMethodDone(dispId, hostCallArgs, Math.abs((new Date()).getTime() - startTime), status);
						}
					}
				});
			}
		}
		catch (ex) {
			onException(ex, asyncMethodCall, suppliedArguments, callArgs);
		}
	};
	this[OSF.DDA.DispIdHost.Methods.AddEventHandler]=function OSF_DDA_DispIdHost_Facade$AddEventHandler(suppliedArguments, eventDispatch, caller, isPopupWindow) {
		var callArgs;
		var eventType, handler;
		var isObjectEvent=false;
		function onEnsureRegistration(status) {
			if (status==OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess) {
				var added=!isObjectEvent ? eventDispatch.addEventHandler(eventType, handler) :
					eventDispatch.addObjectEventHandler(eventType, callArgs[Microsoft.Office.WebExtension.Parameters.Id], handler);
				if (!added) {
					status=OSF.DDA.ErrorCodeManager.errorCodes.ooeEventHandlerAdditionFailed;
				}
			}
			var error;
			if (status !=OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess) {
				error=OSF.DDA.ErrorCodeManager.getErrorArgs(status);
			}
			OSF.DDA.issueAsyncResult(callArgs, status, error);
		}
		try {
			var asyncMethodCall=OSF.DDA.AsyncMethodCalls[OSF.DDA.AsyncMethodNames.AddHandlerAsync.id];
			callArgs=asyncMethodCall.verifyAndExtractCall(suppliedArguments, caller, eventDispatch);
			eventType=callArgs[Microsoft.Office.WebExtension.Parameters.EventType];
			handler=callArgs[Microsoft.Office.WebExtension.Parameters.Handler];
			if (isPopupWindow) {
				onEnsureRegistration(OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess);
				return;
			}
			var dispId=dispIdMap[eventType];
			isObjectEvent=IsObjectEvent(dispId);
			var targetId=(isObjectEvent ? callArgs[Microsoft.Office.WebExtension.Parameters.Id] : (caller.id || ""));
			var count=isObjectEvent ? eventDispatch.getObjectEventHandlerCount(eventType, targetId) : eventDispatch.getEventHandlerCount(eventType);
			if (count==0) {
				var invoker=getDelegateMethods(eventType)[OSF.DDA.DispIdHost.Delegates.RegisterEventAsync];
				invoker({
					"eventType": eventType,
					"dispId": dispId,
					"targetId": targetId,
					"onCalling": function OSF_DDA_DispIdFacade$Execute_onCalling() { OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.IssueCall); },
					"onReceiving": function OSF_DDA_DispIdFacade$Execute_onReceiving() { OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.ReceiveResponse); },
					"onComplete": onEnsureRegistration,
					"onEvent": function handleEvent(hostArgs) {
						var args=parameterMap.fromHost(dispId, hostArgs);
						if (!isObjectEvent)
							eventDispatch.fireEvent(OSF.DDA.OMFactory.manufactureEventArgs(eventType, caller, args));
						else
							eventDispatch.fireObjectEvent(targetId, OSF.DDA.OMFactory.manufactureEventArgs(eventType, targetId, args));
					}
				});
			}
			else {
				onEnsureRegistration(OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess);
			}
		}
		catch (ex) {
			onException(ex, asyncMethodCall, suppliedArguments, callArgs);
		}
	};
	this[OSF.DDA.DispIdHost.Methods.RemoveEventHandler]=function OSF_DDA_DispIdHost_Facade$RemoveEventHandler(suppliedArguments, eventDispatch, caller) {
		var callArgs;
		var eventType, handler;
		var isObjectEvent=false;
		function onEnsureRegistration(status) {
			var error;
			if (status !=OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess) {
				error=OSF.DDA.ErrorCodeManager.getErrorArgs(status);
			}
			OSF.DDA.issueAsyncResult(callArgs, status, error);
		}
		try {
			var asyncMethodCall=OSF.DDA.AsyncMethodCalls[OSF.DDA.AsyncMethodNames.RemoveHandlerAsync.id];
			callArgs=asyncMethodCall.verifyAndExtractCall(suppliedArguments, caller, eventDispatch);
			eventType=callArgs[Microsoft.Office.WebExtension.Parameters.EventType];
			handler=callArgs[Microsoft.Office.WebExtension.Parameters.Handler];
			var dispId=dispIdMap[eventType];
			isObjectEvent=IsObjectEvent(dispId);
			var targetId=(isObjectEvent ? callArgs[Microsoft.Office.WebExtension.Parameters.Id] : (caller.id || ""));
			var status, removeSuccess;
			if (handler===null) {
				removeSuccess=isObjectEvent ? eventDispatch.clearObjectEventHandlers(eventType, targetId) : eventDispatch.clearEventHandlers(eventType);
				status=OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess;
			}
			else {
				removeSuccess=isObjectEvent ? eventDispatch.removeObjectEventHandler(eventType, targetId, handler) : eventDispatch.removeEventHandler(eventType, handler);
				status=removeSuccess ? OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess : OSF.DDA.ErrorCodeManager.errorCodes.ooeEventHandlerNotExist;
			}
			var count=isObjectEvent ? eventDispatch.getObjectEventHandlerCount(eventType, targetId) : eventDispatch.getEventHandlerCount(eventType);
			if (removeSuccess && count==0) {
				var invoker=getDelegateMethods(eventType)[OSF.DDA.DispIdHost.Delegates.UnregisterEventAsync];
				invoker({
					"eventType": eventType,
					"dispId": dispId,
					"targetId": targetId,
					"onCalling": function OSF_DDA_DispIdFacade$Execute_onCalling() { OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.IssueCall); },
					"onReceiving": function OSF_DDA_DispIdFacade$Execute_onReceiving() { OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.ReceiveResponse); },
					"onComplete": onEnsureRegistration
				});
			}
			else {
				onEnsureRegistration(status);
			}
		}
		catch (ex) {
			onException(ex, asyncMethodCall, suppliedArguments, callArgs);
		}
	};
	this[OSF.DDA.DispIdHost.Methods.OpenDialog]=function OSF_DDA_DispIdHost_Facade$OpenDialog(suppliedArguments, eventDispatch, caller) {
		var callArgs;
		var targetId;
		var dialogMessageEvent=Microsoft.Office.WebExtension.EventType.DialogMessageReceived;
		var dialogOtherEvent=Microsoft.Office.WebExtension.EventType.DialogEventReceived;
		function onEnsureRegistration(status) {
			var payload;
			if (status !=OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess) {
				payload=OSF.DDA.ErrorCodeManager.getErrorArgs(status);
			}
			else {
				var onSucceedArgs={};
				onSucceedArgs[Microsoft.Office.WebExtension.Parameters.Id]=targetId;
				onSucceedArgs[Microsoft.Office.WebExtension.Parameters.Data]=eventDispatch;
				var payload=asyncMethodCall.processResponse(status, onSucceedArgs, caller, callArgs);
				OSF.DialogShownStatus.hasDialogShown=true;
				eventDispatch.clearEventHandlers(dialogMessageEvent);
				eventDispatch.clearEventHandlers(dialogOtherEvent);
			}
			OSF.DDA.issueAsyncResult(callArgs, status, payload);
		}
		try {
			if (dialogMessageEvent==undefined || dialogOtherEvent==undefined) {
				onEnsureRegistration(OSF.DDA.ErrorCodeManager.ooeOperationNotSupported);
			}
			if (OSF.DDA.AsyncMethodNames.DisplayDialogAsync==null) {
				onEnsureRegistration(OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError);
				return;
			}
			var asyncMethodCall=OSF.DDA.AsyncMethodCalls[OSF.DDA.AsyncMethodNames.DisplayDialogAsync.id];
			callArgs=asyncMethodCall.verifyAndExtractCall(suppliedArguments, caller, eventDispatch);
			var dispId=dispIdMap[dialogMessageEvent];
			var delegateMethods=getDelegateMethods(dialogMessageEvent);
			var invoker=delegateMethods[OSF.DDA.DispIdHost.Delegates.OpenDialog] !=undefined ?
				delegateMethods[OSF.DDA.DispIdHost.Delegates.OpenDialog] :
				delegateMethods[OSF.DDA.DispIdHost.Delegates.RegisterEventAsync];
			targetId=JSON.stringify(callArgs);
			if (!OSF.DialogShownStatus.hasDialogShown) {
				eventDispatch.clearQueuedEvent(dialogMessageEvent);
				eventDispatch.clearQueuedEvent(dialogOtherEvent);
				eventDispatch.clearQueuedEvent(Microsoft.Office.WebExtension.EventType.DialogParentMessageReceived);
			}
			invoker({
				"eventType": dialogMessageEvent,
				"dispId": dispId,
				"targetId": targetId,
				"onCalling": function OSF_DDA_DispIdFacade$Execute_onCalling() { OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.IssueCall); },
				"onReceiving": function OSF_DDA_DispIdFacade$Execute_onReceiving() { OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.ReceiveResponse); },
				"onComplete": onEnsureRegistration,
				"onEvent": function handleEvent(hostArgs) {
					var args=parameterMap.fromHost(dispId, hostArgs);
					var event=OSF.DDA.OMFactory.manufactureEventArgs(dialogMessageEvent, caller, args);
					if (event.type==dialogOtherEvent) {
						var payload=OSF.DDA.ErrorCodeManager.getErrorArgs(event.error);
						var errorArgs={};
						errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Code]=status || OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError;
						errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Name]=payload.name || payload;
						errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Message]=payload.message || payload;
						event.error=new OSF.DDA.Error(errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Name], errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Message], errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Code]);
					}
					eventDispatch.fireOrQueueEvent(event);
					if (args[OSF.DDA.PropertyDescriptors.MessageType]==OSF.DialogMessageType.DialogClosed) {
						eventDispatch.clearEventHandlers(dialogMessageEvent);
						eventDispatch.clearEventHandlers(dialogOtherEvent);
						eventDispatch.clearEventHandlers(Microsoft.Office.WebExtension.EventType.DialogParentMessageReceived);
						OSF.DialogShownStatus.hasDialogShown=false;
					}
				}
			});
		}
		catch (ex) {
			onException(ex, asyncMethodCall, suppliedArguments, callArgs);
		}
	};
	this[OSF.DDA.DispIdHost.Methods.CloseDialog]=function OSF_DDA_DispIdHost_Facade$CloseDialog(suppliedArguments, targetId, eventDispatch, caller) {
		var callArgs;
		var dialogMessageEvent, dialogOtherEvent;
		var closeStatus=OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess;
		function closeCallback(status) {
			closeStatus=status;
			OSF.DialogShownStatus.hasDialogShown=false;
		}
		try {
			var asyncMethodCall=OSF.DDA.AsyncMethodCalls[OSF.DDA.AsyncMethodNames.CloseAsync.id];
			callArgs=asyncMethodCall.verifyAndExtractCall(suppliedArguments, caller, eventDispatch);
			dialogMessageEvent=Microsoft.Office.WebExtension.EventType.DialogMessageReceived;
			dialogOtherEvent=Microsoft.Office.WebExtension.EventType.DialogEventReceived;
			eventDispatch.clearEventHandlers(dialogMessageEvent);
			eventDispatch.clearEventHandlers(dialogOtherEvent);
			var dispId=dispIdMap[dialogMessageEvent];
			var delegateMethods=getDelegateMethods(dialogMessageEvent);
			var invoker=delegateMethods[OSF.DDA.DispIdHost.Delegates.CloseDialog] !=undefined ?
				delegateMethods[OSF.DDA.DispIdHost.Delegates.CloseDialog] :
				delegateMethods[OSF.DDA.DispIdHost.Delegates.UnregisterEventAsync];
			invoker({
				"eventType": dialogMessageEvent,
				"dispId": dispId,
				"targetId": targetId,
				"onCalling": function OSF_DDA_DispIdFacade$Execute_onCalling() { OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.IssueCall); },
				"onReceiving": function OSF_DDA_DispIdFacade$Execute_onReceiving() { OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.ReceiveResponse); },
				"onComplete": closeCallback
			});
		}
		catch (ex) {
			onException(ex, asyncMethodCall, suppliedArguments, callArgs);
		}
		if (closeStatus !=OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess) {
			throw OSF.OUtil.formatString(Strings.OfficeOM.L_FunctionCallFailed, OSF.DDA.AsyncMethodNames.CloseAsync.displayName, closeStatus);
		}
	};
	this[OSF.DDA.DispIdHost.Methods.MessageParent]=function OSF_DDA_DispIdHost_Facade$MessageParent(suppliedArguments, caller) {
		var stateInfo={};
		var syncMethodCall=OSF.DDA.SyncMethodCalls[OSF.DDA.SyncMethodNames.MessageParent.id];
		var callArgs=syncMethodCall.verifyAndExtractCall(suppliedArguments, caller, stateInfo);
		var delegate=getDelegateMethods(OSF.DDA.SyncMethodNames.MessageParent.id);
		var invoker=delegate[OSF.DDA.DispIdHost.Delegates.MessageParent];
		var dispId=dispIdMap[OSF.DDA.SyncMethodNames.MessageParent.id];
		return invoker({
			"dispId": dispId,
			"hostCallArgs": callArgs,
			"onCalling": function OSF_DDA_DispIdFacade$Execute_onCalling() { OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.IssueCall); },
			"onReceiving": function OSF_DDA_DispIdFacade$Execute_onReceiving() { OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.ReceiveResponse); }
		});
	};
	this[OSF.DDA.DispIdHost.Methods.SendMessage]=function OSF_DDA_DispIdHost_Facade$SendMessage(suppliedArguments, eventDispatch, caller) {
		var stateInfo={};
		var syncMethodCall=OSF.DDA.SyncMethodCalls[OSF.DDA.SyncMethodNames.SendMessage.id];
		var callArgs=syncMethodCall.verifyAndExtractCall(suppliedArguments, caller, stateInfo);
		var delegate=getDelegateMethods(OSF.DDA.SyncMethodNames.SendMessage.id);
		var invoker=delegate[OSF.DDA.DispIdHost.Delegates.SendMessage];
		var dispId=dispIdMap[OSF.DDA.SyncMethodNames.SendMessage.id];
		return invoker({
			"dispId": dispId,
			"hostCallArgs": callArgs,
			"onCalling": function OSF_DDA_DispIdFacade$Execute_onCalling() { OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.IssueCall); },
			"onReceiving": function OSF_DDA_DispIdFacade$Execute_onReceiving() { OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.ReceiveResponse); }
		});
	};
};
OSF.DDA.DispIdHost.addAsyncMethods=function OSF_DDA_DispIdHost$AddAsyncMethods(target, asyncMethodNames, privateState) {
	for (var entry in asyncMethodNames) {
		var method=asyncMethodNames[entry];
		var name=method.displayName;
		if (!target[name]) {
			OSF.OUtil.defineEnumerableProperty(target, name, {
				value: (function (asyncMethod) {
					return function () {
						var invokeMethod=OSF._OfficeAppFactory.getHostFacade()[OSF.DDA.DispIdHost.Methods.InvokeMethod];
						invokeMethod(asyncMethod, arguments, target, privateState);
					};
				})(method)
			});
		}
	}
};
OSF.DDA.DispIdHost.addEventSupport=function OSF_DDA_DispIdHost$AddEventSupport(target, eventDispatch, isPopupWindow) {
	var add=OSF.DDA.AsyncMethodNames.AddHandlerAsync.displayName;
	var remove=OSF.DDA.AsyncMethodNames.RemoveHandlerAsync.displayName;
	if (!target[add]) {
		OSF.OUtil.defineEnumerableProperty(target, add, {
			value: function () {
				var addEventHandler=OSF._OfficeAppFactory.getHostFacade()[OSF.DDA.DispIdHost.Methods.AddEventHandler];
				addEventHandler(arguments, eventDispatch, target, isPopupWindow);
			}
		});
	}
	if (!target[remove]) {
		OSF.OUtil.defineEnumerableProperty(target, remove, {
			value: function () {
				var removeEventHandler=OSF._OfficeAppFactory.getHostFacade()[OSF.DDA.DispIdHost.Methods.RemoveEventHandler];
				removeEventHandler(arguments, eventDispatch, target);
			}
		});
	}
};
var OfficeExt;
(function (OfficeExt) {
	var MsAjaxTypeHelper=(function () {
		function MsAjaxTypeHelper() {
		}
		MsAjaxTypeHelper.isInstanceOfType=function (type, instance) {
			if (typeof (instance)==="undefined" || instance===null)
				return false;
			if (instance instanceof type)
				return true;
			var instanceType=instance.constructor;
			if (!instanceType || (typeof (instanceType) !=="function") || !instanceType.__typeName || instanceType.__typeName==='Object') {
				instanceType=Object;
			}
			return !!(instanceType===type) ||
				(instanceType.__typeName && type.__typeName && instanceType.__typeName===type.__typeName);
		};
		return MsAjaxTypeHelper;
	})();
	OfficeExt.MsAjaxTypeHelper=MsAjaxTypeHelper;
	var MsAjaxError=(function () {
		function MsAjaxError() {
		}
		MsAjaxError.create=function (message, errorInfo) {
			var err=new Error(message);
			err.message=message;
			if (errorInfo) {
				for (var v in errorInfo) {
					err[v]=errorInfo[v];
				}
			}
			err.popStackFrame();
			return err;
		};
		MsAjaxError.parameterCount=function (message) {
			var displayMessage="Sys.ParameterCountException: "+(message ? message : "Parameter count mismatch.");
			var err=MsAjaxError.create(displayMessage, { name: 'Sys.ParameterCountException' });
			err.popStackFrame();
			return err;
		};
		MsAjaxError.argument=function (paramName, message) {
			var displayMessage="Sys.ArgumentException: "+(message ? message : "Value does not fall within the expected range.");
			if (paramName) {
				displayMessage+="\n"+MsAjaxString.format("Parameter name: {0}", paramName);
			}
			var err=MsAjaxError.create(displayMessage, { name: "Sys.ArgumentException", paramName: paramName });
			err.popStackFrame();
			return err;
		};
		MsAjaxError.argumentNull=function (paramName, message) {
			var displayMessage="Sys.ArgumentNullException: "+(message ? message : "Value cannot be null.");
			if (paramName) {
				displayMessage+="\n"+MsAjaxString.format("Parameter name: {0}", paramName);
			}
			var err=MsAjaxError.create(displayMessage, { name: "Sys.ArgumentNullException", paramName: paramName });
			err.popStackFrame();
			return err;
		};
		MsAjaxError.argumentOutOfRange=function (paramName, actualValue, message) {
			var displayMessage="Sys.ArgumentOutOfRangeException: "+(message ? message : "Specified argument was out of the range of valid values.");
			if (paramName) {
				displayMessage+="\n"+MsAjaxString.format("Parameter name: {0}", paramName);
			}
			if (typeof (actualValue) !=="undefined" && actualValue !==null) {
				displayMessage+="\n"+MsAjaxString.format("Actual value was {0}.", actualValue);
			}
			var err=MsAjaxError.create(displayMessage, {
				name: "Sys.ArgumentOutOfRangeException",
				paramName: paramName,
				actualValue: actualValue
			});
			err.popStackFrame();
			return err;
		};
		MsAjaxError.argumentType=function (paramName, actualType, expectedType, message) {
			var displayMessage="Sys.ArgumentTypeException: ";
			if (message) {
				displayMessage+=message;
			}
			else if (actualType && expectedType) {
				displayMessage+=MsAjaxString.format("Object of type '{0}' cannot be converted to type '{1}'.", actualType.getName ? actualType.getName() : actualType, expectedType.getName ? expectedType.getName() : expectedType);
			}
			else {
				displayMessage+="Object cannot be converted to the required type.";
			}
			if (paramName) {
				displayMessage+="\n"+MsAjaxString.format("Parameter name: {0}", paramName);
			}
			var err=MsAjaxError.create(displayMessage, {
				name: "Sys.ArgumentTypeException",
				paramName: paramName,
				actualType: actualType,
				expectedType: expectedType
			});
			err.popStackFrame();
			return err;
		};
		MsAjaxError.argumentUndefined=function (paramName, message) {
			var displayMessage="Sys.ArgumentUndefinedException: "+(message ? message : "Value cannot be undefined.");
			if (paramName) {
				displayMessage+="\n"+MsAjaxString.format("Parameter name: {0}", paramName);
			}
			var err=MsAjaxError.create(displayMessage, { name: "Sys.ArgumentUndefinedException", paramName: paramName });
			err.popStackFrame();
			return err;
		};
		MsAjaxError.invalidOperation=function (message) {
			var displayMessage="Sys.InvalidOperationException: "+(message ? message : "Operation is not valid due to the current state of the object.");
			var err=MsAjaxError.create(displayMessage, { name: 'Sys.InvalidOperationException' });
			err.popStackFrame();
			return err;
		};
		return MsAjaxError;
	})();
	OfficeExt.MsAjaxError=MsAjaxError;
	var MsAjaxString=(function () {
		function MsAjaxString() {
		}
		MsAjaxString.format=function (format) {
			var args=[];
			for (var _i=1; _i < arguments.length; _i++) {
				args[_i - 1]=arguments[_i];
			}
			var source=format;
			return source.replace(/{(\d+)}/gm, function (match, number) {
				var index=parseInt(number, 10);
				return args[index]===undefined ? '{'+number+'}' : args[index];
			});
		};
		MsAjaxString.startsWith=function (str, prefix) {
			return (str.substr(0, prefix.length)===prefix);
		};
		return MsAjaxString;
	})();
	OfficeExt.MsAjaxString=MsAjaxString;
	var MsAjaxDebug=(function () {
		function MsAjaxDebug() {
		}
		MsAjaxDebug.trace=function (text) {
			if (typeof Debug !=="undefined" && Debug.writeln)
				Debug.writeln(text);
			if (window.console && window.console.log)
				window.console.log(text);
			if (window.opera && window.opera.postError)
				window.opera.postError(text);
			if (window.debugService && window.debugService.trace)
				window.debugService.trace(text);
			var a=document.getElementById("TraceConsole");
			if (a && a.tagName.toUpperCase()==="TEXTAREA") {
				a.innerHTML+=text+"\n";
			}
		};
		return MsAjaxDebug;
	})();
	OfficeExt.MsAjaxDebug=MsAjaxDebug;
	if (!OsfMsAjaxFactory.isMsAjaxLoaded()) {
		var registerTypeInternal=function registerTypeInternal(type, name, isClass) {
			if (type.__typeName===undefined || type.__typeName===null) {
				type.__typeName=name;
			}
			if (type.__class===undefined || type.__class===null) {
				type.__class=isClass;
			}
		};
		registerTypeInternal(Function, "Function", true);
		registerTypeInternal(Error, "Error", true);
		registerTypeInternal(Object, "Object", true);
		registerTypeInternal(String, "String", true);
		registerTypeInternal(Boolean, "Boolean", true);
		registerTypeInternal(Date, "Date", true);
		registerTypeInternal(Number, "Number", true);
		registerTypeInternal(RegExp, "RegExp", true);
		registerTypeInternal(Array, "Array", true);
		if (!Function.createCallback) {
			Function.createCallback=function Function$createCallback(method, context) {
				var e=Function._validateParams(arguments, [
					{ name: "method", type: Function },
					{ name: "context", mayBeNull: true }
				]);
				if (e)
					throw e;
				return function () {
					var l=arguments.length;
					if (l > 0) {
						var args=[];
						for (var i=0; i < l; i++) {
							args[i]=arguments[i];
						}
						args[l]=context;
						return method.apply(this, args);
					}
					return method.call(this, context);
				};
			};
		}
		if (!Function.createDelegate) {
			Function.createDelegate=function Function$createDelegate(instance, method) {
				var e=Function._validateParams(arguments, [
					{ name: "instance", mayBeNull: true },
					{ name: "method", type: Function }
				]);
				if (e)
					throw e;
				return function () {
					return method.apply(instance, arguments);
				};
			};
		}
		if (!Function._validateParams) {
			Function._validateParams=function (params, expectedParams, validateParameterCount) {
				var e, expectedLength=expectedParams.length;
				validateParameterCount=validateParameterCount || (typeof (validateParameterCount)==="undefined");
				e=Function._validateParameterCount(params, expectedParams, validateParameterCount);
				if (e) {
					e.popStackFrame();
					return e;
				}
				for (var i=0, l=params.length; i < l; i++) {
					var expectedParam=expectedParams[Math.min(i, expectedLength - 1)], paramName=expectedParam.name;
					if (expectedParam.parameterArray) {
						paramName+="["+(i - expectedLength+1)+"]";
					}
					else if (!validateParameterCount && (i >=expectedLength)) {
						break;
					}
					e=Function._validateParameter(params[i], expectedParam, paramName);
					if (e) {
						e.popStackFrame();
						return e;
					}
				}
				return null;
			};
		}
		if (!Function._validateParameterCount) {
			Function._validateParameterCount=function (params, expectedParams, validateParameterCount) {
				var i, error, expectedLen=expectedParams.length, actualLen=params.length;
				if (actualLen < expectedLen) {
					var minParams=expectedLen;
					for (i=0; i < expectedLen; i++) {
						var param=expectedParams[i];
						if (param.optional || param.parameterArray) {
							minParams--;
						}
					}
					if (actualLen < minParams) {
						error=true;
					}
				}
				else if (validateParameterCount && (actualLen > expectedLen)) {
					error=true;
					for (i=0; i < expectedLen; i++) {
						if (expectedParams[i].parameterArray) {
							error=false;
							break;
						}
					}
				}
				if (error) {
					var e=MsAjaxError.parameterCount();
					e.popStackFrame();
					return e;
				}
				return null;
			};
		}
		if (!Function._validateParameter) {
			Function._validateParameter=function (param, expectedParam, paramName) {
				var e, expectedType=expectedParam.type, expectedInteger=!!expectedParam.integer, expectedDomElement=!!expectedParam.domElement, mayBeNull=!!expectedParam.mayBeNull;
				e=Function._validateParameterType(param, expectedType, expectedInteger, expectedDomElement, mayBeNull, paramName);
				if (e) {
					e.popStackFrame();
					return e;
				}
				var expectedElementType=expectedParam.elementType, elementMayBeNull=!!expectedParam.elementMayBeNull;
				if (expectedType===Array && typeof (param) !=="undefined" && param !==null &&
					(expectedElementType || !elementMayBeNull)) {
					var expectedElementInteger=!!expectedParam.elementInteger, expectedElementDomElement=!!expectedParam.elementDomElement;
					for (var i=0; i < param.length; i++) {
						var elem=param[i];
						e=Function._validateParameterType(elem, expectedElementType, expectedElementInteger, expectedElementDomElement, elementMayBeNull, paramName+"["+i+"]");
						if (e) {
							e.popStackFrame();
							return e;
						}
					}
				}
				return null;
			};
		}
		if (!Function._validateParameterType) {
			Function._validateParameterType=function (param, expectedType, expectedInteger, expectedDomElement, mayBeNull, paramName) {
				var e, i;
				if (typeof (param)==="undefined") {
					if (mayBeNull) {
						return null;
					}
					else {
						e=OfficeExt.MsAjaxError.argumentUndefined(paramName);
						e.popStackFrame();
						return e;
					}
				}
				if (param===null) {
					if (mayBeNull) {
						return null;
					}
					else {
						e=OfficeExt.MsAjaxError.argumentNull(paramName);
						e.popStackFrame();
						return e;
					}
				}
				if (expectedType && !OfficeExt.MsAjaxTypeHelper.isInstanceOfType(expectedType, param)) {
					e=OfficeExt.MsAjaxError.argumentType(paramName, typeof (param), expectedType);
					e.popStackFrame();
					return e;
				}
				return null;
			};
		}
		if (!window.Type) {
			window.Type=Function;
		}
		if (!Type.registerNamespace) {
			Type.registerNamespace=function (ns) {
				var namespaceParts=ns.split('.');
				var currentNamespace=window;
				for (var i=0; i < namespaceParts.length; i++) {
					currentNamespace[namespaceParts[i]]=currentNamespace[namespaceParts[i]] || {};
					currentNamespace=currentNamespace[namespaceParts[i]];
				}
			};
		}
		if (!Type.prototype.registerClass) {
			Type.prototype.registerClass=function (cls) { cls={}; };
		}
		if (typeof (Sys)==="undefined") {
			Type.registerNamespace('Sys');
		}
		if (!Error.prototype.popStackFrame) {
			Error.prototype.popStackFrame=function () {
				if (arguments.length !==0)
					throw MsAjaxError.parameterCount();
				if (typeof (this.stack)==="undefined" || this.stack===null ||
					typeof (this.fileName)==="undefined" || this.fileName===null ||
					typeof (this.lineNumber)==="undefined" || this.lineNumber===null) {
					return;
				}
				var stackFrames=this.stack.split("\n");
				var currentFrame=stackFrames[0];
				var pattern=this.fileName+":"+this.lineNumber;
				while (typeof (currentFrame) !=="undefined" &&
					currentFrame !==null &&
					currentFrame.indexOf(pattern)===-1) {
					stackFrames.shift();
					currentFrame=stackFrames[0];
				}
				var nextFrame=stackFrames[1];
				if (typeof (nextFrame)==="undefined" || nextFrame===null) {
					return;
				}
				var nextFrameParts=nextFrame.match(/@(.*):(\d+)$/);
				if (typeof (nextFrameParts)==="undefined" || nextFrameParts===null) {
					return;
				}
				this.fileName=nextFrameParts[1];
				this.lineNumber=parseInt(nextFrameParts[2]);
				stackFrames.shift();
				this.stack=stackFrames.join("\n");
			};
		}
		OsfMsAjaxFactory.msAjaxError=MsAjaxError;
		OsfMsAjaxFactory.msAjaxString=MsAjaxString;
		OsfMsAjaxFactory.msAjaxDebug=MsAjaxDebug;
	}
})(OfficeExt || (OfficeExt={}));
OSF.OUtil.setNamespace("SafeArray", OSF.DDA);
OSF.DDA.SafeArray.Response={
	Status: 0,
	Payload: 1
};
OSF.DDA.SafeArray.UniqueArguments={
	Offset: "offset",
	Run: "run",
	BindingSpecificData: "bindingSpecificData",
	MergedCellGuid: "{66e7831f-81b2-42e2-823c-89e872d541b3}"
};
OSF.OUtil.setNamespace("Delegate", OSF.DDA.SafeArray);
OSF.DDA.SafeArray.Delegate._onException=function OSF_DDA_SafeArray_Delegate$OnException(ex, args) {
	var status;
	var statusNumber=ex.number;
	if (statusNumber) {
		switch (statusNumber) {
			case -2146828218:
				status=OSF.DDA.ErrorCodeManager.errorCodes.ooeNoCapability;
				break;
			case -2147467259:
				status=OSF.DDA.ErrorCodeManager.errorCodes.ooeDialogAlreadyOpened;
				break;
			case -2146828283:
				status=OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidParam;
				break;
			case -2147209089:
				status=OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidParam;
				break;
			case -2147208704:
				status=OSF.DDA.ErrorCodeManager.errorCodes.ooeTooManyIncompleteRequests;
				break;
			case -2146827850:
			default:
				status=OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError;
				break;
		}
	}
	if (args.onComplete) {
		args.onComplete(status || OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError);
	}
};
OSF.DDA.SafeArray.Delegate._onExceptionSyncMethod=function OSF_DDA_SafeArray_Delegate$OnExceptionSyncMethod(ex, args) {
	var status;
	var number=ex.number;
	if (number) {
		switch (number) {
			case -2146828218:
				status=OSF.DDA.ErrorCodeManager.errorCodes.ooeNoCapability;
				break;
			case -2146827850:
			default:
				status=OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError;
				break;
		}
	}
	return status || OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError;
};
OSF.DDA.SafeArray.Delegate.SpecialProcessor=function OSF_DDA_SafeArray_Delegate_SpecialProcessor() {
	function _2DVBArrayToJaggedArray(vbArr) {
		var ret;
		try {
			var rows=vbArr.ubound(1);
			var cols=vbArr.ubound(2);
			vbArr=vbArr.toArray();
			if (rows==1 && cols==1) {
				ret=[vbArr];
			}
			else {
				ret=[];
				for (var row=0; row < rows; row++) {
					var rowArr=[];
					for (var col=0; col < cols; col++) {
						var datum=vbArr[row * cols+col];
						if (datum !=OSF.DDA.SafeArray.UniqueArguments.MergedCellGuid) {
							rowArr.push(datum);
						}
					}
					if (rowArr.length > 0) {
						ret.push(rowArr);
					}
				}
			}
		}
		catch (ex) {
		}
		return ret;
	}
	var complexTypes=[];
	var dynamicTypes={};
	dynamicTypes[Microsoft.Office.WebExtension.Parameters.Data]=(function () {
		var tableRows=0;
		var tableHeaders=1;
		return {
			toHost: function OSF_DDA_SafeArray_Delegate_SpecialProcessor_Data$toHost(data) {
				if (OSF.DDA.TableDataProperties && typeof data !="string" && data[OSF.DDA.TableDataProperties.TableRows] !==undefined) {
					var tableData=[];
					tableData[tableRows]=data[OSF.DDA.TableDataProperties.TableRows];
					tableData[tableHeaders]=data[OSF.DDA.TableDataProperties.TableHeaders];
					data=tableData;
				}
				return data;
			},
			fromHost: function OSF_DDA_SafeArray_Delegate_SpecialProcessor_Data$fromHost(hostArgs) {
				var ret;
				if (hostArgs.toArray) {
					var dimensions=hostArgs.dimensions();
					if (dimensions===2) {
						ret=_2DVBArrayToJaggedArray(hostArgs);
					}
					else {
						var array=hostArgs.toArray();
						if (array.length===2 && ((array[0] !=null && array[0].toArray) || (array[1] !=null && array[1].toArray))) {
							ret={};
							ret[OSF.DDA.TableDataProperties.TableRows]=_2DVBArrayToJaggedArray(array[tableRows]);
							ret[OSF.DDA.TableDataProperties.TableHeaders]=_2DVBArrayToJaggedArray(array[tableHeaders]);
						}
						else {
							ret=array;
						}
					}
				}
				else {
					ret=hostArgs;
				}
				return ret;
			}
		};
	})();
	OSF.DDA.SafeArray.Delegate.SpecialProcessor.uber.constructor.call(this, complexTypes, dynamicTypes);
	this.unpack=function OSF_DDA_SafeArray_Delegate_SpecialProcessor$unpack(param, arg) {
		var value;
		if (this.isComplexType(param) || OSF.DDA.ListType.isListType(param)) {
			var toArraySupported=(arg || typeof arg==="unknown") && arg.toArray;
			value=toArraySupported ? arg.toArray() : arg || {};
		}
		else if (this.isDynamicType(param)) {
			value=dynamicTypes[param].fromHost(arg);
		}
		else {
			value=arg;
		}
		return value;
	};
};
OSF.OUtil.extend(OSF.DDA.SafeArray.Delegate.SpecialProcessor, OSF.DDA.SpecialProcessor);
OSF.DDA.SafeArray.Delegate.ParameterMap=OSF.DDA.getDecoratedParameterMap(new OSF.DDA.SafeArray.Delegate.SpecialProcessor(), [
	{
		type: Microsoft.Office.WebExtension.Parameters.ValueFormat,
		toHost: [
			{ name: Microsoft.Office.WebExtension.ValueFormat.Unformatted, value: 0 },
			{ name: Microsoft.Office.WebExtension.ValueFormat.Formatted, value: 1 }
		]
	},
	{
		type: Microsoft.Office.WebExtension.Parameters.FilterType,
		toHost: [
			{ name: Microsoft.Office.WebExtension.FilterType.All, value: 0 }
		]
	}
]);
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.PropertyDescriptors.AsyncResultStatus,
	fromHost: [
		{ name: Microsoft.Office.WebExtension.AsyncResultStatus.Succeeded, value: 0 },
		{ name: Microsoft.Office.WebExtension.AsyncResultStatus.Failed, value: 1 }
	]
});
OSF.DDA.SafeArray.Delegate.executeAsync=function OSF_DDA_SafeArray_Delegate$ExecuteAsync(args) {
	function toArray(args) {
		var arrArgs=args;
		if (OSF.OUtil.isArray(args)) {
			var len=arrArgs.length;
			for (var i=0; i < len; i++) {
				arrArgs[i]=toArray(arrArgs[i]);
			}
		}
		else if (OSF.OUtil.isDate(args)) {
			arrArgs=args.getVarDate();
		}
		else if (typeof args==="object" && !OSF.OUtil.isArray(args)) {
			arrArgs=[];
			for (var index in args) {
				if (!OSF.OUtil.isFunction(args[index])) {
					arrArgs[index]=toArray(args[index]);
				}
			}
		}
		return arrArgs;
	}
	function fromSafeArray(value) {
		var ret=value;
		if (value !=null && value.toArray) {
			var arrayResult=value.toArray();
			ret=new Array(arrayResult.length);
			for (var i=0; i < arrayResult.length; i++) {
				ret[i]=fromSafeArray(arrayResult[i]);
			}
		}
		return ret;
	}
	try {
		if (args.onCalling) {
			args.onCalling();
		}
		OSF.ClientHostController.execute(args.dispId, toArray(args.hostCallArgs), function OSF_DDA_SafeArrayFacade$Execute_OnResponse(hostResponseArgs, resultCode) {
			var result=hostResponseArgs.toArray();
			var status=result[OSF.DDA.SafeArray.Response.Status];
			if (status==OSF.DDA.ErrorCodeManager.errorCodes.ooeChunkResult) {
				var payload=result[OSF.DDA.SafeArray.Response.Payload];
				payload=fromSafeArray(payload);
				if (payload !=null) {
					if (!args._chunkResultData) {
						args._chunkResultData=new Array();
					}
					args._chunkResultData[payload[0]]=payload[1];
				}
				return false;
			}
			if (args.onReceiving) {
				args.onReceiving();
			}
			if (args.onComplete) {
				var payload;
				if (status==OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess) {
					if (result.length > 2) {
						payload=[];
						for (var i=1; i < result.length; i++)
							payload[i - 1]=result[i];
					}
					else {
						payload=result[OSF.DDA.SafeArray.Response.Payload];
					}
					if (args._chunkResultData) {
						payload=fromSafeArray(payload);
						if (payload !=null) {
							var expectedChunkCount=payload[payload.length - 1];
							if (args._chunkResultData.length==expectedChunkCount) {
								payload[payload.length - 1]=args._chunkResultData;
							}
							else {
								status=OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError;
							}
						}
					}
				}
				else {
					payload=result[OSF.DDA.SafeArray.Response.Payload];
				}
				args.onComplete(status, payload);
			}
			return true;
		});
	}
	catch (ex) {
		OSF.DDA.SafeArray.Delegate._onException(ex, args);
	}
};
OSF.DDA.SafeArray.Delegate._getOnAfterRegisterEvent=function OSF_DDA_SafeArrayDelegate$GetOnAfterRegisterEvent(register, args) {
	var startTime=(new Date()).getTime();
	return function OSF_DDA_SafeArrayDelegate$OnAfterRegisterEvent(hostResponseArgs) {
		if (args.onReceiving) {
			args.onReceiving();
		}
		var status=hostResponseArgs.toArray ? hostResponseArgs.toArray()[OSF.DDA.SafeArray.Response.Status] : hostResponseArgs;
		if (args.onComplete) {
			args.onComplete(status);
		}
		if (OSF.AppTelemetry) {
			OSF.AppTelemetry.onRegisterDone(register, args.dispId, Math.abs((new Date()).getTime() - startTime), status);
		}
	};
};
OSF.DDA.SafeArray.Delegate.registerEventAsync=function OSF_DDA_SafeArray_Delegate$RegisterEventAsync(args) {
	if (args.onCalling) {
		args.onCalling();
	}
	var callback=OSF.DDA.SafeArray.Delegate._getOnAfterRegisterEvent(true, args);
	try {
		OSF.ClientHostController.registerEvent(args.dispId, args.targetId, function OSF_DDA_SafeArrayDelegate$RegisterEventAsync_OnEvent(eventDispId, payload) {
			if (args.onEvent) {
				args.onEvent(payload);
			}
			if (OSF.AppTelemetry) {
				OSF.AppTelemetry.onEventDone(args.dispId);
			}
		}, callback);
	}
	catch (ex) {
		OSF.DDA.SafeArray.Delegate._onException(ex, args);
	}
};
OSF.DDA.SafeArray.Delegate.unregisterEventAsync=function OSF_DDA_SafeArray_Delegate$UnregisterEventAsync(args) {
	if (args.onCalling) {
		args.onCalling();
	}
	var callback=OSF.DDA.SafeArray.Delegate._getOnAfterRegisterEvent(false, args);
	try {
		OSF.ClientHostController.unregisterEvent(args.dispId, args.targetId, callback);
	}
	catch (ex) {
		OSF.DDA.SafeArray.Delegate._onException(ex, args);
	}
};
OSF.ClientMode={
	ReadWrite: 0,
	ReadOnly: 1
};
OSF.DDA.RichInitializationReason={
	1: Microsoft.Office.WebExtension.InitializationReason.Inserted,
	2: Microsoft.Office.WebExtension.InitializationReason.DocumentOpened
};
OSF.InitializationHelper=function OSF_InitializationHelper(hostInfo, webAppState, context, settings, hostFacade) {
	this._hostInfo=hostInfo;
	this._webAppState=webAppState;
	this._context=context;
	this._settings=settings;
	this._hostFacade=hostFacade;
	this._initializeSettings=this.initializeSettings;
};
OSF.InitializationHelper.prototype.deserializeSettings=function OSF_InitializationHelper$deserializeSettings(serializedSettings, refreshSupported) {
	var settings;
	var osfSessionStorage=OSF.OUtil.getSessionStorage();
	if (osfSessionStorage) {
		var storageSettings=osfSessionStorage.getItem(OSF._OfficeAppFactory.getCachedSessionSettingsKey());
		if (storageSettings) {
			serializedSettings=JSON.parse(storageSettings);
		}
		else {
			storageSettings=JSON.stringify(serializedSettings);
			osfSessionStorage.setItem(OSF._OfficeAppFactory.getCachedSessionSettingsKey(), storageSettings);
		}
	}
	var deserializedSettings=OSF.DDA.SettingsManager.deserializeSettings(serializedSettings);
	if (refreshSupported) {
		settings=new OSF.DDA.RefreshableSettings(deserializedSettings);
	}
	else {
		settings=new OSF.DDA.Settings(deserializedSettings);
	}
	return settings;
};
OSF.InitializationHelper.prototype.saveAndSetDialogInfo=function OSF_InitializationHelper$saveAndSetDialogInfo(hostInfoValue) {
};
OSF.InitializationHelper.prototype.setAgaveHostCommunication=function OSF_InitializationHelper$setAgaveHostCommunication() {
};
OSF.InitializationHelper.prototype.prepareRightBeforeWebExtensionInitialize=function OSF_InitializationHelper$prepareRightBeforeWebExtensionInitialize(appContext) {
	this.prepareApiSurface(appContext);
	Microsoft.Office.WebExtension.initialize(this.getInitializationReason(appContext));
};
OSF.InitializationHelper.prototype.prepareApiSurface=function OSF_InitializationHelper$prepareApiSurfaceAndInitialize(appContext) {
	var license=new OSF.DDA.License(appContext.get_eToken());
	var getOfficeThemeHandler=(OSF.DDA.OfficeTheme && OSF.DDA.OfficeTheme.getOfficeTheme) ? OSF.DDA.OfficeTheme.getOfficeTheme : null;
	if (appContext.get_isDialog()) {
		if (OSF.DDA.UI.ChildUI) {
			appContext.ui=new OSF.DDA.UI.ChildUI();
		}
	}
	else {
		if (OSF.DDA.UI.ParentUI) {
			appContext.ui=new OSF.DDA.UI.ParentUI();
			if (OfficeExt.Container) {
				OSF.DDA.DispIdHost.addAsyncMethods(appContext.ui, [OSF.DDA.AsyncMethodNames.CloseContainerAsync]);
			}
		}
	}
	if (OSF.DDA.OpenBrowser) {
		OSF.DDA.DispIdHost.addAsyncMethods(appContext.ui, [OSF.DDA.AsyncMethodNames.OpenBrowserWindow]);
	}
	if (OSF.DDA.ExecuteFeature) {
		OSF.DDA.DispIdHost.addAsyncMethods(appContext.ui, [OSF.DDA.AsyncMethodNames.ExecuteFeature]);
	}
	if (OSF.DDA.QueryFeature) {
		OSF.DDA.DispIdHost.addAsyncMethods(appContext.ui, [OSF.DDA.AsyncMethodNames.QueryFeature]);
	}
	if (OSF.DDA.Auth) {
		appContext.auth=new OSF.DDA.Auth();
		OSF.DDA.DispIdHost.addAsyncMethods(appContext.auth, [OSF.DDA.AsyncMethodNames.GetAccessTokenAsync]);
	}
	OSF._OfficeAppFactory.setContext(new OSF.DDA.Context(appContext, appContext.doc, license, null, getOfficeThemeHandler));
	var getDelegateMethods, parameterMap;
	getDelegateMethods=OSF.DDA.DispIdHost.getClientDelegateMethods;
	parameterMap=OSF.DDA.SafeArray.Delegate.ParameterMap;
	OSF._OfficeAppFactory.setHostFacade(new OSF.DDA.DispIdHost.Facade(getDelegateMethods, parameterMap));
};
OSF.InitializationHelper.prototype.getInitializationReason=function (appContext) { return OSF.DDA.RichInitializationReason[appContext.get_reason()]; };
OSF.DDA.DispIdHost.getClientDelegateMethods=function (actionId) {
	var delegateMethods={};
	delegateMethods[OSF.DDA.DispIdHost.Delegates.ExecuteAsync]=OSF.DDA.SafeArray.Delegate.executeAsync;
	delegateMethods[OSF.DDA.DispIdHost.Delegates.RegisterEventAsync]=OSF.DDA.SafeArray.Delegate.registerEventAsync;
	delegateMethods[OSF.DDA.DispIdHost.Delegates.UnregisterEventAsync]=OSF.DDA.SafeArray.Delegate.unregisterEventAsync;
	delegateMethods[OSF.DDA.DispIdHost.Delegates.OpenDialog]=OSF.DDA.SafeArray.Delegate.openDialog;
	delegateMethods[OSF.DDA.DispIdHost.Delegates.CloseDialog]=OSF.DDA.SafeArray.Delegate.closeDialog;
	delegateMethods[OSF.DDA.DispIdHost.Delegates.MessageParent]=OSF.DDA.SafeArray.Delegate.messageParent;
	delegateMethods[OSF.DDA.DispIdHost.Delegates.SendMessage]=OSF.DDA.SafeArray.Delegate.sendMessage;
	if (OSF.DDA.AsyncMethodNames.RefreshAsync && actionId==OSF.DDA.AsyncMethodNames.RefreshAsync.id) {
		var readSerializedSettings=function (hostCallArgs, onCalling, onReceiving) {
			return OSF.DDA.ClientSettingsManager.read(onCalling, onReceiving);
		};
		delegateMethods[OSF.DDA.DispIdHost.Delegates.ExecuteAsync]=OSF.DDA.ClientSettingsManager.getSettingsExecuteMethod(readSerializedSettings);
	}
	if (OSF.DDA.AsyncMethodNames.SaveAsync && actionId==OSF.DDA.AsyncMethodNames.SaveAsync.id) {
		var writeSerializedSettings=function (hostCallArgs, onCalling, onReceiving) {
			return OSF.DDA.ClientSettingsManager.write(hostCallArgs[OSF.DDA.SettingsManager.SerializedSettings], hostCallArgs[Microsoft.Office.WebExtension.Parameters.OverwriteIfStale], onCalling, onReceiving);
		};
		delegateMethods[OSF.DDA.DispIdHost.Delegates.ExecuteAsync]=OSF.DDA.ClientSettingsManager.getSettingsExecuteMethod(writeSerializedSettings);
	}
	return delegateMethods;
};
var OfficeExt;
(function (OfficeExt) {
	var MacRichClientHostController=(function () {
		function MacRichClientHostController() {
		}
		MacRichClientHostController.prototype.execute=function (id, params, callback) {
			setTimeout(function () {
				window.external.Execute(id, params, callback);
			}, 0);
		};
		MacRichClientHostController.prototype.registerEvent=function (id, targetId, handler, callback) {
			setTimeout(function () {
				window.external.RegisterEvent(id, targetId, handler, callback);
			}, 0);
		};
		MacRichClientHostController.prototype.unregisterEvent=function (id, targetId, callback) {
			setTimeout(function () {
				window.external.UnregisterEvent(id, targetId, callback);
			}, 0);
		};
		MacRichClientHostController.prototype.openDialog=function (id, targetId, handler, callback) {
			if (MacRichClientHostController.popup && !MacRichClientHostController.popup.closed) {
				callback(OSF.DDA.ErrorCodeManager.errorCodes.ooeDialogAlreadyOpened);
				return;
			}
			var magicWord="action=displayDialog";
			window.dialogAPIErrorCode=undefined;
			var fragmentSeparator='#';
			var callArgs=JSON.parse(targetId);
			var callUrl=callArgs.url;
			if (!callUrl) {
				return;
			}
			var urlParts=callUrl.split(fragmentSeparator);
			var seperator="?";
			if (urlParts[0].indexOf("?") > -1) {
				seperator="&";
			}
			var width=screen.width * callArgs.width / 100;
			var height=screen.height * callArgs.height / 100;
			var params="width="+width+", height="+height;
			urlParts[0]=urlParts[0].concat(seperator).concat(magicWord);
			var openUrl=urlParts.join(fragmentSeparator);
			MacRichClientHostController.popup=window.open(openUrl, "", params);
			function receiveMessage(event) {
				if (event.source==MacRichClientHostController.popup) {
					try {
						var messageObj=JSON.parse(event.data);
						if (messageObj.dialogMessage) {
							handler(id, [OSF.DialogMessageType.DialogMessageReceived, messageObj.dialogMessage.messageContent]);
						}
					}
					catch (e) {
						OsfMsAjaxFactory.msAjaxDebug.trace("messages received cannot be handlered. Message:"+event.data);
					}
				}
			}
			function checkWindowClose() {
				try {
					if (MacRichClientHostController.popup==null || MacRichClientHostController.popup.closed) {
						window.clearInterval(MacRichClientHostController.interval);
						window.removeEventListener("message", receiveMessage);
						MacRichClientHostController.NotifyError=null;
						handler(id, [OSF.DialogMessageType.DialogClosed]);
					}
				}
				catch (e) {
					OsfMsAjaxFactory.msAjaxDebug.trace("Error happened when popup window closed.");
				}
			}
			if (MacRichClientHostController.popup !=undefined && window.dialogAPIErrorCode==undefined) {
				window.addEventListener("message", receiveMessage);
				MacRichClientHostController.interval=window.setInterval(checkWindowClose, 500);
				function notifyError(errorCode) {
					handler(id, [errorCode]);
				}
				MacRichClientHostController.NotifyError=notifyError;
				callback(OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess);
			}
			else {
				var error=OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError;
				if (window.dialogAPIErrorCode) {
					error=window.dialogAPIErrorCode;
				}
				callback(error);
			}
		};
		MacRichClientHostController.prototype.messageParent=function (params) {
			var message=params[Microsoft.Office.WebExtension.Parameters.MessageToParent];
			var messageObj={ dialogMessage: { messageType: OSF.DialogMessageType.DialogMessageReceived, messageContent: message } };
			window.opener.postMessage(JSON.stringify(messageObj), window.location.origin);
		};
		MacRichClientHostController.prototype.closeDialog=function (id, targetId, callback) {
			if (MacRichClientHostController.popup) {
				if (MacRichClientHostController.interval) {
					window.clearInterval(MacRichClientHostController.interval);
				}
				MacRichClientHostController.popup.close();
				MacRichClientHostController.popup=null;
				MacRichClientHostController.NotifyError=null;
				callback(OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess);
			}
			else {
				callback(OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError);
			}
		};
		MacRichClientHostController.prototype.sendMessage=function (params) {
		};
		return MacRichClientHostController;
	})();
	OfficeExt.MacRichClientHostController=MacRichClientHostController;
})(OfficeExt || (OfficeExt={}));
OSF.ClientHostController=new OfficeExt.MacRichClientHostController();
var OfficeExt;
(function (OfficeExt) {
	var OfficeTheme;
	(function (OfficeTheme) {
		var OfficeThemeManager=(function () {
			function OfficeThemeManager() {
				this._osfOfficeTheme=null;
				this._osfOfficeThemeTimeStamp=null;
			}
			OfficeThemeManager.prototype.getOfficeTheme=function () {
				if (OSF.DDA._OsfControlContext) {
					if (this._osfOfficeTheme && this._osfOfficeThemeTimeStamp && ((new Date()).getTime() - this._osfOfficeThemeTimeStamp < OfficeThemeManager._osfOfficeThemeCacheValidPeriod)) {
						if (OSF.AppTelemetry) {
							OSF.AppTelemetry.onPropertyDone("GetOfficeThemeInfo", 0);
						}
					}
					else {
						var startTime=(new Date()).getTime();
						var osfOfficeTheme=OSF.DDA._OsfControlContext.GetOfficeThemeInfo();
						var endTime=(new Date()).getTime();
						if (OSF.AppTelemetry) {
							OSF.AppTelemetry.onPropertyDone("GetOfficeThemeInfo", Math.abs(endTime - startTime));
						}
						this._osfOfficeTheme=JSON.parse(osfOfficeTheme);
						for (var color in this._osfOfficeTheme) {
							this._osfOfficeTheme[color]=OSF.OUtil.convertIntToCssHexColor(this._osfOfficeTheme[color]);
						}
						this._osfOfficeThemeTimeStamp=endTime;
					}
					return this._osfOfficeTheme;
				}
			};
			OfficeThemeManager.instance=function () {
				if (OfficeThemeManager._instance==null) {
					OfficeThemeManager._instance=new OfficeThemeManager();
				}
				return OfficeThemeManager._instance;
			};
			OfficeThemeManager._osfOfficeThemeCacheValidPeriod=5000;
			OfficeThemeManager._instance=null;
			return OfficeThemeManager;
		})();
		OfficeTheme.OfficeThemeManager=OfficeThemeManager;
		OSF.OUtil.setNamespace("OfficeTheme", OSF.DDA);
		OSF.DDA.OfficeTheme.getOfficeTheme=OfficeExt.OfficeTheme.OfficeThemeManager.instance().getOfficeTheme;
	})(OfficeTheme=OfficeExt.OfficeTheme || (OfficeExt.OfficeTheme={}));
})(OfficeExt || (OfficeExt={}));
OSF.DDA.ClientSettingsManager={
	getSettingsExecuteMethod: function OSF_DDA_ClientSettingsManager$getSettingsExecuteMethod(hostDelegateMethod) {
		return function (args) {
			var status, response;
			try {
				response=hostDelegateMethod(args.hostCallArgs, args.onCalling, args.onReceiving);
				status=OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess;
			}
			catch (ex) {
				status=OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError;
				response={ name: Strings.OfficeOM.L_InternalError, message: ex };
			}
			if (args.onComplete) {
				args.onComplete(status, response);
			}
		};
	},
	read: function OSF_DDA_ClientSettingsManager$read(onCalling, onReceiving) {
		var keys=[];
		var values=[];
		if (onCalling) {
			onCalling();
		}
		OSF.DDA._OsfControlContext.GetSettings().Read(keys, values);
		if (onReceiving) {
			onReceiving();
		}
		var serializedSettings={};
		for (var index=0; index < keys.length; index++) {
			serializedSettings[keys[index]]=values[index];
		}
		return serializedSettings;
	},
	write: function OSF_DDA_ClientSettingsManager$write(serializedSettings, overwriteIfStale, onCalling, onReceiving) {
		var keys=[];
		var values=[];
		for (var key in serializedSettings) {
			keys.push(key);
			values.push(serializedSettings[key]);
		}
		if (onCalling) {
			onCalling();
		}
		OSF.DDA._OsfControlContext.GetSettings().Write(keys, values);
		if (onReceiving) {
			onReceiving();
		}
	}
};
OSF.InitializationHelper.prototype.initializeSettings=function OSF_InitializationHelper$initializeSettings(refreshSupported) {
	var serializedSettings=OSF.DDA.ClientSettingsManager.read();
	var settings=this.deserializeSettings(serializedSettings, refreshSupported);
	return settings;
};
OSF.InitializationHelper.prototype.getAppContext=function OSF_InitializationHelper$getAppContext(wnd, gotAppContext) {
	var returnedContext;
	var context;
	var warningText="Warning: Office.js is loaded outside of Office client";
	try {
		if (window.external && typeof window.external.GetContext !=='undefined') {
			context=OSF.DDA._OsfControlContext=window.external.GetContext();
		}
		else {
			OsfMsAjaxFactory.msAjaxDebug.trace(warningText);
			return;
		}
	}
	catch (e) {
		OsfMsAjaxFactory.msAjaxDebug.trace(warningText);
		return;
	}
	var appType=context.GetAppType();
	var id=context.GetSolutionRef();
	var version=context.GetAppVersionMajor();
	var minorVersion=context.GetAppVersionMinor();
	var UILocale=context.GetAppUILocale();
	var dataLocale=context.GetAppDataLocale();
	var docUrl=context.GetDocUrl();
	var clientMode=context.GetAppCapabilities();
	var reason=context.GetActivationMode();
	var osfControlType=context.GetControlIntegrationLevel();
	var settings=[];
	var eToken;
	try {
		eToken=context.GetSolutionToken();
	}
	catch (ex) {
	}
	var correlationId;
	if (typeof context.GetCorrelationId !=="undefined") {
		correlationId=context.GetCorrelationId();
	}
	var appInstanceId;
	if (typeof context.GetInstanceId !=="undefined") {
		appInstanceId=context.GetInstanceId();
	}
	var touchEnabled;
	if (typeof context.GetTouchEnabled !=="undefined") {
		touchEnabled=context.GetTouchEnabled();
	}
	var commerceAllowed;
	if (typeof context.GetCommerceAllowed !=="undefined") {
		commerceAllowed=context.GetCommerceAllowed();
	}
	var requirementMatrix;
	if (typeof context.GetSupportedMatrix !=="undefined") {
		requirementMatrix=context.GetSupportedMatrix();
	}
	var hostCustomMessage;
	if (typeof context.GetHostCustomMessage !=="undefined") {
		hostCustomMessage=context.GetHostCustomMessage();
	}
	var hostFullVersion;
	if (typeof context.GetHostFullVersion !=="undefined") {
		hostFullVersion=context.GetHostFullVersion();
	}
	var dialogRequirementMatrix;
	if (typeof context.GetDialogRequirementMatrix !="undefined") {
		dialogRequirementMatrix=context.GetDialogRequirementMatrix();
	}
	eToken=eToken ? eToken.toString() : "";
	returnedContext=new OSF.OfficeAppContext(id, appType, version, UILocale, dataLocale, docUrl, clientMode, settings, reason, osfControlType, eToken, correlationId, appInstanceId, touchEnabled, commerceAllowed, minorVersion, requirementMatrix, hostCustomMessage, hostFullVersion, undefined, undefined, undefined, dialogRequirementMatrix);
	if (OSF.AppTelemetry) {
		OSF.AppTelemetry.initialize(returnedContext);
	}
	gotAppContext(returnedContext);
};
var OSFLog;
(function (OSFLog) {
	var BaseUsageData=(function () {
		function BaseUsageData(table) {
			this._table=table;
			this._fields={};
		}
		Object.defineProperty(BaseUsageData.prototype, "Fields", {
			get: function () {
				return this._fields;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(BaseUsageData.prototype, "Table", {
			get: function () {
				return this._table;
			},
			enumerable: true,
			configurable: true
		});
		BaseUsageData.prototype.SerializeFields=function () {
		};
		BaseUsageData.prototype.SetSerializedField=function (key, value) {
			if (typeof (value) !=="undefined" && value !==null) {
				this._serializedFields[key]=value.toString();
			}
		};
		BaseUsageData.prototype.SerializeRow=function () {
			this._serializedFields={};
			this.SetSerializedField("Table", this._table);
			this.SerializeFields();
			return JSON.stringify(this._serializedFields);
		};
		return BaseUsageData;
	})();
	OSFLog.BaseUsageData=BaseUsageData;
	var AppActivatedUsageData=(function (_super) {
		__extends(AppActivatedUsageData, _super);
		function AppActivatedUsageData() {
			_super.call(this, "AppActivated");
		}
		Object.defineProperty(AppActivatedUsageData.prototype, "CorrelationId", {
			get: function () { return this.Fields["CorrelationId"]; },
			set: function (value) { this.Fields["CorrelationId"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(AppActivatedUsageData.prototype, "SessionId", {
			get: function () { return this.Fields["SessionId"]; },
			set: function (value) { this.Fields["SessionId"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(AppActivatedUsageData.prototype, "AppId", {
			get: function () { return this.Fields["AppId"]; },
			set: function (value) { this.Fields["AppId"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(AppActivatedUsageData.prototype, "AppInstanceId", {
			get: function () { return this.Fields["AppInstanceId"]; },
			set: function (value) { this.Fields["AppInstanceId"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(AppActivatedUsageData.prototype, "AppURL", {
			get: function () { return this.Fields["AppURL"]; },
			set: function (value) { this.Fields["AppURL"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(AppActivatedUsageData.prototype, "AssetId", {
			get: function () { return this.Fields["AssetId"]; },
			set: function (value) { this.Fields["AssetId"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(AppActivatedUsageData.prototype, "Browser", {
			get: function () { return this.Fields["Browser"]; },
			set: function (value) { this.Fields["Browser"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(AppActivatedUsageData.prototype, "UserId", {
			get: function () { return this.Fields["UserId"]; },
			set: function (value) { this.Fields["UserId"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(AppActivatedUsageData.prototype, "Host", {
			get: function () { return this.Fields["Host"]; },
			set: function (value) { this.Fields["Host"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(AppActivatedUsageData.prototype, "HostVersion", {
			get: function () { return this.Fields["HostVersion"]; },
			set: function (value) { this.Fields["HostVersion"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(AppActivatedUsageData.prototype, "ClientId", {
			get: function () { return this.Fields["ClientId"]; },
			set: function (value) { this.Fields["ClientId"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(AppActivatedUsageData.prototype, "AppSizeWidth", {
			get: function () { return this.Fields["AppSizeWidth"]; },
			set: function (value) { this.Fields["AppSizeWidth"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(AppActivatedUsageData.prototype, "AppSizeHeight", {
			get: function () { return this.Fields["AppSizeHeight"]; },
			set: function (value) { this.Fields["AppSizeHeight"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(AppActivatedUsageData.prototype, "Message", {
			get: function () { return this.Fields["Message"]; },
			set: function (value) { this.Fields["Message"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(AppActivatedUsageData.prototype, "DocUrl", {
			get: function () { return this.Fields["DocUrl"]; },
			set: function (value) { this.Fields["DocUrl"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(AppActivatedUsageData.prototype, "OfficeJSVersion", {
			get: function () { return this.Fields["OfficeJSVersion"]; },
			set: function (value) { this.Fields["OfficeJSVersion"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(AppActivatedUsageData.prototype, "HostJSVersion", {
			get: function () { return this.Fields["HostJSVersion"]; },
			set: function (value) { this.Fields["HostJSVersion"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(AppActivatedUsageData.prototype, "WacHostEnvironment", {
			get: function () { return this.Fields["WacHostEnvironment"]; },
			set: function (value) { this.Fields["WacHostEnvironment"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(AppActivatedUsageData.prototype, "IsFromWacAutomation", {
			get: function () { return this.Fields["IsFromWacAutomation"]; },
			set: function (value) { this.Fields["IsFromWacAutomation"]=value; },
			enumerable: true,
			configurable: true
		});
		AppActivatedUsageData.prototype.SerializeFields=function () {
			this.SetSerializedField("CorrelationId", this.CorrelationId);
			this.SetSerializedField("SessionId", this.SessionId);
			this.SetSerializedField("AppId", this.AppId);
			this.SetSerializedField("AppInstanceId", this.AppInstanceId);
			this.SetSerializedField("AppURL", this.AppURL);
			this.SetSerializedField("AssetId", this.AssetId);
			this.SetSerializedField("Browser", this.Browser);
			this.SetSerializedField("UserId", this.UserId);
			this.SetSerializedField("Host", this.Host);
			this.SetSerializedField("HostVersion", this.HostVersion);
			this.SetSerializedField("ClientId", this.ClientId);
			this.SetSerializedField("AppSizeWidth", this.AppSizeWidth);
			this.SetSerializedField("AppSizeHeight", this.AppSizeHeight);
			this.SetSerializedField("Message", this.Message);
			this.SetSerializedField("DocUrl", this.DocUrl);
			this.SetSerializedField("OfficeJSVersion", this.OfficeJSVersion);
			this.SetSerializedField("HostJSVersion", this.HostJSVersion);
			this.SetSerializedField("WacHostEnvironment", this.WacHostEnvironment);
			this.SetSerializedField("IsFromWacAutomation", this.IsFromWacAutomation);
		};
		return AppActivatedUsageData;
	})(BaseUsageData);
	OSFLog.AppActivatedUsageData=AppActivatedUsageData;
	var ScriptLoadUsageData=(function (_super) {
		__extends(ScriptLoadUsageData, _super);
		function ScriptLoadUsageData() {
			_super.call(this, "ScriptLoad");
		}
		Object.defineProperty(ScriptLoadUsageData.prototype, "CorrelationId", {
			get: function () { return this.Fields["CorrelationId"]; },
			set: function (value) { this.Fields["CorrelationId"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ScriptLoadUsageData.prototype, "SessionId", {
			get: function () { return this.Fields["SessionId"]; },
			set: function (value) { this.Fields["SessionId"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ScriptLoadUsageData.prototype, "ScriptId", {
			get: function () { return this.Fields["ScriptId"]; },
			set: function (value) { this.Fields["ScriptId"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ScriptLoadUsageData.prototype, "StartTime", {
			get: function () { return this.Fields["StartTime"]; },
			set: function (value) { this.Fields["StartTime"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ScriptLoadUsageData.prototype, "ResponseTime", {
			get: function () { return this.Fields["ResponseTime"]; },
			set: function (value) { this.Fields["ResponseTime"]=value; },
			enumerable: true,
			configurable: true
		});
		ScriptLoadUsageData.prototype.SerializeFields=function () {
			this.SetSerializedField("CorrelationId", this.CorrelationId);
			this.SetSerializedField("SessionId", this.SessionId);
			this.SetSerializedField("ScriptId", this.ScriptId);
			this.SetSerializedField("StartTime", this.StartTime);
			this.SetSerializedField("ResponseTime", this.ResponseTime);
		};
		return ScriptLoadUsageData;
	})(BaseUsageData);
	OSFLog.ScriptLoadUsageData=ScriptLoadUsageData;
	var AppClosedUsageData=(function (_super) {
		__extends(AppClosedUsageData, _super);
		function AppClosedUsageData() {
			_super.call(this, "AppClosed");
		}
		Object.defineProperty(AppClosedUsageData.prototype, "CorrelationId", {
			get: function () { return this.Fields["CorrelationId"]; },
			set: function (value) { this.Fields["CorrelationId"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(AppClosedUsageData.prototype, "SessionId", {
			get: function () { return this.Fields["SessionId"]; },
			set: function (value) { this.Fields["SessionId"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(AppClosedUsageData.prototype, "FocusTime", {
			get: function () { return this.Fields["FocusTime"]; },
			set: function (value) { this.Fields["FocusTime"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(AppClosedUsageData.prototype, "AppSizeFinalWidth", {
			get: function () { return this.Fields["AppSizeFinalWidth"]; },
			set: function (value) { this.Fields["AppSizeFinalWidth"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(AppClosedUsageData.prototype, "AppSizeFinalHeight", {
			get: function () { return this.Fields["AppSizeFinalHeight"]; },
			set: function (value) { this.Fields["AppSizeFinalHeight"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(AppClosedUsageData.prototype, "OpenTime", {
			get: function () { return this.Fields["OpenTime"]; },
			set: function (value) { this.Fields["OpenTime"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(AppClosedUsageData.prototype, "CloseMethod", {
			get: function () { return this.Fields["CloseMethod"]; },
			set: function (value) { this.Fields["CloseMethod"]=value; },
			enumerable: true,
			configurable: true
		});
		AppClosedUsageData.prototype.SerializeFields=function () {
			this.SetSerializedField("CorrelationId", this.CorrelationId);
			this.SetSerializedField("SessionId", this.SessionId);
			this.SetSerializedField("FocusTime", this.FocusTime);
			this.SetSerializedField("AppSizeFinalWidth", this.AppSizeFinalWidth);
			this.SetSerializedField("AppSizeFinalHeight", this.AppSizeFinalHeight);
			this.SetSerializedField("OpenTime", this.OpenTime);
			this.SetSerializedField("CloseMethod", this.CloseMethod);
		};
		return AppClosedUsageData;
	})(BaseUsageData);
	OSFLog.AppClosedUsageData=AppClosedUsageData;
	var APIUsageUsageData=(function (_super) {
		__extends(APIUsageUsageData, _super);
		function APIUsageUsageData() {
			_super.call(this, "APIUsage");
		}
		Object.defineProperty(APIUsageUsageData.prototype, "CorrelationId", {
			get: function () { return this.Fields["CorrelationId"]; },
			set: function (value) { this.Fields["CorrelationId"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(APIUsageUsageData.prototype, "SessionId", {
			get: function () { return this.Fields["SessionId"]; },
			set: function (value) { this.Fields["SessionId"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(APIUsageUsageData.prototype, "APIType", {
			get: function () { return this.Fields["APIType"]; },
			set: function (value) { this.Fields["APIType"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(APIUsageUsageData.prototype, "APIID", {
			get: function () { return this.Fields["APIID"]; },
			set: function (value) { this.Fields["APIID"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(APIUsageUsageData.prototype, "Parameters", {
			get: function () { return this.Fields["Parameters"]; },
			set: function (value) { this.Fields["Parameters"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(APIUsageUsageData.prototype, "ResponseTime", {
			get: function () { return this.Fields["ResponseTime"]; },
			set: function (value) { this.Fields["ResponseTime"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(APIUsageUsageData.prototype, "ErrorType", {
			get: function () { return this.Fields["ErrorType"]; },
			set: function (value) { this.Fields["ErrorType"]=value; },
			enumerable: true,
			configurable: true
		});
		APIUsageUsageData.prototype.SerializeFields=function () {
			this.SetSerializedField("CorrelationId", this.CorrelationId);
			this.SetSerializedField("SessionId", this.SessionId);
			this.SetSerializedField("APIType", this.APIType);
			this.SetSerializedField("APIID", this.APIID);
			this.SetSerializedField("Parameters", this.Parameters);
			this.SetSerializedField("ResponseTime", this.ResponseTime);
			this.SetSerializedField("ErrorType", this.ErrorType);
		};
		return APIUsageUsageData;
	})(BaseUsageData);
	OSFLog.APIUsageUsageData=APIUsageUsageData;
	var AppInitializationUsageData=(function (_super) {
		__extends(AppInitializationUsageData, _super);
		function AppInitializationUsageData() {
			_super.call(this, "AppInitialization");
		}
		Object.defineProperty(AppInitializationUsageData.prototype, "CorrelationId", {
			get: function () { return this.Fields["CorrelationId"]; },
			set: function (value) { this.Fields["CorrelationId"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(AppInitializationUsageData.prototype, "SessionId", {
			get: function () { return this.Fields["SessionId"]; },
			set: function (value) { this.Fields["SessionId"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(AppInitializationUsageData.prototype, "SuccessCode", {
			get: function () { return this.Fields["SuccessCode"]; },
			set: function (value) { this.Fields["SuccessCode"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(AppInitializationUsageData.prototype, "Message", {
			get: function () { return this.Fields["Message"]; },
			set: function (value) { this.Fields["Message"]=value; },
			enumerable: true,
			configurable: true
		});
		AppInitializationUsageData.prototype.SerializeFields=function () {
			this.SetSerializedField("CorrelationId", this.CorrelationId);
			this.SetSerializedField("SessionId", this.SessionId);
			this.SetSerializedField("SuccessCode", this.SuccessCode);
			this.SetSerializedField("Message", this.Message);
		};
		return AppInitializationUsageData;
	})(BaseUsageData);
	OSFLog.AppInitializationUsageData=AppInitializationUsageData;
})(OSFLog || (OSFLog={}));
var Logger;
(function (Logger) {
	"use strict";
	(function (TraceLevel) {
		TraceLevel[TraceLevel["info"]=0]="info";
		TraceLevel[TraceLevel["warning"]=1]="warning";
		TraceLevel[TraceLevel["error"]=2]="error";
	})(Logger.TraceLevel || (Logger.TraceLevel={}));
	var TraceLevel=Logger.TraceLevel;
	(function (SendFlag) {
		SendFlag[SendFlag["none"]=0]="none";
		SendFlag[SendFlag["flush"]=1]="flush";
	})(Logger.SendFlag || (Logger.SendFlag={}));
	var SendFlag=Logger.SendFlag;
	function allowUploadingData() {
	}
	Logger.allowUploadingData=allowUploadingData;
	function sendLog(traceLevel, message, flag) {
	}
	Logger.sendLog=sendLog;
	function creatULSEndpoint() {
		try {
			return new ULSEndpointProxy();
		}
		catch (e) {
			return null;
		}
	}
	var ULSEndpointProxy=(function () {
		function ULSEndpointProxy() {
		}
		ULSEndpointProxy.prototype.writeLog=function (log) {
		};
		ULSEndpointProxy.prototype.loadProxyFrame=function () {
		};
		return ULSEndpointProxy;
	})();
	if (!OSF.Logger) {
		OSF.Logger=Logger;
	}
	Logger.ulsEndpoint=creatULSEndpoint();
})(Logger || (Logger={}));
var OSFAriaLogger;
(function (OSFAriaLogger) {
	var TelemetryEventAppActivated={ name: "AppActivated", enabled: true, basic: true, critical: true, points: [
			{ name: "Browser", type: "string" },
			{ name: "Message", type: "string" },
			{ name: "AppId", type: "string" },
			{ name: "AppURL", type: "string" },
			{ name: "UserId", type: "string" },
			{ name: "Host", type: "string" },
			{ name: "HostVersion", type: "string" },
			{ name: "CorrelationId", type: "string", rename: "HostSessionId" },
			{ name: "AppSizeWidth", type: "int64" },
			{ name: "AppSizeHeight", type: "int64" },
			{ name: "AppInstanceId", type: "string" },
			{ name: "OfficeJSVersion", type: "string" },
			{ name: "HostJSVersion", type: "string" },
			{ name: "IsFromWacAutomation", type: "string" },
		] };
	var TelemetryEventScriptLoad={ name: "ScriptLoad", enabled: true, basic: false, critical: false, points: [
			{ name: "ScriptId", type: "string" },
			{ name: "StartTime", type: "double" },
			{ name: "ResponseTime", type: "double" },
		] };
	var TelemetryEventApiUsage={ name: "APIUsage", enabled: false, basic: false, critical: false, points: [
			{ name: "APIType", type: "string" },
			{ name: "APIID", type: "int64" },
			{ name: "Parameters", type: "string" },
			{ name: "ResponseTime", type: "int64" },
			{ name: "ErrorType", type: "int64" },
		] };
	var TelemetryEventAppInitialization={ name: "AppInitialization", enabled: true, basic: false, critical: false, points: [
			{ name: "SuccessCode", type: "int64" },
			{ name: "Message", type: "string" },
		] };
	var TelemetryEventAppClosed={ name: "AppClosed", enabled: true, basic: false, critical: false, points: [
			{ name: "FocusTime", type: "int64" },
			{ name: "AppSizeFinalWidth", type: "int64" },
			{ name: "AppSizeFinalHeight", type: "int64" },
			{ name: "OpenTime", type: "int64" },
		] };
	var TelemetryEvents=[
		TelemetryEventAppActivated,
		TelemetryEventScriptLoad,
		TelemetryEventApiUsage,
		TelemetryEventAppInitialization,
		TelemetryEventAppClosed,
	];
	function createDataField(value, point) {
		var key=point.rename===undefined ? point.name : point.rename;
		var type=point.type;
		var field=undefined;
		switch (type) {
			case "string":
				field=oteljs.makeStringDataField(key, value);
				break;
			case "double":
				if (typeof value==="string") {
					value=parseFloat(value);
				}
				field=oteljs.makeDoubleDataField(key, value);
				break;
			case "int64":
				if (typeof value==="string") {
					value=parseInt(value);
				}
				field=oteljs.makeInt64DataField(key, value);
				break;
			case "boolean":
				if (typeof value==="string") {
					value=value==="true";
				}
				field=oteljs.makeBooleanDataField(key, value);
				break;
		}
		return field;
	}
	function getEventDefinition(eventName) {
		for (var _i=0; _i < TelemetryEvents.length; _i++) {
			var event_1=TelemetryEvents[_i];
			if (event_1.name===eventName) {
				return event_1;
			}
		}
		return undefined;
	}
	function eventEnabled(eventName) {
		var eventDefinition=getEventDefinition(eventName);
		if (eventDefinition===undefined) {
			return false;
		}
		return eventDefinition.enabled;
	}
	function generateTelemetryEvent(eventName, telemetryData) {
		var eventDefinition=getEventDefinition(eventName);
		if (eventDefinition===undefined) {
			return undefined;
		}
		var dataFields=[];
		for (var _i=0, _a=eventDefinition.points; _i < _a.length; _i++) {
			var point=_a[_i];
			var key=point.name;
			var value=telemetryData[key];
			if (value===undefined) {
				continue;
			}
			var field=createDataField(value, point);
			if (field !==undefined) {
				dataFields.push(field);
			}
		}
		var flags={ dataCategories: oteljs.DataCategories.ProductServiceUsage };
		if (eventDefinition.critical) {
			flags.samplingPolicy=oteljs.SamplingPolicy.CriticalBusinessImpact;
		}
		if (eventDefinition.basic) {
			flags.diagnosticLevel=oteljs.DiagnosticLevel.BasicEvent;
		}
		var eventNameFull="Office.Extensibility.OfficeJs."+eventName+"X";
		var event={ eventName: eventNameFull, dataFields: dataFields, eventFlags: flags };
		return event;
	}
	function sendOtelTelemetryEvent(eventName, telemetryData) {
		if (eventEnabled(eventName)) {
			if (typeof OTel !=="undefined") {
				OTel.OTelLogger.onTelemetryLoaded(function () {
					var event=generateTelemetryEvent(eventName, telemetryData);
					if (event===undefined) {
						return;
					}
					Microsoft.Office.WebExtension.sendTelemetryEvent(event);
				});
			}
		}
	}
	var AriaLogger=(function () {
		function AriaLogger() {
		}
		AriaLogger.prototype.getAriaCDNLocation=function () {
			return (OSF._OfficeAppFactory.getLoadScriptHelper().getOfficeJsBasePath()+"ariatelemetry/aria-web-telemetry.js");
		};
		AriaLogger.getInstance=function () {
			if (AriaLogger.AriaLoggerObj===undefined) {
				AriaLogger.AriaLoggerObj=new AriaLogger();
			}
			return AriaLogger.AriaLoggerObj;
		};
		AriaLogger.prototype.isIUsageData=function (arg) {
			return arg["Fields"] !==undefined;
		};
		AriaLogger.prototype.sendTelemetry=function (tableName, telemetryData) {
			var startAfterMs=1000;
			if (AriaLogger.EnableSendingTelemetryWithLegacyAria) {
				OSF.OUtil.loadScript(this.getAriaCDNLocation(), function () {
					try {
						if (!this.ALogger) {
							var OfficeExtensibilityTenantID="db334b301e7b474db5e0f02f07c51a47-a1b5bc36-1bbe-482f-a64a-c2d9cb606706-7439";
							this.ALogger=AWTLogManager.initialize(OfficeExtensibilityTenantID);
						}
						var eventProperties=new AWTEventProperties();
						eventProperties.setName("Office.Extensibility.OfficeJS."+tableName);
						for (var key in telemetryData) {
							if (key.toLowerCase() !=="table") {
								eventProperties.setProperty(key, telemetryData[key]);
							}
						}
						var today=new Date();
						eventProperties.setProperty("Date", today.toISOString());
						this.ALogger.logEvent(eventProperties);
					}
					catch (e) {
					}
				}, startAfterMs);
			}
			if (AriaLogger.EnableSendingTelemetryWithOTel) {
				sendOtelTelemetryEvent(tableName, telemetryData);
			}
		};
		AriaLogger.prototype.logData=function (data) {
			if (this.isIUsageData(data)) {
				this.sendTelemetry(data["Table"], data["Fields"]);
			}
			else {
				this.sendTelemetry(data["Table"], data);
			}
		};
		AriaLogger.EnableSendingTelemetryWithOTel=true;
		AriaLogger.EnableSendingTelemetryWithLegacyAria=true;
		return AriaLogger;
	})();
	OSFAriaLogger.AriaLogger=AriaLogger;
})(OSFAriaLogger || (OSFAriaLogger={}));
var OSFAppTelemetry;
(function (OSFAppTelemetry) {
	"use strict";
	var appInfo;
	var sessionId=OSF.OUtil.Guid.generateNewGuid();
	var osfControlAppCorrelationId="";
	var omexDomainRegex=new RegExp("^https?://store\\.office(ppe|-int)?\\.com/", "i");
	OSFAppTelemetry.enableTelemetry=true;
	;
	var AppInfo=(function () {
		function AppInfo() {
		}
		return AppInfo;
	})();
	OSFAppTelemetry.AppInfo=AppInfo;
	var Event=(function () {
		function Event(name, handler) {
			this.name=name;
			this.handler=handler;
		}
		return Event;
	})();
	var AppStorage=(function () {
		function AppStorage() {
			this.clientIDKey="Office API client";
			this.logIdSetKey="Office App Log Id Set";
		}
		AppStorage.prototype.getClientId=function () {
			var clientId=this.getValue(this.clientIDKey);
			if (!clientId || clientId.length <=0 || clientId.length > 40) {
				clientId=OSF.OUtil.Guid.generateNewGuid();
				this.setValue(this.clientIDKey, clientId);
			}
			return clientId;
		};
		AppStorage.prototype.saveLog=function (logId, log) {
			var logIdSet=this.getValue(this.logIdSetKey);
			logIdSet=((logIdSet && logIdSet.length > 0) ? (logIdSet+";") : "")+logId;
			this.setValue(this.logIdSetKey, logIdSet);
			this.setValue(logId, log);
		};
		AppStorage.prototype.enumerateLog=function (callback, clean) {
			var logIdSet=this.getValue(this.logIdSetKey);
			if (logIdSet) {
				var ids=logIdSet.split(";");
				for (var id in ids) {
					var logId=ids[id];
					var log=this.getValue(logId);
					if (log) {
						if (callback) {
							callback(logId, log);
						}
						if (clean) {
							this.remove(logId);
						}
					}
				}
				if (clean) {
					this.remove(this.logIdSetKey);
				}
			}
		};
		AppStorage.prototype.getValue=function (key) {
			var osfLocalStorage=OSF.OUtil.getLocalStorage();
			var value="";
			if (osfLocalStorage) {
				value=osfLocalStorage.getItem(key);
			}
			return value;
		};
		AppStorage.prototype.setValue=function (key, value) {
			var osfLocalStorage=OSF.OUtil.getLocalStorage();
			if (osfLocalStorage) {
				osfLocalStorage.setItem(key, value);
			}
		};
		AppStorage.prototype.remove=function (key) {
			var osfLocalStorage=OSF.OUtil.getLocalStorage();
			if (osfLocalStorage) {
				try {
					osfLocalStorage.removeItem(key);
				}
				catch (ex) {
				}
			}
		};
		return AppStorage;
	})();
	var AppLogger=(function () {
		function AppLogger() {
		}
		AppLogger.prototype.LogData=function (data) {
			if (!OSFAppTelemetry.enableTelemetry) {
				return;
			}
			try {
				OSFAriaLogger.AriaLogger.getInstance().logData(data);
			}
			catch (e) {
			}
		};
		AppLogger.prototype.LogRawData=function (log) {
			if (!OSFAppTelemetry.enableTelemetry) {
				return;
			}
			try {
				OSFAriaLogger.AriaLogger.getInstance().logData(JSON.parse(log));
			}
			catch (e) {
			}
		};
		return AppLogger;
	})();
	function trimStringToLowerCase(input) {
		if (input) {
			input=input.replace(/[{}]/g, "").toLowerCase();
		}
		return (input || "");
	}
	var UrlFilter=(function () {
		function UrlFilter() {
		}
		UrlFilter.hashString=function (s) {
			var hash=0;
			if (s.length===0) {
				return hash;
			}
			for (var i=0; i < s.length; i++) {
				var c=s.charCodeAt(i);
				hash=((hash << 5) - hash)+c;
				hash |=0;
			}
			return hash;
		};
		;
		UrlFilter.stringToHash=function (s) {
			var hash=UrlFilter.hashString(s);
			var stringHash=hash.toString();
			if (hash < 0) {
				stringHash="1"+stringHash.substring(1);
			}
			else {
				stringHash="0"+stringHash;
			}
			return stringHash;
		};
		UrlFilter.startsWith=function (s, prefix) {
			return s.indexOf(prefix)==-0;
		};
		UrlFilter.isFileUrl=function (url) {
			return UrlFilter.startsWith(url.toLowerCase(), "file:");
		};
		UrlFilter.removeHttpPrefix=function (url) {
			var prefix="";
			if (UrlFilter.startsWith(url.toLowerCase(), UrlFilter.httpsPrefix)) {
				prefix=UrlFilter.httpsPrefix;
			}
			else if (UrlFilter.startsWith(url.toLowerCase(), UrlFilter.httpPrefix)) {
				prefix=UrlFilter.httpPrefix;
			}
			var clean=url.slice(prefix.length);
			return clean;
		};
		UrlFilter.getUrlDomain=function (url) {
			var domain=UrlFilter.removeHttpPrefix(url);
			domain=domain.split("/")[0];
			domain=domain.split(":")[0];
			return domain;
		};
		UrlFilter.isIp4Address=function (domain) {
			var ipv4Regex=/^(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)$/;
			return ipv4Regex.test(domain);
		};
		UrlFilter.filter=function (url) {
			if (UrlFilter.isFileUrl(url)) {
				var hash=UrlFilter.stringToHash(url);
				return "file://"+hash;
			}
			var domain=UrlFilter.getUrlDomain(url);
			if (UrlFilter.isIp4Address(domain)) {
				var hash=UrlFilter.stringToHash(url);
				if (UrlFilter.startsWith(domain, "10.")) {
					return "IP10Range_"+hash;
				}
				else if (UrlFilter.startsWith(domain, "192.")) {
					return "IP192Range_"+hash;
				}
				else if (UrlFilter.startsWith(domain, "127.")) {
					return "IP127Range_"+hash;
				}
				return "IPOther_"+hash;
			}
			return domain;
		};
		UrlFilter.httpPrefix="http://";
		UrlFilter.httpsPrefix="https://";
		return UrlFilter;
	})();
	function initialize(context) {
		if (!OSFAppTelemetry.enableTelemetry) {
			return;
		}
		if (appInfo) {
			return;
		}
		appInfo=new AppInfo();
		if (context.get_hostFullVersion()) {
			appInfo.hostVersion=context.get_hostFullVersion();
		}
		else {
			appInfo.hostVersion=context.get_appVersion();
		}
		appInfo.appId=context.get_id();
		appInfo.host=context.get_appName();
		appInfo.browser=window.navigator.userAgent;
		appInfo.correlationId=trimStringToLowerCase(context.get_correlationId());
		appInfo.clientId=(new AppStorage()).getClientId();
		appInfo.appInstanceId=context.get_appInstanceId();
		if (appInfo.appInstanceId) {
			appInfo.appInstanceId=appInfo.appInstanceId.replace(/[{}]/g, "").toLowerCase();
		}
		appInfo.message=context.get_hostCustomMessage();
		appInfo.officeJSVersion=OSF.ConstantNames.FileVersion;
		appInfo.hostJSVersion="16.0.11329.10000";
		if (context._wacHostEnvironment) {
			appInfo.wacHostEnvironment=context._wacHostEnvironment;
		}
		if (context._isFromWacAutomation !==undefined && context._isFromWacAutomation !==null) {
			appInfo.isFromWacAutomation=context._isFromWacAutomation.toString().toLowerCase();
		}
		var docUrl=context.get_docUrl();
		appInfo.docUrl=omexDomainRegex.test(docUrl) ? docUrl : "";
		var url=location.href;
		if (url) {
			url=url.split("?")[0].split("#")[0];
		}
		appInfo.appURL=UrlFilter.filter(url);
		(function getUserIdAndAssetIdFromToken(token, appInfo) {
			var xmlContent;
			var parser;
			var xmlDoc;
			appInfo.assetId="";
			appInfo.userId="";
			try {
				xmlContent=decodeURIComponent(token);
				parser=new DOMParser();
				xmlDoc=parser.parseFromString(xmlContent, "text/xml");
				var cidNode=xmlDoc.getElementsByTagName("t")[0].attributes.getNamedItem("cid");
				var oidNode=xmlDoc.getElementsByTagName("t")[0].attributes.getNamedItem("oid");
				if (cidNode && cidNode.nodeValue) {
					appInfo.userId=cidNode.nodeValue;
				}
				else if (oidNode && oidNode.nodeValue) {
					appInfo.userId=oidNode.nodeValue;
				}
				appInfo.assetId=xmlDoc.getElementsByTagName("t")[0].attributes.getNamedItem("aid").nodeValue;
			}
			catch (e) {
			}
			finally {
				xmlContent=null;
				xmlDoc=null;
				parser=null;
			}
		})(context.get_eToken(), appInfo);
		appInfo.sessionId=sessionId;
		appInfo.name=context.get_addinName();
		if (typeof OTel !=="undefined") {
			OTel.OTelLogger.initialize(appInfo);
		}
		(function handleLifecycle() {
			var startTime=new Date();
			var lastFocus=null;
			var focusTime=0;
			var finished=false;
			var adjustFocusTime=function () {
				if (document.hasFocus()) {
					if (lastFocus==null) {
						lastFocus=new Date();
					}
				}
				else if (lastFocus) {
					focusTime+=Math.abs((new Date()).getTime() - lastFocus.getTime());
					lastFocus=null;
				}
			};
			var eventList=[];
			eventList.push(new Event("focus", adjustFocusTime));
			eventList.push(new Event("blur", adjustFocusTime));
			eventList.push(new Event("focusout", adjustFocusTime));
			eventList.push(new Event("focusin", adjustFocusTime));
			var exitFunction=function () {
				for (var i=0; i < eventList.length; i++) {
					OSF.OUtil.removeEventListener(window, eventList[i].name, eventList[i].handler);
				}
				eventList.length=0;
				if (!finished) {
					if (document.hasFocus() && lastFocus) {
						focusTime+=Math.abs((new Date()).getTime() - lastFocus.getTime());
						lastFocus=null;
					}
					OSFAppTelemetry.onAppClosed(Math.abs((new Date()).getTime() - startTime.getTime()), focusTime);
					finished=true;
				}
			};
			eventList.push(new Event("beforeunload", exitFunction));
			eventList.push(new Event("unload", exitFunction));
			for (var i=0; i < eventList.length; i++) {
				OSF.OUtil.addEventListener(window, eventList[i].name, eventList[i].handler);
			}
			adjustFocusTime();
		})();
		OSFAppTelemetry.onAppActivated();
	}
	OSFAppTelemetry.initialize=initialize;
	function onAppActivated() {
		if (!appInfo) {
			return;
		}
		(new AppStorage()).enumerateLog(function (id, log) { return (new AppLogger()).LogRawData(log); }, true);
		var data=new OSFLog.AppActivatedUsageData();
		data.SessionId=sessionId;
		data.AppId=appInfo.appId;
		data.AssetId=appInfo.assetId;
		data.AppURL=appInfo.appURL;
		data.UserId="";
		data.ClientId=appInfo.clientId;
		data.Browser=appInfo.browser;
		data.Host=appInfo.host;
		data.HostVersion=appInfo.hostVersion;
		data.CorrelationId=trimStringToLowerCase(appInfo.correlationId);
		data.AppSizeWidth=window.innerWidth;
		data.AppSizeHeight=window.innerHeight;
		data.AppInstanceId=appInfo.appInstanceId;
		data.Message=appInfo.message;
		data.DocUrl=appInfo.docUrl;
		data.OfficeJSVersion=appInfo.officeJSVersion;
		data.HostJSVersion=appInfo.hostJSVersion;
		if (appInfo.wacHostEnvironment) {
			data.WacHostEnvironment=appInfo.wacHostEnvironment;
		}
		if (appInfo.isFromWacAutomation !==undefined && appInfo.isFromWacAutomation !==null) {
			data.IsFromWacAutomation=appInfo.isFromWacAutomation;
		}
		(new AppLogger()).LogData(data);
	}
	OSFAppTelemetry.onAppActivated=onAppActivated;
	function onScriptDone(scriptId, msStartTime, msResponseTime, appCorrelationId) {
		var data=new OSFLog.ScriptLoadUsageData();
		data.CorrelationId=trimStringToLowerCase(appCorrelationId);
		data.SessionId=sessionId;
		data.ScriptId=scriptId;
		data.StartTime=msStartTime;
		data.ResponseTime=msResponseTime;
		(new AppLogger()).LogData(data);
	}
	OSFAppTelemetry.onScriptDone=onScriptDone;
	function onCallDone(apiType, id, parameters, msResponseTime, errorType) {
		if (!appInfo) {
			return;
		}
		var data=new OSFLog.APIUsageUsageData();
		data.CorrelationId=trimStringToLowerCase(osfControlAppCorrelationId);
		data.SessionId=sessionId;
		data.APIType=apiType;
		data.APIID=id;
		data.Parameters=parameters;
		data.ResponseTime=msResponseTime;
		data.ErrorType=errorType;
		(new AppLogger()).LogData(data);
	}
	OSFAppTelemetry.onCallDone=onCallDone;
	;
	function onMethodDone(id, args, msResponseTime, errorType) {
		var parameters=null;
		if (args) {
			if (typeof args=="number") {
				parameters=String(args);
			}
			else if (typeof args==="object") {
				for (var index in args) {
					if (parameters !==null) {
						parameters+=",";
					}
					else {
						parameters="";
					}
					if (typeof args[index]=="number") {
						parameters+=String(args[index]);
					}
				}
			}
			else {
				parameters="";
			}
		}
		OSF.AppTelemetry.onCallDone("method", id, parameters, msResponseTime, errorType);
	}
	OSFAppTelemetry.onMethodDone=onMethodDone;
	function onPropertyDone(propertyName, msResponseTime) {
		OSF.AppTelemetry.onCallDone("property", -1, propertyName, msResponseTime);
	}
	OSFAppTelemetry.onPropertyDone=onPropertyDone;
	function onEventDone(id, errorType) {
		OSF.AppTelemetry.onCallDone("event", id, null, 0, errorType);
	}
	OSFAppTelemetry.onEventDone=onEventDone;
	function onRegisterDone(register, id, msResponseTime, errorType) {
		OSF.AppTelemetry.onCallDone(register ? "registerevent" : "unregisterevent", id, null, msResponseTime, errorType);
	}
	OSFAppTelemetry.onRegisterDone=onRegisterDone;
	function onAppClosed(openTime, focusTime) {
		if (!appInfo) {
			return;
		}
		var data=new OSFLog.AppClosedUsageData();
		data.CorrelationId=trimStringToLowerCase(osfControlAppCorrelationId);
		data.SessionId=sessionId;
		data.FocusTime=focusTime;
		data.OpenTime=openTime;
		data.AppSizeFinalWidth=window.innerWidth;
		data.AppSizeFinalHeight=window.innerHeight;
		(new AppStorage()).saveLog(sessionId, data.SerializeRow());
	}
	OSFAppTelemetry.onAppClosed=onAppClosed;
	function setOsfControlAppCorrelationId(correlationId) {
		osfControlAppCorrelationId=trimStringToLowerCase(correlationId);
	}
	OSFAppTelemetry.setOsfControlAppCorrelationId=setOsfControlAppCorrelationId;
	function doAppInitializationLogging(isException, message) {
		var data=new OSFLog.AppInitializationUsageData();
		data.CorrelationId=trimStringToLowerCase(osfControlAppCorrelationId);
		data.SessionId=sessionId;
		data.SuccessCode=isException ? 1 : 0;
		data.Message=message;
		(new AppLogger()).LogData(data);
	}
	OSFAppTelemetry.doAppInitializationLogging=doAppInitializationLogging;
	function logAppCommonMessage(message) {
		doAppInitializationLogging(false, message);
	}
	OSFAppTelemetry.logAppCommonMessage=logAppCommonMessage;
	function logAppException(errorMessage) {
		doAppInitializationLogging(true, errorMessage);
	}
	OSFAppTelemetry.logAppException=logAppException;
	OSF.AppTelemetry=OSFAppTelemetry;
})(OSFAppTelemetry || (OSFAppTelemetry={}));
Microsoft.Office.WebExtension.EventType={};
OSF.EventDispatch=function OSF_EventDispatch(eventTypes) {
	this._eventHandlers={};
	this._objectEventHandlers={};
	this._queuedEventsArgs={};
	if (eventTypes !=null) {
		for (var i=0; i < eventTypes.length; i++) {
			var eventType=eventTypes[i];
			var isObjectEvent=(eventType=="objectDeleted" || eventType=="objectSelectionChanged" || eventType=="objectDataChanged" || eventType=="contentControlAdded");
			if (!isObjectEvent)
				this._eventHandlers[eventType]=[];
			else
				this._objectEventHandlers[eventType]={};
			this._queuedEventsArgs[eventType]=[];
		}
	}
};
OSF.EventDispatch.prototype={
	getSupportedEvents: function OSF_EventDispatch$getSupportedEvents() {
		var events=[];
		for (var eventName in this._eventHandlers)
			events.push(eventName);
		for (var eventName in this._objectEventHandlers)
			events.push(eventName);
		return events;
	},
	supportsEvent: function OSF_EventDispatch$supportsEvent(event) {
		for (var eventName in this._eventHandlers) {
			if (event==eventName)
				return true;
		}
		for (var eventName in this._objectEventHandlers) {
			if (event==eventName)
				return true;
		}
		return false;
	},
	hasEventHandler: function OSF_EventDispatch$hasEventHandler(eventType, handler) {
		var handlers=this._eventHandlers[eventType];
		if (handlers && handlers.length > 0) {
			for (var i=0; i < handlers.length; i++) {
				if (handlers[i]===handler)
					return true;
			}
		}
		return false;
	},
	hasObjectEventHandler: function OSF_EventDispatch$hasObjectEventHandler(eventType, objectId, handler) {
		var handlers=this._objectEventHandlers[eventType];
		if (handlers !=null) {
			var _handlers=handlers[objectId];
			for (var i=0; _handlers !=null && i < _handlers.length; i++) {
				if (_handlers[i]===handler)
					return true;
			}
		}
		return false;
	},
	addEventHandler: function OSF_EventDispatch$addEventHandler(eventType, handler) {
		if (typeof handler !="function") {
			return false;
		}
		var handlers=this._eventHandlers[eventType];
		if (handlers && !this.hasEventHandler(eventType, handler)) {
			handlers.push(handler);
			return true;
		}
		else {
			return false;
		}
	},
	addObjectEventHandler: function OSF_EventDispatch$addObjectEventHandler(eventType, objectId, handler) {
		if (typeof handler !="function") {
			return false;
		}
		var handlers=this._objectEventHandlers[eventType];
		if (handlers && !this.hasObjectEventHandler(eventType, objectId, handler)) {
			if (handlers[objectId]==null)
				handlers[objectId]=[];
			handlers[objectId].push(handler);
			return true;
		}
		return false;
	},
	addEventHandlerAndFireQueuedEvent: function OSF_EventDispatch$addEventHandlerAndFireQueuedEvent(eventType, handler) {
		var handlers=this._eventHandlers[eventType];
		var isFirstHandler=handlers.length==0;
		var succeed=this.addEventHandler(eventType, handler);
		if (isFirstHandler && succeed) {
			this.fireQueuedEvent(eventType);
		}
		return succeed;
	},
	removeEventHandler: function OSF_EventDispatch$removeEventHandler(eventType, handler) {
		var handlers=this._eventHandlers[eventType];
		if (handlers && handlers.length > 0) {
			for (var index=0; index < handlers.length; index++) {
				if (handlers[index]===handler) {
					handlers.splice(index, 1);
					return true;
				}
			}
		}
		return false;
	},
	removeObjectEventHandler: function OSF_EventDispatch$removeObjectEventHandler(eventType, objectId, handler) {
		var handlers=this._objectEventHandlers[eventType];
		if (handlers !=null) {
			var _handlers=handlers[objectId];
			for (var i=0; _handlers !=null && i < _handlers.length; i++) {
				if (_handlers[i]===handler) {
					_handlers.splice(i, 1);
					return true;
				}
			}
		}
		return false;
	},
	clearEventHandlers: function OSF_EventDispatch$clearEventHandlers(eventType) {
		if (typeof this._eventHandlers[eventType] !="undefined" && this._eventHandlers[eventType].length > 0) {
			this._eventHandlers[eventType]=[];
			return true;
		}
		return false;
	},
	clearObjectEventHandlers: function OSF_EventDispatch$clearObjectEventHandlers(eventType, objectId) {
		if (this._objectEventHandlers[eventType] !=null && this._objectEventHandlers[eventType][objectId] !=null) {
			this._objectEventHandlers[eventType][objectId]=[];
			return true;
		}
		return false;
	},
	getEventHandlerCount: function OSF_EventDispatch$getEventHandlerCount(eventType) {
		return this._eventHandlers[eventType] !=undefined ? this._eventHandlers[eventType].length : -1;
	},
	getObjectEventHandlerCount: function OSF_EventDispatch$getObjectEventHandlerCount(eventType, objectId) {
		if (this._objectEventHandlers[eventType]==null || this._objectEventHandlers[eventType][objectId]==null)
			return 0;
		return this._objectEventHandlers[eventType][objectId].length;
	},
	fireEvent: function OSF_EventDispatch$fireEvent(eventArgs) {
		if (eventArgs.type==undefined)
			return false;
		var eventType=eventArgs.type;
		if (eventType && this._eventHandlers[eventType]) {
			var eventHandlers=this._eventHandlers[eventType];
			for (var i=0; i < eventHandlers.length; i++) {
				eventHandlers[i](eventArgs);
			}
			return true;
		}
		else {
			return false;
		}
	},
	fireObjectEvent: function OSF_EventDispatch$fireObjectEvent(objectId, eventArgs) {
		if (eventArgs.type==undefined)
			return false;
		var eventType=eventArgs.type;
		if (eventType && this._objectEventHandlers[eventType]) {
			var eventHandlers=this._objectEventHandlers[eventType];
			var _handlers=eventHandlers[objectId];
			if (_handlers !=null) {
				for (var i=0; i < _handlers.length; i++)
					_handlers[i](eventArgs);
				return true;
			}
		}
		return false;
	},
	fireOrQueueEvent: function OSF_EventDispatch$fireOrQueueEvent(eventArgs) {
		var eventType=eventArgs.type;
		if (eventType && this._eventHandlers[eventType]) {
			var eventHandlers=this._eventHandlers[eventType];
			var queuedEvents=this._queuedEventsArgs[eventType];
			if (eventHandlers.length==0) {
				queuedEvents.push(eventArgs);
			}
			else {
				this.fireEvent(eventArgs);
			}
			return true;
		}
		else {
			return false;
		}
	},
	fireQueuedEvent: function OSF_EventDispatch$queueEvent(eventType) {
		if (eventType && this._eventHandlers[eventType]) {
			var eventHandlers=this._eventHandlers[eventType];
			var queuedEvents=this._queuedEventsArgs[eventType];
			if (eventHandlers.length > 0) {
				var eventHandler=eventHandlers[0];
				while (queuedEvents.length > 0) {
					var eventArgs=queuedEvents.shift();
					eventHandler(eventArgs);
				}
				return true;
			}
		}
		return false;
	},
	clearQueuedEvent: function OSF_EventDispatch$clearQueuedEvent(eventType) {
		if (eventType && this._eventHandlers[eventType]) {
			var queuedEvents=this._queuedEventsArgs[eventType];
			if (queuedEvents) {
				this._queuedEventsArgs[eventType]=[];
			}
		}
	}
};
OSF.DDA.OMFactory=OSF.DDA.OMFactory || {};
OSF.DDA.OMFactory.manufactureEventArgs=function OSF_DDA_OMFactory$manufactureEventArgs(eventType, target, eventProperties) {
	var args;
	switch (eventType) {
		case Microsoft.Office.WebExtension.EventType.DocumentSelectionChanged:
			args=new OSF.DDA.DocumentSelectionChangedEventArgs(target);
			break;
		case Microsoft.Office.WebExtension.EventType.BindingSelectionChanged:
			args=new OSF.DDA.BindingSelectionChangedEventArgs(this.manufactureBinding(eventProperties, target.document), eventProperties[OSF.DDA.PropertyDescriptors.Subset]);
			break;
		case Microsoft.Office.WebExtension.EventType.BindingDataChanged:
			args=new OSF.DDA.BindingDataChangedEventArgs(this.manufactureBinding(eventProperties, target.document));
			break;
		case Microsoft.Office.WebExtension.EventType.SettingsChanged:
			args=new OSF.DDA.SettingsChangedEventArgs(target);
			break;
		case Microsoft.Office.WebExtension.EventType.ActiveViewChanged:
			args=new OSF.DDA.ActiveViewChangedEventArgs(eventProperties);
			break;
		case Microsoft.Office.WebExtension.EventType.OfficeThemeChanged:
			args=new OSF.DDA.Theming.OfficeThemeChangedEventArgs(eventProperties);
			break;
		case Microsoft.Office.WebExtension.EventType.DocumentThemeChanged:
			args=new OSF.DDA.Theming.DocumentThemeChangedEventArgs(eventProperties);
			break;
		case Microsoft.Office.WebExtension.EventType.AppCommandInvoked:
			args=OSF.DDA.AppCommand.AppCommandInvokedEventArgs.create(eventProperties);
			break;
		case Microsoft.Office.WebExtension.EventType.ObjectDeleted:
		case Microsoft.Office.WebExtension.EventType.ObjectSelectionChanged:
		case Microsoft.Office.WebExtension.EventType.ObjectDataChanged:
		case Microsoft.Office.WebExtension.EventType.ContentControlAdded:
			args=new OSF.DDA.ObjectEventArgs(eventType, eventProperties[Microsoft.Office.WebExtension.Parameters.Id]);
			break;
		case Microsoft.Office.WebExtension.EventType.RichApiMessage:
			args=new OSF.DDA.RichApiMessageEventArgs(eventType, eventProperties);
			break;
		case Microsoft.Office.WebExtension.EventType.DataNodeInserted:
			args=new OSF.DDA.NodeInsertedEventArgs(this.manufactureDataNode(eventProperties[OSF.DDA.DataNodeEventProperties.NewNode]), eventProperties[OSF.DDA.DataNodeEventProperties.InUndoRedo]);
			break;
		case Microsoft.Office.WebExtension.EventType.DataNodeReplaced:
			args=new OSF.DDA.NodeReplacedEventArgs(this.manufactureDataNode(eventProperties[OSF.DDA.DataNodeEventProperties.OldNode]), this.manufactureDataNode(eventProperties[OSF.DDA.DataNodeEventProperties.NewNode]), eventProperties[OSF.DDA.DataNodeEventProperties.InUndoRedo]);
			break;
		case Microsoft.Office.WebExtension.EventType.DataNodeDeleted:
			args=new OSF.DDA.NodeDeletedEventArgs(this.manufactureDataNode(eventProperties[OSF.DDA.DataNodeEventProperties.OldNode]), this.manufactureDataNode(eventProperties[OSF.DDA.DataNodeEventProperties.NextSiblingNode]), eventProperties[OSF.DDA.DataNodeEventProperties.InUndoRedo]);
			break;
		case Microsoft.Office.WebExtension.EventType.TaskSelectionChanged:
			args=new OSF.DDA.TaskSelectionChangedEventArgs(target);
			break;
		case Microsoft.Office.WebExtension.EventType.ResourceSelectionChanged:
			args=new OSF.DDA.ResourceSelectionChangedEventArgs(target);
			break;
		case Microsoft.Office.WebExtension.EventType.ViewSelectionChanged:
			args=new OSF.DDA.ViewSelectionChangedEventArgs(target);
			break;
		case Microsoft.Office.WebExtension.EventType.DialogMessageReceived:
			args=new OSF.DDA.DialogEventArgs(eventProperties);
			break;
		case Microsoft.Office.WebExtension.EventType.DialogParentMessageReceived:
			args=new OSF.DDA.DialogParentEventArgs(eventProperties);
			break;
		case Microsoft.Office.WebExtension.EventType.ItemChanged:
			if (OSF._OfficeAppFactory.getHostInfo()["hostType"]=="outlook") {
				args=new OSF.DDA.OlkItemSelectedChangedEventArgs(eventProperties);
				target.initialize(args["initialData"]);
				if (OSF._OfficeAppFactory.getHostInfo()["hostPlatform"]=="win32" || OSF._OfficeAppFactory.getHostInfo()["hostPlatform"]=="mac") {
					target.setCurrentItemNumber(args["itemNumber"].itemNumber);
				}
			}
			else {
				throw OsfMsAjaxFactory.msAjaxError.argument(Microsoft.Office.WebExtension.Parameters.EventType, OSF.OUtil.formatString(Strings.OfficeOM.L_NotSupportedEventType, eventType));
			}
			break;
		case Microsoft.Office.WebExtension.EventType.RecipientsChanged:
			if (OSF._OfficeAppFactory.getHostInfo()["hostType"]=="outlook") {
				args=new OSF.DDA.OlkRecipientsChangedEventArgs(eventProperties);
			}
			else {
				throw OsfMsAjaxFactory.msAjaxError.argument(Microsoft.Office.WebExtension.Parameters.EventType, OSF.OUtil.formatString(Strings.OfficeOM.L_NotSupportedEventType, eventType));
			}
			break;
		case Microsoft.Office.WebExtension.EventType.AppointmentTimeChanged:
			if (OSF._OfficeAppFactory.getHostInfo()["hostType"]=="outlook") {
				args=new OSF.DDA.OlkAppointmentTimeChangedEventArgs(eventProperties);
			}
			else {
				throw OsfMsAjaxFactory.msAjaxError.argument(Microsoft.Office.WebExtension.Parameters.EventType, OSF.OUtil.formatString(Strings.OfficeOM.L_NotSupportedEventType, eventType));
			}
			break;
		case Microsoft.Office.WebExtension.EventType.RecurrenceChanged:
			if (OSF._OfficeAppFactory.getHostInfo()["hostType"]=="outlook") {
				args=new OSF.DDA.OlkRecurrenceChangedEventArgs(eventProperties);
			}
			else {
				throw OsfMsAjaxFactory.msAjaxError.argument(Microsoft.Office.WebExtension.Parameters.EventType, OSF.OUtil.formatString(Strings.OfficeOM.L_NotSupportedEventType, eventType));
			}
			break;
		case Microsoft.Office.WebExtension.EventType.AttachmentsChanged:
			if (OSF._OfficeAppFactory.getHostInfo()["hostType"]=="outlook") {
				args=new OSF.DDA.OlkAttachmentsChangedEventArgs(eventProperties);
			}
			else {
				throw OsfMsAjaxFactory.msAjaxError.argument(Microsoft.Office.WebExtension.Parameters.EventType, OSF.OUtil.formatString(Strings.OfficeOM.L_NotSupportedEventType, eventType));
			}
			break;
		case Microsoft.Office.WebExtension.EventType.EnhancedLocationsChanged:
			if (OSF._OfficeAppFactory.getHostInfo()["hostType"]=="outlook") {
				args=new OSF.DDA.OlkEnhancedLocationsChangedEventArgs(eventProperties);
			}
			else {
				throw OsfMsAjaxFactory.msAjaxError.argument(Microsoft.Office.WebExtension.Parameters.EventType, OSF.OUtil.formatString(Strings.OfficeOM.L_NotSupportedEventType, eventType));
			}
			break;
		case Microsoft.Office.WebExtension.EventType.InfobarClicked:
			if (OSF._OfficeAppFactory.getHostInfo()["hostType"]=="outlook") {
				args=new OSF.DDA.OlkInfobarClickedEventArgs(eventProperties);
			}
			else {
				throw OsfMsAjaxFactory.msAjaxError.argument(Microsoft.Office.WebExtension.Parameters.EventType, OSF.OUtil.formatString(Strings.OfficeOM.L_NotSupportedEventType, eventType));
			}
			break;
		default:
			throw OsfMsAjaxFactory.msAjaxError.argument(Microsoft.Office.WebExtension.Parameters.EventType, OSF.OUtil.formatString(Strings.OfficeOM.L_NotSupportedEventType, eventType));
	}
	return args;
};
OSF.DDA.AsyncMethodNames.addNames({
	AddHandlerAsync: "addHandlerAsync",
	RemoveHandlerAsync: "removeHandlerAsync"
});
OSF.DDA.AsyncMethodCalls.define({
	method: OSF.DDA.AsyncMethodNames.AddHandlerAsync,
	requiredArguments: [{
			"name": Microsoft.Office.WebExtension.Parameters.EventType,
			"enum": Microsoft.Office.WebExtension.EventType,
			"verify": function (eventType, caller, eventDispatch) { return eventDispatch.supportsEvent(eventType); }
		},
		{
			"name": Microsoft.Office.WebExtension.Parameters.Handler,
			"types": ["function"]
		}
	],
	supportedOptions: [],
	privateStateCallbacks: []
});
OSF.DDA.AsyncMethodCalls.define({
	method: OSF.DDA.AsyncMethodNames.RemoveHandlerAsync,
	requiredArguments: [
		{
			"name": Microsoft.Office.WebExtension.Parameters.EventType,
			"enum": Microsoft.Office.WebExtension.EventType,
			"verify": function (eventType, caller, eventDispatch) { return eventDispatch.supportsEvent(eventType); }
		}
	],
	supportedOptions: [
		{
			name: Microsoft.Office.WebExtension.Parameters.Handler,
			value: {
				"types": ["function", "object"],
				"defaultValue": null
			}
		}
	],
	privateStateCallbacks: []
});
OSF.DialogShownStatus={ hasDialogShown: false, isWindowDialog: false };
OSF.OUtil.augmentList(OSF.DDA.EventDescriptors, {
	DialogMessageReceivedEvent: "DialogMessageReceivedEvent"
});
OSF.OUtil.augmentList(Microsoft.Office.WebExtension.EventType, {
	DialogMessageReceived: "dialogMessageReceived",
	DialogEventReceived: "dialogEventReceived"
});
OSF.OUtil.augmentList(OSF.DDA.PropertyDescriptors, {
	MessageType: "messageType",
	MessageContent: "messageContent"
});
OSF.DDA.DialogEventType={};
OSF.OUtil.augmentList(OSF.DDA.DialogEventType, {
	DialogClosed: "dialogClosed",
	NavigationFailed: "naviationFailed"
});
OSF.DDA.AsyncMethodNames.addNames({
	DisplayDialogAsync: "displayDialogAsync",
	CloseAsync: "close"
});
OSF.DDA.SyncMethodNames.addNames({
	MessageParent: "messageParent",
	AddMessageHandler: "addEventHandler",
	SendMessage: "sendMessage"
});
OSF.DDA.UI.ParentUI=function OSF_DDA_ParentUI() {
	var eventDispatch;
	if (Microsoft.Office.WebExtension.EventType.DialogParentMessageReceived !=null) {
		eventDispatch=new OSF.EventDispatch([
			Microsoft.Office.WebExtension.EventType.DialogMessageReceived,
			Microsoft.Office.WebExtension.EventType.DialogEventReceived,
			Microsoft.Office.WebExtension.EventType.DialogParentMessageReceived
		]);
	}
	else {
		eventDispatch=new OSF.EventDispatch([
			Microsoft.Office.WebExtension.EventType.DialogMessageReceived,
			Microsoft.Office.WebExtension.EventType.DialogEventReceived
		]);
	}
	var openDialogName=OSF.DDA.AsyncMethodNames.DisplayDialogAsync.displayName;
	var target=this;
	if (!target[openDialogName]) {
		OSF.OUtil.defineEnumerableProperty(target, openDialogName, {
			value: function () {
				var openDialog=OSF._OfficeAppFactory.getHostFacade()[OSF.DDA.DispIdHost.Methods.OpenDialog];
				openDialog(arguments, eventDispatch, target);
			}
		});
	}
	OSF.OUtil.finalizeProperties(this);
};
OSF.DDA.UI.ChildUI=function OSF_DDA_ChildUI(isPopupWindow) {
	var messageParentName=OSF.DDA.SyncMethodNames.MessageParent.displayName;
	var target=this;
	if (!target[messageParentName]) {
		OSF.OUtil.defineEnumerableProperty(target, messageParentName, {
			value: function () {
				var messageParent=OSF._OfficeAppFactory.getHostFacade()[OSF.DDA.DispIdHost.Methods.MessageParent];
				return messageParent(arguments, target);
			}
		});
	}
	var addEventHandler=OSF.DDA.SyncMethodNames.AddMessageHandler.displayName;
	if (!target[addEventHandler] && typeof OSF.DialogParentMessageEventDispatch !="undefined") {
		OSF.DDA.DispIdHost.addEventSupport(target, OSF.DialogParentMessageEventDispatch, isPopupWindow);
	}
	OSF.OUtil.finalizeProperties(this);
};
OSF.DialogHandler=function OSF_DialogHandler() { };
OSF.DDA.DialogEventArgs=function OSF_DDA_DialogEventArgs(message) {
	if (message[OSF.DDA.PropertyDescriptors.MessageType]==OSF.DialogMessageType.DialogMessageReceived) {
		OSF.OUtil.defineEnumerableProperties(this, {
			"type": {
				value: Microsoft.Office.WebExtension.EventType.DialogMessageReceived
			},
			"message": {
				value: message[OSF.DDA.PropertyDescriptors.MessageContent]
			}
		});
	}
	else {
		OSF.OUtil.defineEnumerableProperties(this, {
			"type": {
				value: Microsoft.Office.WebExtension.EventType.DialogEventReceived
			},
			"error": {
				value: message[OSF.DDA.PropertyDescriptors.MessageType]
			}
		});
	}
};
OSF.DDA.DialogParentEventArgs=function OSF_DDA_DialogParentEventArgs(message) {
	OSF.OUtil.defineEnumerableProperties(this, {
		"type": {
			value: Microsoft.Office.WebExtension.EventType.DialogParentMessageReceived
		},
		"message": {
			value: message[OSF.DDA.PropertyDescriptors.MessageContent]
		}
	});
};
OSF.DDA.AsyncMethodCalls.define({
	method: OSF.DDA.AsyncMethodNames.DisplayDialogAsync,
	requiredArguments: [
		{
			"name": Microsoft.Office.WebExtension.Parameters.Url,
			"types": ["string"]
		}
	],
	supportedOptions: [
		{
			name: Microsoft.Office.WebExtension.Parameters.Width,
			value: {
				"types": ["number"],
				"defaultValue": 99
			}
		},
		{
			name: Microsoft.Office.WebExtension.Parameters.Height,
			value: {
				"types": ["number"],
				"defaultValue": 99
			}
		},
		{
			name: Microsoft.Office.WebExtension.Parameters.RequireHTTPs,
			value: {
				"types": ["boolean"],
				"defaultValue": true
			}
		},
		{
			name: Microsoft.Office.WebExtension.Parameters.DisplayInIframe,
			value: {
				"types": ["boolean"],
				"defaultValue": false
			}
		},
		{
			name: Microsoft.Office.WebExtension.Parameters.HideTitle,
			value: {
				"types": ["boolean"],
				"defaultValue": false
			}
		},
		{
			name: Microsoft.Office.WebExtension.Parameters.UseDeviceIndependentPixels,
			value: {
				"types": ["boolean"],
				"defaultValue": false
			}
		},
		{
			name: Microsoft.Office.WebExtension.Parameters.PromptBeforeOpen,
			value: {
				"types": ["boolean"],
				"defaultValue": true
			}
		}
	],
	privateStateCallbacks: [],
	onSucceeded: function (args, caller, callArgs) {
		var targetId=args[Microsoft.Office.WebExtension.Parameters.Id];
		var eventDispatch=args[Microsoft.Office.WebExtension.Parameters.Data];
		var dialog=new OSF.DialogHandler();
		var closeDialog=OSF.DDA.AsyncMethodNames.CloseAsync.displayName;
		OSF.OUtil.defineEnumerableProperty(dialog, closeDialog, {
			value: function () {
				var closeDialogfunction=OSF._OfficeAppFactory.getHostFacade()[OSF.DDA.DispIdHost.Methods.CloseDialog];
				closeDialogfunction(arguments, targetId, eventDispatch, dialog);
			}
		});
		var addHandler=OSF.DDA.SyncMethodNames.AddMessageHandler.displayName;
		OSF.OUtil.defineEnumerableProperty(dialog, addHandler, {
			value: function () {
				var syncMethodCall=OSF.DDA.SyncMethodCalls[OSF.DDA.SyncMethodNames.AddMessageHandler.id];
				var callArgs=syncMethodCall.verifyAndExtractCall(arguments, dialog, eventDispatch);
				var eventType=callArgs[Microsoft.Office.WebExtension.Parameters.EventType];
				var handler=callArgs[Microsoft.Office.WebExtension.Parameters.Handler];
				return eventDispatch.addEventHandlerAndFireQueuedEvent(eventType, handler);
			}
		});
		var sendMessage=OSF.DDA.SyncMethodNames.SendMessage.displayName;
		OSF.OUtil.defineEnumerableProperty(dialog, sendMessage, {
			value: function () {
				var execute=OSF._OfficeAppFactory.getHostFacade()[OSF.DDA.DispIdHost.Methods.SendMessage];
				return execute(arguments, eventDispatch, dialog);
			}
		});
		return dialog;
	},
	checkCallArgs: function (callArgs, caller, stateInfo) {
		if (callArgs[Microsoft.Office.WebExtension.Parameters.Width] <=0) {
			callArgs[Microsoft.Office.WebExtension.Parameters.Width]=1;
		}
		if (!callArgs[Microsoft.Office.WebExtension.Parameters.UseDeviceIndependentPixels] && callArgs[Microsoft.Office.WebExtension.Parameters.Width] > 100) {
			callArgs[Microsoft.Office.WebExtension.Parameters.Width]=99;
		}
		if (callArgs[Microsoft.Office.WebExtension.Parameters.Height] <=0) {
			callArgs[Microsoft.Office.WebExtension.Parameters.Height]=1;
		}
		if (!callArgs[Microsoft.Office.WebExtension.Parameters.UseDeviceIndependentPixels] && callArgs[Microsoft.Office.WebExtension.Parameters.Height] > 100) {
			callArgs[Microsoft.Office.WebExtension.Parameters.Height]=99;
		}
		if (!callArgs[Microsoft.Office.WebExtension.Parameters.RequireHTTPs]) {
			callArgs[Microsoft.Office.WebExtension.Parameters.RequireHTTPs]=true;
		}
		return callArgs;
	}
});
OSF.DDA.AsyncMethodCalls.define({
	method: OSF.DDA.AsyncMethodNames.CloseAsync,
	requiredArguments: [],
	supportedOptions: [],
	privateStateCallbacks: []
});
OSF.DDA.SyncMethodCalls.define({
	method: OSF.DDA.SyncMethodNames.MessageParent,
	requiredArguments: [
		{
			"name": Microsoft.Office.WebExtension.Parameters.MessageToParent,
			"types": ["string", "number", "boolean"]
		}
	],
	supportedOptions: []
});
OSF.DDA.SyncMethodCalls.define({
	method: OSF.DDA.SyncMethodNames.AddMessageHandler,
	requiredArguments: [
		{
			"name": Microsoft.Office.WebExtension.Parameters.EventType,
			"enum": Microsoft.Office.WebExtension.EventType,
			"verify": function (eventType, caller, eventDispatch) { return eventDispatch.supportsEvent(eventType); }
		},
		{
			"name": Microsoft.Office.WebExtension.Parameters.Handler,
			"types": ["function"]
		}
	],
	supportedOptions: []
});
OSF.DDA.SyncMethodCalls.define({
	method: OSF.DDA.SyncMethodNames.SendMessage,
	requiredArguments: [
		{
			"name": Microsoft.Office.WebExtension.Parameters.MessageContent,
			"types": ["string"]
		}
	],
	supportedOptions: [],
	privateStateCallbacks: []
});
OSF.DDA.SafeArray.Delegate.openDialog=function OSF_DDA_SafeArray_Delegate$OpenDialog(args) {
	try {
		if (args.onCalling) {
			args.onCalling();
		}
		var callback=OSF.DDA.SafeArray.Delegate._getOnAfterRegisterEvent(true, args);
		OSF.ClientHostController.openDialog(args.dispId, args.targetId, function OSF_DDA_SafeArrayDelegate$RegisterEventAsync_OnEvent(eventDispId, payload) {
			if (args.onEvent) {
				args.onEvent(payload);
			}
			if (OSF.AppTelemetry) {
				OSF.AppTelemetry.onEventDone(args.dispId);
			}
		}, callback);
	}
	catch (ex) {
		OSF.DDA.SafeArray.Delegate._onException(ex, args);
	}
};
OSF.DDA.SafeArray.Delegate.closeDialog=function OSF_DDA_SafeArray_Delegate$CloseDialog(args) {
	if (args.onCalling) {
		args.onCalling();
	}
	var callback=OSF.DDA.SafeArray.Delegate._getOnAfterRegisterEvent(false, args);
	try {
		OSF.ClientHostController.closeDialog(args.dispId, args.targetId, callback);
	}
	catch (ex) {
		OSF.DDA.SafeArray.Delegate._onException(ex, args);
	}
};
OSF.DDA.SafeArray.Delegate.messageParent=function OSF_DDA_SafeArray_Delegate$MessageParent(args) {
	try {
		if (args.onCalling) {
			args.onCalling();
		}
		var startTime=(new Date()).getTime();
		var result=OSF.ClientHostController.messageParent(args.hostCallArgs);
		if (args.onReceiving) {
			args.onReceiving();
		}
		if (OSF.AppTelemetry) {
			OSF.AppTelemetry.onMethodDone(args.dispId, args.hostCallArgs, Math.abs((new Date()).getTime() - startTime), result);
		}
		return result;
	}
	catch (ex) {
		return OSF.DDA.SafeArray.Delegate._onExceptionSyncMethod(ex);
	}
};
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.EventDispId.dispidDialogMessageReceivedEvent,
	fromHost: [
		{ name: OSF.DDA.EventDescriptors.DialogMessageReceivedEvent, value: OSF.DDA.SafeArray.Delegate.ParameterMap.self }
	],
	isComplexType: true
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.EventDescriptors.DialogMessageReceivedEvent,
	fromHost: [
		{ name: OSF.DDA.PropertyDescriptors.MessageType, value: 0 },
		{ name: OSF.DDA.PropertyDescriptors.MessageContent, value: 1 }
	],
	isComplexType: true
});
OSF.DDA.SafeArray.Delegate.sendMessage=function OSF_DDA_SafeArray_Delegate$SendMessage(args) {
	try {
		if (args.onCalling) {
			args.onCalling();
		}
		var startTime=(new Date()).getTime();
		var result=OSF.ClientHostController.sendMessage(args.hostCallArgs);
		if (args.onReceiving) {
			args.onReceiving();
		}
		return result;
	}
	catch (ex) {
		return OSF.DDA.SafeArray.Delegate._onExceptionSyncMethod(ex);
	}
};
Microsoft.Office.WebExtension.TableData=function Microsoft_Office_WebExtension_TableData(rows, headers) {
	function fixData(data) {
		if (data==null || data==undefined) {
			return null;
		}
		try {
			for (var dim=OSF.DDA.DataCoercion.findArrayDimensionality(data, 2); dim < 2; dim++) {
				data=[data];
			}
			return data;
		}
		catch (ex) {
		}
	}
	;
	OSF.OUtil.defineEnumerableProperties(this, {
		"headers": {
			get: function () { return headers; },
			set: function (value) {
				headers=fixData(value);
			}
		},
		"rows": {
			get: function () { return rows; },
			set: function (value) {
				rows=(value==null || (OSF.OUtil.isArray(value) && (value.length==0))) ?
					[] :
					fixData(value);
			}
		}
	});
	this.headers=headers;
	this.rows=rows;
};
OSF.DDA.OMFactory=OSF.DDA.OMFactory || {};
OSF.DDA.OMFactory.manufactureTableData=function OSF_DDA_OMFactory$manufactureTableData(tableDataProperties) {
	return new Microsoft.Office.WebExtension.TableData(tableDataProperties[OSF.DDA.TableDataProperties.TableRows], tableDataProperties[OSF.DDA.TableDataProperties.TableHeaders]);
};
Microsoft.Office.WebExtension.CoercionType={
	Text: "text",
	Matrix: "matrix",
	Table: "table"
};
OSF.DDA.DataCoercion=(function OSF_DDA_DataCoercion() {
	return {
		findArrayDimensionality: function OSF_DDA_DataCoercion$findArrayDimensionality(obj) {
			if (OSF.OUtil.isArray(obj)) {
				var dim=0;
				for (var index=0; index < obj.length; index++) {
					dim=Math.max(dim, OSF.DDA.DataCoercion.findArrayDimensionality(obj[index]));
				}
				return dim+1;
			}
			else {
				return 0;
			}
		},
		getCoercionDefaultForBinding: function OSF_DDA_DataCoercion$getCoercionDefaultForBinding(bindingType) {
			switch (bindingType) {
				case Microsoft.Office.WebExtension.BindingType.Matrix: return Microsoft.Office.WebExtension.CoercionType.Matrix;
				case Microsoft.Office.WebExtension.BindingType.Table: return Microsoft.Office.WebExtension.CoercionType.Table;
				case Microsoft.Office.WebExtension.BindingType.Text:
				default:
					return Microsoft.Office.WebExtension.CoercionType.Text;
			}
		},
		getBindingDefaultForCoercion: function OSF_DDA_DataCoercion$getBindingDefaultForCoercion(coercionType) {
			switch (coercionType) {
				case Microsoft.Office.WebExtension.CoercionType.Matrix: return Microsoft.Office.WebExtension.BindingType.Matrix;
				case Microsoft.Office.WebExtension.CoercionType.Table: return Microsoft.Office.WebExtension.BindingType.Table;
				case Microsoft.Office.WebExtension.CoercionType.Text:
				case Microsoft.Office.WebExtension.CoercionType.Html:
				case Microsoft.Office.WebExtension.CoercionType.Ooxml:
				default:
					return Microsoft.Office.WebExtension.BindingType.Text;
			}
		},
		determineCoercionType: function OSF_DDA_DataCoercion$determineCoercionType(data) {
			if (data==null || data==undefined)
				return null;
			var sourceType=null;
			var runtimeType=typeof data;
			if (data.rows !==undefined) {
				sourceType=Microsoft.Office.WebExtension.CoercionType.Table;
			}
			else if (OSF.OUtil.isArray(data)) {
				sourceType=Microsoft.Office.WebExtension.CoercionType.Matrix;
			}
			else if (runtimeType=="string" || runtimeType=="number" || runtimeType=="boolean" || OSF.OUtil.isDate(data)) {
				sourceType=Microsoft.Office.WebExtension.CoercionType.Text;
			}
			else {
				throw OSF.DDA.ErrorCodeManager.errorCodes.ooeUnsupportedDataObject;
			}
			return sourceType;
		},
		coerceData: function OSF_DDA_DataCoercion$coerceData(data, destinationType, sourceType) {
			sourceType=sourceType || OSF.DDA.DataCoercion.determineCoercionType(data);
			if (sourceType && sourceType !=destinationType) {
				OSF.OUtil.writeProfilerMark(OSF.InternalPerfMarker.DataCoercionBegin);
				data=OSF.DDA.DataCoercion._coerceDataFromTable(destinationType, OSF.DDA.DataCoercion._coerceDataToTable(data, sourceType));
				OSF.OUtil.writeProfilerMark(OSF.InternalPerfMarker.DataCoercionEnd);
			}
			return data;
		},
		_matrixToText: function OSF_DDA_DataCoercion$_matrixToText(matrix) {
			if (matrix.length==1 && matrix[0].length==1)
				return ""+matrix[0][0];
			var val="";
			for (var i=0; i < matrix.length; i++) {
				val+=matrix[i].join("\t")+"\n";
			}
			return val.substring(0, val.length - 1);
		},
		_textToMatrix: function OSF_DDA_DataCoercion$_textToMatrix(text) {
			var ret=text.split("\n");
			for (var i=0; i < ret.length; i++)
				ret[i]=ret[i].split("\t");
			return ret;
		},
		_tableToText: function OSF_DDA_DataCoercion$_tableToText(table) {
			var headers="";
			if (table.headers !=null) {
				headers=OSF.DDA.DataCoercion._matrixToText([table.headers])+"\n";
			}
			var rows=OSF.DDA.DataCoercion._matrixToText(table.rows);
			if (rows=="") {
				headers=headers.substring(0, headers.length - 1);
			}
			return headers+rows;
		},
		_tableToMatrix: function OSF_DDA_DataCoercion$_tableToMatrix(table) {
			var matrix=table.rows;
			if (table.headers !=null) {
				matrix.unshift(table.headers);
			}
			return matrix;
		},
		_coerceDataFromTable: function OSF_DDA_DataCoercion$_coerceDataFromTable(coercionType, table) {
			var value;
			switch (coercionType) {
				case Microsoft.Office.WebExtension.CoercionType.Table:
					value=table;
					break;
				case Microsoft.Office.WebExtension.CoercionType.Matrix:
					value=OSF.DDA.DataCoercion._tableToMatrix(table);
					break;
				case Microsoft.Office.WebExtension.CoercionType.SlideRange:
					value=null;
					if (OSF.DDA.OMFactory.manufactureSlideRange) {
						value=OSF.DDA.OMFactory.manufactureSlideRange(OSF.DDA.DataCoercion._tableToText(table));
					}
					if (value==null) {
						value=OSF.DDA.DataCoercion._tableToText(table);
					}
					break;
				case Microsoft.Office.WebExtension.CoercionType.Text:
				case Microsoft.Office.WebExtension.CoercionType.Html:
				case Microsoft.Office.WebExtension.CoercionType.Ooxml:
				default:
					value=OSF.DDA.DataCoercion._tableToText(table);
					break;
			}
			return value;
		},
		_coerceDataToTable: function OSF_DDA_DataCoercion$_coerceDataToTable(data, sourceType) {
			if (sourceType==undefined) {
				sourceType=OSF.DDA.DataCoercion.determineCoercionType(data);
			}
			var value;
			switch (sourceType) {
				case Microsoft.Office.WebExtension.CoercionType.Table:
					value=data;
					break;
				case Microsoft.Office.WebExtension.CoercionType.Matrix:
					value=new Microsoft.Office.WebExtension.TableData(data);
					break;
				case Microsoft.Office.WebExtension.CoercionType.Text:
				case Microsoft.Office.WebExtension.CoercionType.Html:
				case Microsoft.Office.WebExtension.CoercionType.Ooxml:
				default:
					value=new Microsoft.Office.WebExtension.TableData(OSF.DDA.DataCoercion._textToMatrix(data));
					break;
			}
			return value;
		}
	};
})();
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: Microsoft.Office.WebExtension.Parameters.CoercionType,
	toHost: [
		{ name: Microsoft.Office.WebExtension.CoercionType.Text, value: 0 },
		{ name: Microsoft.Office.WebExtension.CoercionType.Matrix, value: 1 },
		{ name: Microsoft.Office.WebExtension.CoercionType.Table, value: 2 }
	]
});
OSF.DDA.AsyncMethodNames.addNames({
	GetSelectedDataAsync: "getSelectedDataAsync",
	SetSelectedDataAsync: "setSelectedDataAsync"
});
(function () {
	function processData(dataDescriptor, caller, callArgs) {
		var data=dataDescriptor[Microsoft.Office.WebExtension.Parameters.Data];
		if (OSF.DDA.TableDataProperties && data && (data[OSF.DDA.TableDataProperties.TableRows] !=undefined || data[OSF.DDA.TableDataProperties.TableHeaders] !=undefined)) {
			data=OSF.DDA.OMFactory.manufactureTableData(data);
		}
		data=OSF.DDA.DataCoercion.coerceData(data, callArgs[Microsoft.Office.WebExtension.Parameters.CoercionType]);
		return data==undefined ? null : data;
	}
	OSF.DDA.AsyncMethodCalls.define({
		method: OSF.DDA.AsyncMethodNames.GetSelectedDataAsync,
		requiredArguments: [
			{
				"name": Microsoft.Office.WebExtension.Parameters.CoercionType,
				"enum": Microsoft.Office.WebExtension.CoercionType
			}
		],
		supportedOptions: [
			{
				name: Microsoft.Office.WebExtension.Parameters.ValueFormat,
				value: {
					"enum": Microsoft.Office.WebExtension.ValueFormat,
					"defaultValue": Microsoft.Office.WebExtension.ValueFormat.Unformatted
				}
			},
			{
				name: Microsoft.Office.WebExtension.Parameters.FilterType,
				value: {
					"enum": Microsoft.Office.WebExtension.FilterType,
					"defaultValue": Microsoft.Office.WebExtension.FilterType.All
				}
			}
		],
		privateStateCallbacks: [],
		onSucceeded: processData
	});
	OSF.DDA.AsyncMethodCalls.define({
		method: OSF.DDA.AsyncMethodNames.SetSelectedDataAsync,
		requiredArguments: [
			{
				"name": Microsoft.Office.WebExtension.Parameters.Data,
				"types": ["string", "object", "number", "boolean"]
			}
		],
		supportedOptions: [
			{
				name: Microsoft.Office.WebExtension.Parameters.CoercionType,
				value: {
					"enum": Microsoft.Office.WebExtension.CoercionType,
					"calculate": function (requiredArgs) {
						return OSF.DDA.DataCoercion.determineCoercionType(requiredArgs[Microsoft.Office.WebExtension.Parameters.Data]);
					}
				}
			},
			{
				name: Microsoft.Office.WebExtension.Parameters.ImageLeft,
				value: {
					"types": ["number", "boolean"],
					"defaultValue": false
				}
			},
			{
				name: Microsoft.Office.WebExtension.Parameters.ImageTop,
				value: {
					"types": ["number", "boolean"],
					"defaultValue": false
				}
			},
			{
				name: Microsoft.Office.WebExtension.Parameters.ImageWidth,
				value: {
					"types": ["number", "boolean"],
					"defaultValue": false
				}
			},
			{
				name: Microsoft.Office.WebExtension.Parameters.ImageHeight,
				value: {
					"types": ["number", "boolean"],
					"defaultValue": false
				}
			}
		],
		privateStateCallbacks: []
	});
})();
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.MethodDispId.dispidGetSelectedDataMethod,
	fromHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.Data, value: OSF.DDA.SafeArray.Delegate.ParameterMap.self }
	],
	toHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.CoercionType, value: 0 },
		{ name: Microsoft.Office.WebExtension.Parameters.ValueFormat, value: 1 },
		{ name: Microsoft.Office.WebExtension.Parameters.FilterType, value: 2 }
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.MethodDispId.dispidSetSelectedDataMethod,
	toHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.CoercionType, value: 0 },
		{ name: Microsoft.Office.WebExtension.Parameters.Data, value: 1 },
		{ name: Microsoft.Office.WebExtension.Parameters.ImageLeft, value: 2 },
		{ name: Microsoft.Office.WebExtension.Parameters.ImageTop, value: 3 },
		{ name: Microsoft.Office.WebExtension.Parameters.ImageWidth, value: 4 },
		{ name: Microsoft.Office.WebExtension.Parameters.ImageHeight, value: 5 },
	]
});
OSF.DDA.SettingsManager={
	SerializedSettings: "serializedSettings",
	RefreshingSettings: "refreshingSettings",
	DateJSONPrefix: "Date(",
	DataJSONSuffix: ")",
	serializeSettings: function OSF_DDA_SettingsManager$serializeSettings(settingsCollection) {
		return OSF.OUtil.serializeSettings(settingsCollection);
	},
	deserializeSettings: function OSF_DDA_SettingsManager$deserializeSettings(serializedSettings) {
		return OSF.OUtil.deserializeSettings(serializedSettings);
	}
};
OSF.DDA.Settings=function OSF_DDA_Settings(settings) {
	settings=settings || {};
	var cacheSessionSettings=function (settings) {
		var osfSessionStorage=OSF.OUtil.getSessionStorage();
		if (osfSessionStorage) {
			var serializedSettings=OSF.DDA.SettingsManager.serializeSettings(settings);
			var storageSettings=JSON ? JSON.stringify(serializedSettings) : Sys.Serialization.JavaScriptSerializer.serialize(serializedSettings);
			osfSessionStorage.setItem(OSF._OfficeAppFactory.getCachedSessionSettingsKey(), storageSettings);
		}
	};
	OSF.OUtil.defineEnumerableProperties(this, {
		"get": {
			value: function OSF_DDA_Settings$get(name) {
				var e=Function._validateParams(arguments, [
					{ name: "name", type: String, mayBeNull: false }
				]);
				if (e)
					throw e;
				var setting=settings[name];
				return typeof (setting)==='undefined' ? null : setting;
			}
		},
		"set": {
			value: function OSF_DDA_Settings$set(name, value) {
				var e=Function._validateParams(arguments, [
					{ name: "name", type: String, mayBeNull: false },
					{ name: "value", mayBeNull: true }
				]);
				if (e)
					throw e;
				settings[name]=value;
				cacheSessionSettings(settings);
			}
		},
		"remove": {
			value: function OSF_DDA_Settings$remove(name) {
				var e=Function._validateParams(arguments, [
					{ name: "name", type: String, mayBeNull: false }
				]);
				if (e)
					throw e;
				delete settings[name];
				cacheSessionSettings(settings);
			}
		}
	});
	OSF.DDA.DispIdHost.addAsyncMethods(this, [OSF.DDA.AsyncMethodNames.SaveAsync], settings);
};
OSF.DDA.RefreshableSettings=function OSF_DDA_RefreshableSettings(settings) {
	OSF.DDA.RefreshableSettings.uber.constructor.call(this, settings);
	OSF.DDA.DispIdHost.addAsyncMethods(this, [OSF.DDA.AsyncMethodNames.RefreshAsync], settings);
	OSF.DDA.DispIdHost.addEventSupport(this, new OSF.EventDispatch([Microsoft.Office.WebExtension.EventType.SettingsChanged]));
};
OSF.OUtil.extend(OSF.DDA.RefreshableSettings, OSF.DDA.Settings);
OSF.OUtil.augmentList(Microsoft.Office.WebExtension.EventType, {
	SettingsChanged: "settingsChanged"
});
OSF.DDA.SettingsChangedEventArgs=function OSF_DDA_SettingsChangedEventArgs(settingsInstance) {
	OSF.OUtil.defineEnumerableProperties(this, {
		"type": {
			value: Microsoft.Office.WebExtension.EventType.SettingsChanged
		},
		"settings": {
			value: settingsInstance
		}
	});
};
OSF.DDA.AsyncMethodNames.addNames({
	RefreshAsync: "refreshAsync",
	SaveAsync: "saveAsync"
});
OSF.DDA.AsyncMethodCalls.define({
	method: OSF.DDA.AsyncMethodNames.RefreshAsync,
	requiredArguments: [],
	supportedOptions: [],
	privateStateCallbacks: [
		{
			name: OSF.DDA.SettingsManager.RefreshingSettings,
			value: function getRefreshingSettings(settingsInstance, settingsCollection) {
				return settingsCollection;
			}
		}
	],
	onSucceeded: function deserializeSettings(serializedSettingsDescriptor, refreshingSettings, refreshingSettingsArgs) {
		var serializedSettings=serializedSettingsDescriptor[OSF.DDA.SettingsManager.SerializedSettings];
		var newSettings=OSF.DDA.SettingsManager.deserializeSettings(serializedSettings);
		var oldSettings=refreshingSettingsArgs[OSF.DDA.SettingsManager.RefreshingSettings];
		for (var setting in oldSettings) {
			refreshingSettings.remove(setting);
		}
		for (var setting in newSettings) {
			refreshingSettings.set(setting, newSettings[setting]);
		}
		return refreshingSettings;
	}
});
OSF.DDA.AsyncMethodCalls.define({
	method: OSF.DDA.AsyncMethodNames.SaveAsync,
	requiredArguments: [],
	supportedOptions: [
		{
			name: Microsoft.Office.WebExtension.Parameters.OverwriteIfStale,
			value: {
				"types": ["boolean"],
				"defaultValue": true
			}
		}
	],
	privateStateCallbacks: [
		{
			name: OSF.DDA.SettingsManager.SerializedSettings,
			value: function serializeSettings(settingsInstance, settingsCollection) {
				return OSF.DDA.SettingsManager.serializeSettings(settingsCollection);
			}
		}
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.MethodDispId.dispidLoadSettingsMethod,
	fromHost: [
		{ name: OSF.DDA.SettingsManager.SerializedSettings, value: OSF.DDA.SafeArray.Delegate.ParameterMap.self }
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.MethodDispId.dispidSaveSettingsMethod,
	toHost: [
		{ name: OSF.DDA.SettingsManager.SerializedSettings, value: OSF.DDA.SettingsManager.SerializedSettings },
		{ name: Microsoft.Office.WebExtension.Parameters.OverwriteIfStale, value: Microsoft.Office.WebExtension.Parameters.OverwriteIfStale }
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({ type: OSF.DDA.EventDispId.dispidSettingsChangedEvent });
Microsoft.Office.WebExtension.BindingType={
	Table: "table",
	Text: "text",
	Matrix: "matrix"
};
OSF.DDA.BindingProperties={
	Id: "BindingId",
	Type: Microsoft.Office.WebExtension.Parameters.BindingType
};
OSF.OUtil.augmentList(OSF.DDA.ListDescriptors, { BindingList: "BindingList" });
OSF.OUtil.augmentList(OSF.DDA.PropertyDescriptors, {
	Subset: "subset",
	BindingProperties: "BindingProperties"
});
OSF.DDA.ListType.setListType(OSF.DDA.ListDescriptors.BindingList, OSF.DDA.PropertyDescriptors.BindingProperties);
OSF.DDA.BindingPromise=function OSF_DDA_BindingPromise(bindingId, errorCallback) {
	this._id=bindingId;
	OSF.OUtil.defineEnumerableProperty(this, "onFail", {
		get: function () {
			return errorCallback;
		},
		set: function (onError) {
			var t=typeof onError;
			if (t !="undefined" && t !="function") {
				throw OSF.OUtil.formatString(Strings.OfficeOM.L_CallbackNotAFunction, t);
			}
			errorCallback=onError;
		}
	});
};
OSF.DDA.BindingPromise.prototype={
	_fetch: function OSF_DDA_BindingPromise$_fetch(onComplete) {
		if (this.binding) {
			if (onComplete)
				onComplete(this.binding);
		}
		else {
			if (!this._binding) {
				var me=this;
				Microsoft.Office.WebExtension.context.document.bindings.getByIdAsync(this._id, function (asyncResult) {
					if (asyncResult.status==Microsoft.Office.WebExtension.AsyncResultStatus.Succeeded) {
						OSF.OUtil.defineEnumerableProperty(me, "binding", {
							value: asyncResult.value
						});
						if (onComplete)
							onComplete(me.binding);
					}
					else {
						if (me.onFail)
							me.onFail(asyncResult);
					}
				});
			}
		}
		return this;
	},
	getDataAsync: function OSF_DDA_BindingPromise$getDataAsync() {
		var args=arguments;
		this._fetch(function onComplete(binding) { binding.getDataAsync.apply(binding, args); });
		return this;
	},
	setDataAsync: function OSF_DDA_BindingPromise$setDataAsync() {
		var args=arguments;
		this._fetch(function onComplete(binding) { binding.setDataAsync.apply(binding, args); });
		return this;
	},
	addHandlerAsync: function OSF_DDA_BindingPromise$addHandlerAsync() {
		var args=arguments;
		this._fetch(function onComplete(binding) { binding.addHandlerAsync.apply(binding, args); });
		return this;
	},
	removeHandlerAsync: function OSF_DDA_BindingPromise$removeHandlerAsync() {
		var args=arguments;
		this._fetch(function onComplete(binding) { binding.removeHandlerAsync.apply(binding, args); });
		return this;
	}
};
OSF.DDA.BindingFacade=function OSF_DDA_BindingFacade(docInstance) {
	this._eventDispatches=[];
	OSF.OUtil.defineEnumerableProperty(this, "document", {
		value: docInstance
	});
	var am=OSF.DDA.AsyncMethodNames;
	OSF.DDA.DispIdHost.addAsyncMethods(this, [
		am.AddFromSelectionAsync,
		am.AddFromNamedItemAsync,
		am.GetAllAsync,
		am.GetByIdAsync,
		am.ReleaseByIdAsync
	]);
};
OSF.DDA.UnknownBinding=function OSF_DDA_UknonwnBinding(id, docInstance) {
	OSF.OUtil.defineEnumerableProperties(this, {
		"document": { value: docInstance },
		"id": { value: id }
	});
};
OSF.DDA.Binding=function OSF_DDA_Binding(id, docInstance) {
	OSF.OUtil.defineEnumerableProperties(this, {
		"document": {
			value: docInstance
		},
		"id": {
			value: id
		}
	});
	var am=OSF.DDA.AsyncMethodNames;
	OSF.DDA.DispIdHost.addAsyncMethods(this, [
		am.GetDataAsync,
		am.SetDataAsync
	]);
	var et=Microsoft.Office.WebExtension.EventType;
	var bindingEventDispatches=docInstance.bindings._eventDispatches;
	if (!bindingEventDispatches[id]) {
		bindingEventDispatches[id]=new OSF.EventDispatch([
			et.BindingSelectionChanged,
			et.BindingDataChanged
		]);
	}
	var eventDispatch=bindingEventDispatches[id];
	OSF.DDA.DispIdHost.addEventSupport(this, eventDispatch);
};
OSF.DDA.generateBindingId=function OSF_DDA$GenerateBindingId() {
	return "UnnamedBinding_"+OSF.OUtil.getUniqueId()+"_"+new Date().getTime();
};
OSF.DDA.OMFactory=OSF.DDA.OMFactory || {};
OSF.DDA.OMFactory.manufactureBinding=function OSF_DDA_OMFactory$manufactureBinding(bindingProperties, containingDocument) {
	var id=bindingProperties[OSF.DDA.BindingProperties.Id];
	var rows=bindingProperties[OSF.DDA.BindingProperties.RowCount];
	var cols=bindingProperties[OSF.DDA.BindingProperties.ColumnCount];
	var hasHeaders=bindingProperties[OSF.DDA.BindingProperties.HasHeaders];
	var binding;
	switch (bindingProperties[OSF.DDA.BindingProperties.Type]) {
		case Microsoft.Office.WebExtension.BindingType.Text:
			binding=new OSF.DDA.TextBinding(id, containingDocument);
			break;
		case Microsoft.Office.WebExtension.BindingType.Matrix:
			binding=new OSF.DDA.MatrixBinding(id, containingDocument, rows, cols);
			break;
		case Microsoft.Office.WebExtension.BindingType.Table:
			var isExcelApp=function () {
				return (OSF.DDA.ExcelDocument)
					&& (Microsoft.Office.WebExtension.context.document)
					&& (Microsoft.Office.WebExtension.context.document instanceof OSF.DDA.ExcelDocument);
			};
			var tableBindingObject;
			if (isExcelApp() && OSF.DDA.ExcelTableBinding) {
				tableBindingObject=OSF.DDA.ExcelTableBinding;
			}
			else {
				tableBindingObject=OSF.DDA.TableBinding;
			}
			binding=new tableBindingObject(id, containingDocument, rows, cols, hasHeaders);
			break;
		default:
			binding=new OSF.DDA.UnknownBinding(id, containingDocument);
	}
	return binding;
};
OSF.DDA.AsyncMethodNames.addNames({
	AddFromSelectionAsync: "addFromSelectionAsync",
	AddFromNamedItemAsync: "addFromNamedItemAsync",
	GetAllAsync: "getAllAsync",
	GetByIdAsync: "getByIdAsync",
	ReleaseByIdAsync: "releaseByIdAsync",
	GetDataAsync: "getDataAsync",
	SetDataAsync: "setDataAsync"
});
(function () {
	function processBinding(bindingDescriptor) {
		return OSF.DDA.OMFactory.manufactureBinding(bindingDescriptor, Microsoft.Office.WebExtension.context.document);
	}
	function getObjectId(obj) { return obj.id; }
	function processData(dataDescriptor, caller, callArgs) {
		var data=dataDescriptor[Microsoft.Office.WebExtension.Parameters.Data];
		if (OSF.DDA.TableDataProperties && data && (data[OSF.DDA.TableDataProperties.TableRows] !=undefined || data[OSF.DDA.TableDataProperties.TableHeaders] !=undefined)) {
			data=OSF.DDA.OMFactory.manufactureTableData(data);
		}
		data=OSF.DDA.DataCoercion.coerceData(data, callArgs[Microsoft.Office.WebExtension.Parameters.CoercionType]);
		return data==undefined ? null : data;
	}
	OSF.DDA.AsyncMethodCalls.define({
		method: OSF.DDA.AsyncMethodNames.AddFromSelectionAsync,
		requiredArguments: [
			{
				"name": Microsoft.Office.WebExtension.Parameters.BindingType,
				"enum": Microsoft.Office.WebExtension.BindingType
			}
		],
		supportedOptions: [{
				name: Microsoft.Office.WebExtension.Parameters.Id,
				value: {
					"types": ["string"],
					"calculate": OSF.DDA.generateBindingId
				}
			},
			{
				name: Microsoft.Office.WebExtension.Parameters.Columns,
				value: {
					"types": ["object"],
					"defaultValue": null
				}
			}
		],
		privateStateCallbacks: [],
		onSucceeded: processBinding
	});
	OSF.DDA.AsyncMethodCalls.define({
		method: OSF.DDA.AsyncMethodNames.AddFromNamedItemAsync,
		requiredArguments: [{
				"name": Microsoft.Office.WebExtension.Parameters.ItemName,
				"types": ["string"]
			},
			{
				"name": Microsoft.Office.WebExtension.Parameters.BindingType,
				"enum": Microsoft.Office.WebExtension.BindingType
			}
		],
		supportedOptions: [{
				name: Microsoft.Office.WebExtension.Parameters.Id,
				value: {
					"types": ["string"],
					"calculate": OSF.DDA.generateBindingId
				}
			},
			{
				name: Microsoft.Office.WebExtension.Parameters.Columns,
				value: {
					"types": ["object"],
					"defaultValue": null
				}
			}
		],
		privateStateCallbacks: [
			{
				name: Microsoft.Office.WebExtension.Parameters.FailOnCollision,
				value: function () { return true; }
			}
		],
		onSucceeded: processBinding
	});
	OSF.DDA.AsyncMethodCalls.define({
		method: OSF.DDA.AsyncMethodNames.GetAllAsync,
		requiredArguments: [],
		supportedOptions: [],
		privateStateCallbacks: [],
		onSucceeded: function (response) { return OSF.OUtil.mapList(response[OSF.DDA.ListDescriptors.BindingList], processBinding); }
	});
	OSF.DDA.AsyncMethodCalls.define({
		method: OSF.DDA.AsyncMethodNames.GetByIdAsync,
		requiredArguments: [
			{
				"name": Microsoft.Office.WebExtension.Parameters.Id,
				"types": ["string"]
			}
		],
		supportedOptions: [],
		privateStateCallbacks: [],
		onSucceeded: processBinding
	});
	OSF.DDA.AsyncMethodCalls.define({
		method: OSF.DDA.AsyncMethodNames.ReleaseByIdAsync,
		requiredArguments: [
			{
				"name": Microsoft.Office.WebExtension.Parameters.Id,
				"types": ["string"]
			}
		],
		supportedOptions: [],
		privateStateCallbacks: [],
		onSucceeded: function (response, caller, callArgs) {
			var id=callArgs[Microsoft.Office.WebExtension.Parameters.Id];
			delete caller._eventDispatches[id];
		}
	});
	OSF.DDA.AsyncMethodCalls.define({
		method: OSF.DDA.AsyncMethodNames.GetDataAsync,
		requiredArguments: [],
		supportedOptions: [{
				name: Microsoft.Office.WebExtension.Parameters.CoercionType,
				value: {
					"enum": Microsoft.Office.WebExtension.CoercionType,
					"calculate": function (requiredArgs, binding) { return OSF.DDA.DataCoercion.getCoercionDefaultForBinding(binding.type); }
				}
			},
			{
				name: Microsoft.Office.WebExtension.Parameters.ValueFormat,
				value: {
					"enum": Microsoft.Office.WebExtension.ValueFormat,
					"defaultValue": Microsoft.Office.WebExtension.ValueFormat.Unformatted
				}
			},
			{
				name: Microsoft.Office.WebExtension.Parameters.FilterType,
				value: {
					"enum": Microsoft.Office.WebExtension.FilterType,
					"defaultValue": Microsoft.Office.WebExtension.FilterType.All
				}
			},
			{
				name: Microsoft.Office.WebExtension.Parameters.Rows,
				value: {
					"types": ["object", "string"],
					"defaultValue": null
				}
			},
			{
				name: Microsoft.Office.WebExtension.Parameters.Columns,
				value: {
					"types": ["object"],
					"defaultValue": null
				}
			},
			{
				name: Microsoft.Office.WebExtension.Parameters.StartRow,
				value: {
					"types": ["number"],
					"defaultValue": 0
				}
			},
			{
				name: Microsoft.Office.WebExtension.Parameters.StartColumn,
				value: {
					"types": ["number"],
					"defaultValue": 0
				}
			},
			{
				name: Microsoft.Office.WebExtension.Parameters.RowCount,
				value: {
					"types": ["number"],
					"defaultValue": 0
				}
			},
			{
				name: Microsoft.Office.WebExtension.Parameters.ColumnCount,
				value: {
					"types": ["number"],
					"defaultValue": 0
				}
			}
		],
		checkCallArgs: function (callArgs, caller, stateInfo) {
			if (callArgs[Microsoft.Office.WebExtension.Parameters.StartRow]==0 &&
				callArgs[Microsoft.Office.WebExtension.Parameters.StartColumn]==0 &&
				callArgs[Microsoft.Office.WebExtension.Parameters.RowCount]==0 &&
				callArgs[Microsoft.Office.WebExtension.Parameters.ColumnCount]==0) {
				delete callArgs[Microsoft.Office.WebExtension.Parameters.StartRow];
				delete callArgs[Microsoft.Office.WebExtension.Parameters.StartColumn];
				delete callArgs[Microsoft.Office.WebExtension.Parameters.RowCount];
				delete callArgs[Microsoft.Office.WebExtension.Parameters.ColumnCount];
			}
			if (callArgs[Microsoft.Office.WebExtension.Parameters.CoercionType] !=OSF.DDA.DataCoercion.getCoercionDefaultForBinding(caller.type) &&
				(callArgs[Microsoft.Office.WebExtension.Parameters.StartRow] ||
					callArgs[Microsoft.Office.WebExtension.Parameters.StartColumn] ||
					callArgs[Microsoft.Office.WebExtension.Parameters.RowCount] ||
					callArgs[Microsoft.Office.WebExtension.Parameters.ColumnCount])) {
				throw OSF.DDA.ErrorCodeManager.errorCodes.ooeCoercionTypeNotMatchBinding;
			}
			return callArgs;
		},
		privateStateCallbacks: [
			{
				name: Microsoft.Office.WebExtension.Parameters.Id,
				value: getObjectId
			}
		],
		onSucceeded: processData
	});
	OSF.DDA.AsyncMethodCalls.define({
		method: OSF.DDA.AsyncMethodNames.SetDataAsync,
		requiredArguments: [
			{
				"name": Microsoft.Office.WebExtension.Parameters.Data,
				"types": ["string", "object", "number", "boolean"]
			}
		],
		supportedOptions: [{
				name: Microsoft.Office.WebExtension.Parameters.CoercionType,
				value: {
					"enum": Microsoft.Office.WebExtension.CoercionType,
					"calculate": function (requiredArgs) { return OSF.DDA.DataCoercion.determineCoercionType(requiredArgs[Microsoft.Office.WebExtension.Parameters.Data]); }
				}
			},
			{
				name: Microsoft.Office.WebExtension.Parameters.Rows,
				value: {
					"types": ["object", "string"],
					"defaultValue": null
				}
			},
			{
				name: Microsoft.Office.WebExtension.Parameters.Columns,
				value: {
					"types": ["object"],
					"defaultValue": null
				}
			},
			{
				name: Microsoft.Office.WebExtension.Parameters.StartRow,
				value: {
					"types": ["number"],
					"defaultValue": 0
				}
			},
			{
				name: Microsoft.Office.WebExtension.Parameters.StartColumn,
				value: {
					"types": ["number"],
					"defaultValue": 0
				}
			}
		],
		checkCallArgs: function (callArgs, caller, stateInfo) {
			if (callArgs[Microsoft.Office.WebExtension.Parameters.StartRow]==0 &&
				callArgs[Microsoft.Office.WebExtension.Parameters.StartColumn]==0) {
				delete callArgs[Microsoft.Office.WebExtension.Parameters.StartRow];
				delete callArgs[Microsoft.Office.WebExtension.Parameters.StartColumn];
			}
			if (callArgs[Microsoft.Office.WebExtension.Parameters.CoercionType] !=OSF.DDA.DataCoercion.getCoercionDefaultForBinding(caller.type) &&
				(callArgs[Microsoft.Office.WebExtension.Parameters.StartRow] ||
					callArgs[Microsoft.Office.WebExtension.Parameters.StartColumn])) {
				throw OSF.DDA.ErrorCodeManager.errorCodes.ooeCoercionTypeNotMatchBinding;
			}
			return callArgs;
		},
		privateStateCallbacks: [
			{
				name: Microsoft.Office.WebExtension.Parameters.Id,
				value: getObjectId
			}
		]
	});
})();
OSF.OUtil.augmentList(OSF.DDA.BindingProperties, {
	RowCount: "BindingRowCount",
	ColumnCount: "BindingColumnCount",
	HasHeaders: "HasHeaders"
});
OSF.DDA.MatrixBinding=function OSF_DDA_MatrixBinding(id, docInstance, rows, cols) {
	OSF.DDA.MatrixBinding.uber.constructor.call(this, id, docInstance);
	OSF.OUtil.defineEnumerableProperties(this, {
		"type": {
			value: Microsoft.Office.WebExtension.BindingType.Matrix
		},
		"rowCount": {
			value: rows ? rows : 0
		},
		"columnCount": {
			value: cols ? cols : 0
		}
	});
};
OSF.OUtil.extend(OSF.DDA.MatrixBinding, OSF.DDA.Binding);
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.PropertyDescriptors.BindingProperties,
	fromHost: [
		{ name: OSF.DDA.BindingProperties.Id, value: 0 },
		{ name: OSF.DDA.BindingProperties.Type, value: 1 },
		{ name: OSF.DDA.SafeArray.UniqueArguments.BindingSpecificData, value: 2 }
	],
	isComplexType: true
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: Microsoft.Office.WebExtension.Parameters.BindingType,
	toHost: [
		{ name: Microsoft.Office.WebExtension.BindingType.Text, value: 0 },
		{ name: Microsoft.Office.WebExtension.BindingType.Matrix, value: 1 },
		{ name: Microsoft.Office.WebExtension.BindingType.Table, value: 2 }
	],
	invertible: true
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.MethodDispId.dispidAddBindingFromSelectionMethod,
	fromHost: [
		{ name: OSF.DDA.PropertyDescriptors.BindingProperties, value: OSF.DDA.SafeArray.Delegate.ParameterMap.self }
	],
	toHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.Id, value: 0 },
		{ name: Microsoft.Office.WebExtension.Parameters.BindingType, value: 1 }
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.MethodDispId.dispidAddBindingFromNamedItemMethod,
	fromHost: [
		{ name: OSF.DDA.PropertyDescriptors.BindingProperties, value: OSF.DDA.SafeArray.Delegate.ParameterMap.self }
	],
	toHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.ItemName, value: 0 },
		{ name: Microsoft.Office.WebExtension.Parameters.Id, value: 1 },
		{ name: Microsoft.Office.WebExtension.Parameters.BindingType, value: 2 },
		{ name: Microsoft.Office.WebExtension.Parameters.FailOnCollision, value: 3 }
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.MethodDispId.dispidReleaseBindingMethod,
	toHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.Id, value: 0 }
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.MethodDispId.dispidGetBindingMethod,
	fromHost: [
		{ name: OSF.DDA.PropertyDescriptors.BindingProperties, value: OSF.DDA.SafeArray.Delegate.ParameterMap.self }
	],
	toHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.Id, value: 0 }
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.MethodDispId.dispidGetAllBindingsMethod,
	fromHost: [
		{ name: OSF.DDA.ListDescriptors.BindingList, value: OSF.DDA.SafeArray.Delegate.ParameterMap.self }
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.MethodDispId.dispidGetBindingDataMethod,
	fromHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.Data, value: OSF.DDA.SafeArray.Delegate.ParameterMap.self }
	],
	toHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.Id, value: 0 },
		{ name: Microsoft.Office.WebExtension.Parameters.CoercionType, value: 1 },
		{ name: Microsoft.Office.WebExtension.Parameters.ValueFormat, value: 2 },
		{ name: Microsoft.Office.WebExtension.Parameters.FilterType, value: 3 },
		{ name: OSF.DDA.PropertyDescriptors.Subset, value: 4 }
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.MethodDispId.dispidSetBindingDataMethod,
	toHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.Id, value: 0 },
		{ name: Microsoft.Office.WebExtension.Parameters.CoercionType, value: 1 },
		{ name: Microsoft.Office.WebExtension.Parameters.Data, value: 2 },
		{ name: OSF.DDA.SafeArray.UniqueArguments.Offset, value: 3 }
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.SafeArray.UniqueArguments.BindingSpecificData,
	fromHost: [
		{ name: OSF.DDA.BindingProperties.RowCount, value: 0 },
		{ name: OSF.DDA.BindingProperties.ColumnCount, value: 1 },
		{ name: OSF.DDA.BindingProperties.HasHeaders, value: 2 }
	],
	isComplexType: true
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.PropertyDescriptors.Subset,
	toHost: [
		{ name: OSF.DDA.SafeArray.UniqueArguments.Offset, value: 0 },
		{ name: OSF.DDA.SafeArray.UniqueArguments.Run, value: 1 }
	],
	canonical: true,
	isComplexType: true
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.SafeArray.UniqueArguments.Offset,
	toHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.StartRow, value: 0 },
		{ name: Microsoft.Office.WebExtension.Parameters.StartColumn, value: 1 }
	],
	canonical: true,
	isComplexType: true
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.SafeArray.UniqueArguments.Run,
	toHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.RowCount, value: 0 },
		{ name: Microsoft.Office.WebExtension.Parameters.ColumnCount, value: 1 }
	],
	canonical: true,
	isComplexType: true
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.MethodDispId.dispidAddRowsMethod,
	toHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.Id, value: 0 },
		{ name: Microsoft.Office.WebExtension.Parameters.Data, value: 1 }
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.MethodDispId.dispidAddColumnsMethod,
	toHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.Id, value: 0 },
		{ name: Microsoft.Office.WebExtension.Parameters.Data, value: 1 }
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.MethodDispId.dispidClearAllRowsMethod,
	toHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.Id, value: 0 }
	]
});
OSF.OUtil.augmentList(OSF.DDA.PropertyDescriptors, { TableDataProperties: "TableDataProperties" });
OSF.OUtil.augmentList(OSF.DDA.BindingProperties, {
	RowCount: "BindingRowCount",
	ColumnCount: "BindingColumnCount",
	HasHeaders: "HasHeaders"
});
OSF.DDA.TableDataProperties={
	TableRows: "TableRows",
	TableHeaders: "TableHeaders"
};
OSF.DDA.TableBinding=function OSF_DDA_TableBinding(id, docInstance, rows, cols, hasHeaders) {
	OSF.DDA.TableBinding.uber.constructor.call(this, id, docInstance);
	OSF.OUtil.defineEnumerableProperties(this, {
		"type": {
			value: Microsoft.Office.WebExtension.BindingType.Table
		},
		"rowCount": {
			value: rows ? rows : 0
		},
		"columnCount": {
			value: cols ? cols : 0
		},
		"hasHeaders": {
			value: hasHeaders ? hasHeaders : false
		}
	});
	var am=OSF.DDA.AsyncMethodNames;
	OSF.DDA.DispIdHost.addAsyncMethods(this, [
		am.AddRowsAsync,
		am.AddColumnsAsync,
		am.DeleteAllDataValuesAsync
	]);
};
OSF.OUtil.extend(OSF.DDA.TableBinding, OSF.DDA.Binding);
OSF.DDA.AsyncMethodNames.addNames({
	AddRowsAsync: "addRowsAsync",
	AddColumnsAsync: "addColumnsAsync",
	DeleteAllDataValuesAsync: "deleteAllDataValuesAsync"
});
(function () {
	function getObjectId(obj) { return obj.id; }
	OSF.DDA.AsyncMethodCalls.define({
		method: OSF.DDA.AsyncMethodNames.AddRowsAsync,
		requiredArguments: [
			{
				"name": Microsoft.Office.WebExtension.Parameters.Data,
				"types": ["object"]
			}
		],
		supportedOptions: [],
		privateStateCallbacks: [
			{
				name: Microsoft.Office.WebExtension.Parameters.Id,
				value: getObjectId
			}
		]
	});
	OSF.DDA.AsyncMethodCalls.define({
		method: OSF.DDA.AsyncMethodNames.AddColumnsAsync,
		requiredArguments: [
			{
				"name": Microsoft.Office.WebExtension.Parameters.Data,
				"types": ["object"]
			}
		],
		supportedOptions: [],
		privateStateCallbacks: [
			{
				name: Microsoft.Office.WebExtension.Parameters.Id,
				value: getObjectId
			}
		]
	});
	OSF.DDA.AsyncMethodCalls.define({
		method: OSF.DDA.AsyncMethodNames.DeleteAllDataValuesAsync,
		requiredArguments: [],
		supportedOptions: [],
		privateStateCallbacks: [
			{
				name: Microsoft.Office.WebExtension.Parameters.Id,
				value: getObjectId
			}
		]
	});
})();
OSF.DDA.TextBinding=function OSF_DDA_TextBinding(id, docInstance) {
	OSF.DDA.TextBinding.uber.constructor.call(this, id, docInstance);
	OSF.OUtil.defineEnumerableProperty(this, "type", {
		value: Microsoft.Office.WebExtension.BindingType.Text
	});
};
OSF.OUtil.extend(OSF.DDA.TextBinding, OSF.DDA.Binding);
OSF.DDA.AsyncMethodNames.addNames({ AddFromPromptAsync: "addFromPromptAsync" });
OSF.DDA.AsyncMethodCalls.define({
	method: OSF.DDA.AsyncMethodNames.AddFromPromptAsync,
	requiredArguments: [
		{
			"name": Microsoft.Office.WebExtension.Parameters.BindingType,
			"enum": Microsoft.Office.WebExtension.BindingType
		}
	],
	supportedOptions: [{
			name: Microsoft.Office.WebExtension.Parameters.Id,
			value: {
				"types": ["string"],
				"calculate": OSF.DDA.generateBindingId
			}
		},
		{
			name: Microsoft.Office.WebExtension.Parameters.PromptText,
			value: {
				"types": ["string"],
				"calculate": function () { return Strings.OfficeOM.L_AddBindingFromPromptDefaultText; }
			}
		},
		{
			name: Microsoft.Office.WebExtension.Parameters.SampleData,
			value: {
				"types": ["object"],
				"defaultValue": null
			}
		}
	],
	privateStateCallbacks: [],
	onSucceeded: function (bindingDescriptor) { return OSF.DDA.OMFactory.manufactureBinding(bindingDescriptor, Microsoft.Office.WebExtension.context.document); }
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.MethodDispId.dispidAddBindingFromPromptMethod,
	fromHost: [
		{ name: OSF.DDA.PropertyDescriptors.BindingProperties, value: OSF.DDA.SafeArray.Delegate.ParameterMap.self }
	],
	toHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.Id, value: 0 },
		{ name: Microsoft.Office.WebExtension.Parameters.BindingType, value: 1 },
		{ name: Microsoft.Office.WebExtension.Parameters.PromptText, value: 2 }
	]
});
OSF.OUtil.augmentList(Microsoft.Office.WebExtension.EventType, { DocumentSelectionChanged: "documentSelectionChanged" });
OSF.DDA.DocumentSelectionChangedEventArgs=function OSF_DDA_DocumentSelectionChangedEventArgs(docInstance) {
	OSF.OUtil.defineEnumerableProperties(this, {
		"type": {
			value: Microsoft.Office.WebExtension.EventType.DocumentSelectionChanged
		},
		"document": {
			value: docInstance
		}
	});
};
OSF.OUtil.augmentList(Microsoft.Office.WebExtension.EventType, { ObjectDeleted: "objectDeleted" });
OSF.OUtil.augmentList(Microsoft.Office.WebExtension.EventType, { ObjectSelectionChanged: "objectSelectionChanged" });
OSF.OUtil.augmentList(Microsoft.Office.WebExtension.EventType, { ObjectDataChanged: "objectDataChanged" });
OSF.OUtil.augmentList(Microsoft.Office.WebExtension.EventType, { ContentControlAdded: "contentControlAdded" });
OSF.DDA.ObjectEventArgs=function OSF_DDA_ObjectEventArgs(eventType, object) {
	OSF.OUtil.defineEnumerableProperties(this, {
		"type": { value: eventType },
		"object": { value: object }
	});
};
OSF.DDA.SafeArray.Delegate.ParameterMap.define({ type: OSF.DDA.EventDispId.dispidDocumentSelectionChangedEvent });
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.EventDispId.dispidObjectDeletedEvent,
	toHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.Id, value: 0 }
	],
	fromHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.Id, value: OSF.DDA.SafeArray.Delegate.ParameterMap.sourceData }
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.EventDispId.dispidObjectSelectionChangedEvent,
	toHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.Id, value: 0 }
	],
	fromHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.Id, value: OSF.DDA.SafeArray.Delegate.ParameterMap.sourceData }
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.EventDispId.dispidObjectDataChangedEvent,
	toHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.Id, value: 0 }
	],
	fromHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.Id, value: OSF.DDA.SafeArray.Delegate.ParameterMap.sourceData }
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.EventDispId.dispidContentControlAddedEvent,
	toHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.Id, value: 0 }
	],
	fromHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.Id, value: OSF.DDA.SafeArray.Delegate.ParameterMap.sourceData }
	]
});
OSF.OUtil.augmentList(Microsoft.Office.WebExtension.EventType, {
	BindingSelectionChanged: "bindingSelectionChanged",
	BindingDataChanged: "bindingDataChanged"
});
OSF.OUtil.augmentList(OSF.DDA.EventDescriptors, { BindingSelectionChangedEvent: "BindingSelectionChangedEvent" });
OSF.DDA.BindingSelectionChangedEventArgs=function OSF_DDA_BindingSelectionChangedEventArgs(bindingInstance, subset) {
	OSF.OUtil.defineEnumerableProperties(this, {
		"type": {
			value: Microsoft.Office.WebExtension.EventType.BindingSelectionChanged
		},
		"binding": {
			value: bindingInstance
		}
	});
	for (var prop in subset) {
		OSF.OUtil.defineEnumerableProperty(this, prop, {
			value: subset[prop]
		});
	}
};
OSF.DDA.BindingDataChangedEventArgs=function OSF_DDA_BindingDataChangedEventArgs(bindingInstance) {
	OSF.OUtil.defineEnumerableProperties(this, {
		"type": {
			value: Microsoft.Office.WebExtension.EventType.BindingDataChanged
		},
		"binding": {
			value: bindingInstance
		}
	});
};
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.EventDescriptors.BindingSelectionChangedEvent,
	fromHost: [
		{ name: OSF.DDA.PropertyDescriptors.BindingProperties, value: 0 },
		{ name: OSF.DDA.PropertyDescriptors.Subset, value: 1 }
	],
	isComplexType: true
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.EventDispId.dispidBindingSelectionChangedEvent,
	fromHost: [
		{ name: OSF.DDA.EventDescriptors.BindingSelectionChangedEvent, value: OSF.DDA.SafeArray.Delegate.ParameterMap.self }
	],
	isComplexType: true
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.EventDispId.dispidBindingDataChangedEvent,
	fromHost: [{ name: OSF.DDA.PropertyDescriptors.BindingProperties, value: OSF.DDA.SafeArray.Delegate.ParameterMap.self }]
});
OSF.OUtil.augmentList(Microsoft.Office.WebExtension.FilterType, { OnlyVisible: "onlyVisible" });
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: Microsoft.Office.WebExtension.Parameters.FilterType,
	toHost: [{ name: Microsoft.Office.WebExtension.FilterType.OnlyVisible, value: 1 }]
});
Microsoft.Office.WebExtension.GoToType={
	Binding: "binding",
	NamedItem: "namedItem",
	Slide: "slide",
	Index: "index"
};
Microsoft.Office.WebExtension.SelectionMode={
	Default: "default",
	Selected: "selected",
	None: "none"
};
Microsoft.Office.WebExtension.Index={
	First: "first",
	Last: "last",
	Next: "next",
	Previous: "previous"
};
OSF.DDA.AsyncMethodNames.addNames({ GoToByIdAsync: "goToByIdAsync" });
OSF.DDA.AsyncMethodCalls.define({
	method: OSF.DDA.AsyncMethodNames.GoToByIdAsync,
	requiredArguments: [{
			"name": Microsoft.Office.WebExtension.Parameters.Id,
			"types": ["string", "number"]
		},
		{
			"name": Microsoft.Office.WebExtension.Parameters.GoToType,
			"enum": Microsoft.Office.WebExtension.GoToType
		}
	],
	supportedOptions: [
		{
			name: Microsoft.Office.WebExtension.Parameters.SelectionMode,
			value: {
				"enum": Microsoft.Office.WebExtension.SelectionMode,
				"defaultValue": Microsoft.Office.WebExtension.SelectionMode.Default
			}
		}
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: Microsoft.Office.WebExtension.Parameters.GoToType,
	toHost: [
		{ name: Microsoft.Office.WebExtension.GoToType.Binding, value: 0 },
		{ name: Microsoft.Office.WebExtension.GoToType.NamedItem, value: 1 },
		{ name: Microsoft.Office.WebExtension.GoToType.Slide, value: 2 },
		{ name: Microsoft.Office.WebExtension.GoToType.Index, value: 3 }
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: Microsoft.Office.WebExtension.Parameters.SelectionMode,
	toHost: [
		{ name: Microsoft.Office.WebExtension.SelectionMode.Default, value: 0 },
		{ name: Microsoft.Office.WebExtension.SelectionMode.Selected, value: 1 },
		{ name: Microsoft.Office.WebExtension.SelectionMode.None, value: 2 }
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.MethodDispId.dispidNavigateToMethod,
	toHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.Id, value: 0 },
		{ name: Microsoft.Office.WebExtension.Parameters.GoToType, value: 1 },
		{ name: Microsoft.Office.WebExtension.Parameters.SelectionMode, value: 2 }
	]
});
OSF.OUtil.augmentList(Microsoft.Office.WebExtension.EventType, { RichApiMessage: "richApiMessage" });
OSF.DDA.RichApiMessageEventArgs=function OSF_DDA_RichApiMessageEventArgs(eventType, eventProperties) {
	var entryArray=eventProperties[Microsoft.Office.WebExtension.Parameters.Data];
	var entries=[];
	if (entryArray) {
		for (var i=0; i < entryArray.length; i++) {
			var elem=entryArray[i];
			if (elem.toArray) {
				elem=elem.toArray();
			}
			entries.push({
				messageCategory: elem[0],
				messageType: elem[1],
				targetId: elem[2],
				message: elem[3],
				id: elem[4],
				isRemoteOverride: elem[5]
			});
		}
	}
	OSF.OUtil.defineEnumerableProperties(this, {
		"type": { value: Microsoft.Office.WebExtension.EventType.RichApiMessage },
		"entries": { value: entries }
	});
};
var OfficeExt;
(function (OfficeExt) {
	var RichApiMessageManager=(function () {
		function RichApiMessageManager() {
			this._eventDispatch=null;
			this._eventDispatch=new OSF.EventDispatch([
				Microsoft.Office.WebExtension.EventType.RichApiMessage,
			]);
			OSF.DDA.DispIdHost.addEventSupport(this, this._eventDispatch);
		}
		return RichApiMessageManager;
	})();
	OfficeExt.RichApiMessageManager=RichApiMessageManager;
})(OfficeExt || (OfficeExt={}));
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.EventDispId.dispidRichApiMessageEvent,
	toHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.Data, value: 0 }
	],
	fromHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.Data, value: OSF.DDA.SafeArray.Delegate.ParameterMap.sourceData }
	]
});
OSF.DDA.AsyncMethodNames.addNames({
	ExecuteRichApiRequestAsync: "executeRichApiRequestAsync"
});
OSF.DDA.AsyncMethodCalls.define({
	method: OSF.DDA.AsyncMethodNames.ExecuteRichApiRequestAsync,
	requiredArguments: [
		{
			name: Microsoft.Office.WebExtension.Parameters.Data,
			types: ["object"]
		}
	],
	supportedOptions: []
});
OSF.OUtil.setNamespace("RichApi", OSF.DDA);
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.MethodDispId.dispidExecuteRichApiRequestMethod,
	toHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.Data, value: 0 }
	],
	fromHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.Data, value: OSF.DDA.SafeArray.Delegate.ParameterMap.self }
	]
});
Microsoft.Office.WebExtension.FileType={
	Text: "text",
	Compressed: "compressed",
	Pdf: "pdf"
};
OSF.OUtil.augmentList(OSF.DDA.PropertyDescriptors, {
	FileProperties: "FileProperties",
	FileSliceProperties: "FileSliceProperties"
});
OSF.DDA.FileProperties={
	Handle: "FileHandle",
	FileSize: "FileSize",
	SliceSize: Microsoft.Office.WebExtension.Parameters.SliceSize
};
OSF.DDA.File=function OSF_DDA_File(handle, fileSize, sliceSize) {
	OSF.OUtil.defineEnumerableProperties(this, {
		"size": {
			value: fileSize
		},
		"sliceCount": {
			value: Math.ceil(fileSize / sliceSize)
		}
	});
	var privateState={};
	privateState[OSF.DDA.FileProperties.Handle]=handle;
	privateState[OSF.DDA.FileProperties.SliceSize]=sliceSize;
	var am=OSF.DDA.AsyncMethodNames;
	OSF.DDA.DispIdHost.addAsyncMethods(this, [
		am.GetDocumentCopyChunkAsync,
		am.ReleaseDocumentCopyAsync
	], privateState);
};
OSF.DDA.FileSliceOffset="fileSliceoffset";
OSF.DDA.AsyncMethodNames.addNames({
	GetDocumentCopyAsync: "getFileAsync",
	GetDocumentCopyChunkAsync: "getSliceAsync",
	ReleaseDocumentCopyAsync: "closeAsync"
});
OSF.DDA.AsyncMethodCalls.define({
	method: OSF.DDA.AsyncMethodNames.GetDocumentCopyAsync,
	requiredArguments: [
		{
			"name": Microsoft.Office.WebExtension.Parameters.FileType,
			"enum": Microsoft.Office.WebExtension.FileType
		}
	],
	supportedOptions: [
		{
			name: Microsoft.Office.WebExtension.Parameters.SliceSize,
			value: {
				"types": ["number"],
				"defaultValue": 4 * 1024 * 1024
			}
		}
	],
	checkCallArgs: function (callArgs, caller, stateInfo) {
		var sliceSize=callArgs[Microsoft.Office.WebExtension.Parameters.SliceSize];
		if (sliceSize <=0 || sliceSize > (4 * 1024 * 1024)) {
			throw OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidSliceSize;
		}
		return callArgs;
	},
	onSucceeded: function (fileDescriptor, caller, callArgs) {
		return new OSF.DDA.File(fileDescriptor[OSF.DDA.FileProperties.Handle], fileDescriptor[OSF.DDA.FileProperties.FileSize], callArgs[Microsoft.Office.WebExtension.Parameters.SliceSize]);
	}
});
OSF.DDA.AsyncMethodCalls.define({
	method: OSF.DDA.AsyncMethodNames.GetDocumentCopyChunkAsync,
	requiredArguments: [
		{
			"name": Microsoft.Office.WebExtension.Parameters.SliceIndex,
			"types": ["number"]
		}
	],
	privateStateCallbacks: [
		{
			name: OSF.DDA.FileProperties.Handle,
			value: function (caller, stateInfo) { return stateInfo[OSF.DDA.FileProperties.Handle]; }
		},
		{
			name: OSF.DDA.FileProperties.SliceSize,
			value: function (caller, stateInfo) { return stateInfo[OSF.DDA.FileProperties.SliceSize]; }
		}
	],
	checkCallArgs: function (callArgs, caller, stateInfo) {
		var index=callArgs[Microsoft.Office.WebExtension.Parameters.SliceIndex];
		if (index < 0 || index >=caller.sliceCount) {
			throw OSF.DDA.ErrorCodeManager.errorCodes.ooeIndexOutOfRange;
		}
		callArgs[OSF.DDA.FileSliceOffset]=parseInt((index * stateInfo[OSF.DDA.FileProperties.SliceSize]).toString());
		return callArgs;
	},
	onSucceeded: function (sliceDescriptor, caller, callArgs) {
		var slice={};
		OSF.OUtil.defineEnumerableProperties(slice, {
			"data": {
				value: OSF.OUtil.shallowCopy(sliceDescriptor[Microsoft.Office.WebExtension.Parameters.Data])
			},
			"index": {
				value: callArgs[Microsoft.Office.WebExtension.Parameters.SliceIndex]
			},
			"size": {
				value: sliceDescriptor[OSF.DDA.FileProperties.SliceSize]
			}
		});
		return slice;
	}
});
OSF.DDA.AsyncMethodCalls.define({
	method: OSF.DDA.AsyncMethodNames.ReleaseDocumentCopyAsync,
	privateStateCallbacks: [
		{
			name: OSF.DDA.FileProperties.Handle,
			value: function (caller, stateInfo) { return stateInfo[OSF.DDA.FileProperties.Handle]; }
		}
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.PropertyDescriptors.FileProperties,
	fromHost: [
		{ name: OSF.DDA.FileProperties.Handle, value: 0 },
		{ name: OSF.DDA.FileProperties.FileSize, value: 1 }
	],
	isComplexType: true
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.PropertyDescriptors.FileSliceProperties,
	fromHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.Data, value: 0 },
		{ name: OSF.DDA.FileProperties.SliceSize, value: 1 }
	],
	isComplexType: true
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: Microsoft.Office.WebExtension.Parameters.FileType,
	toHost: [
		{ name: Microsoft.Office.WebExtension.FileType.Text, value: 0 },
		{ name: Microsoft.Office.WebExtension.FileType.Compressed, value: 5 },
		{ name: Microsoft.Office.WebExtension.FileType.Pdf, value: 6 }
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.MethodDispId.dispidGetDocumentCopyMethod,
	toHost: [{ name: Microsoft.Office.WebExtension.Parameters.FileType, value: 0 }],
	fromHost: [
		{ name: OSF.DDA.PropertyDescriptors.FileProperties, value: OSF.DDA.SafeArray.Delegate.ParameterMap.self }
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.MethodDispId.dispidGetDocumentCopyChunkMethod,
	toHost: [
		{ name: OSF.DDA.FileProperties.Handle, value: 0 },
		{ name: OSF.DDA.FileSliceOffset, value: 1 },
		{ name: OSF.DDA.FileProperties.SliceSize, value: 2 }
	],
	fromHost: [
		{ name: OSF.DDA.PropertyDescriptors.FileSliceProperties, value: OSF.DDA.SafeArray.Delegate.ParameterMap.self }
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.MethodDispId.dispidReleaseDocumentCopyMethod,
	toHost: [{ name: OSF.DDA.FileProperties.Handle, value: 0 }]
});
OSF.DDA.FilePropertiesDescriptor={
	Url: "Url"
};
OSF.OUtil.augmentList(OSF.DDA.PropertyDescriptors, {
	FilePropertiesDescriptor: "FilePropertiesDescriptor"
});
Microsoft.Office.WebExtension.FileProperties=function Microsoft_Office_WebExtension_FileProperties(filePropertiesDescriptor) {
	OSF.OUtil.defineEnumerableProperties(this, {
		"url": {
			value: filePropertiesDescriptor[OSF.DDA.FilePropertiesDescriptor.Url]
		}
	});
};
OSF.DDA.AsyncMethodNames.addNames({ GetFilePropertiesAsync: "getFilePropertiesAsync" });
OSF.DDA.AsyncMethodCalls.define({
	method: OSF.DDA.AsyncMethodNames.GetFilePropertiesAsync,
	fromHost: [
		{ name: OSF.DDA.PropertyDescriptors.FilePropertiesDescriptor, value: 0 }
	],
	requiredArguments: [],
	supportedOptions: [],
	onSucceeded: function (filePropertiesDescriptor, caller, callArgs) {
		return new Microsoft.Office.WebExtension.FileProperties(filePropertiesDescriptor);
	}
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.PropertyDescriptors.FilePropertiesDescriptor,
	fromHost: [
		{ name: OSF.DDA.FilePropertiesDescriptor.Url, value: 0 }
	],
	isComplexType: true
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.MethodDispId.dispidGetFilePropertiesMethod,
	fromHost: [
		{ name: OSF.DDA.PropertyDescriptors.FilePropertiesDescriptor, value: OSF.DDA.SafeArray.Delegate.ParameterMap.self }
	]
});
OSF.DDA.ExcelTableBinding=function OSF_DDA_ExcelTableBinding(id, docInstance, rows, cols, hasHeaders) {
	var am=OSF.DDA.AsyncMethodNames;
	OSF.DDA.DispIdHost.addAsyncMethods(this, [
		am.ClearFormatsAsync,
		am.SetTableOptionsAsync,
		am.SetFormatsAsync
	]);
	OSF.DDA.ExcelTableBinding.uber.constructor.call(this, id, docInstance, rows, cols, hasHeaders);
	OSF.OUtil.finalizeProperties(this);
};
OSF.OUtil.extend(OSF.DDA.ExcelTableBinding, OSF.DDA.TableBinding);
(function () {
	OSF.DDA.AsyncMethodCalls.define({
		method: OSF.DDA.AsyncMethodNames.SetSelectedDataAsync,
		requiredArguments: [
			{
				"name": Microsoft.Office.WebExtension.Parameters.Data,
				"types": ["string", "object", "number", "boolean"]
			}
		],
		supportedOptions: [{
				name: Microsoft.Office.WebExtension.Parameters.CoercionType,
				value: {
					"enum": Microsoft.Office.WebExtension.CoercionType,
					"calculate": function (requiredArgs) { return OSF.DDA.DataCoercion.determineCoercionType(requiredArgs[Microsoft.Office.WebExtension.Parameters.Data]); }
				}
			},
			{
				name: Microsoft.Office.WebExtension.Parameters.CellFormat,
				value: {
					"types": ["object"],
					"defaultValue": []
				}
			},
			{
				name: Microsoft.Office.WebExtension.Parameters.TableOptions,
				value: {
					"types": ["object"],
					"defaultValue": []
				}
			}
		],
		privateStateCallbacks: []
	});
	OSF.DDA.AsyncMethodCalls.define({
		method: OSF.DDA.AsyncMethodNames.SetDataAsync,
		requiredArguments: [
			{
				"name": Microsoft.Office.WebExtension.Parameters.Data,
				"types": ["string", "object", "number", "boolean"]
			}
		],
		supportedOptions: [{
				name: Microsoft.Office.WebExtension.Parameters.CoercionType,
				value: {
					"enum": Microsoft.Office.WebExtension.CoercionType,
					"calculate": function (requiredArgs) { return OSF.DDA.DataCoercion.determineCoercionType(requiredArgs[Microsoft.Office.WebExtension.Parameters.Data]); }
				}
			},
			{
				name: Microsoft.Office.WebExtension.Parameters.Rows,
				value: {
					"types": ["object", "string"],
					"defaultValue": null
				}
			},
			{
				name: Microsoft.Office.WebExtension.Parameters.Columns,
				value: {
					"types": ["object"],
					"defaultValue": null
				}
			},
			{
				name: Microsoft.Office.WebExtension.Parameters.StartRow,
				value: {
					"types": ["number"],
					"defaultValue": 0
				}
			},
			{
				name: Microsoft.Office.WebExtension.Parameters.StartColumn,
				value: {
					"types": ["number"],
					"defaultValue": 0
				}
			},
			{
				name: Microsoft.Office.WebExtension.Parameters.CellFormat,
				value: {
					"types": ["object"],
					"defaultValue": []
				}
			},
			{
				name: Microsoft.Office.WebExtension.Parameters.TableOptions,
				value: {
					"types": ["object"],
					"defaultValue": []
				}
			}
		],
		checkCallArgs: function (callArgs, caller, stateInfo) {
			var Parameters=Microsoft.Office.WebExtension.Parameters;
			if (callArgs[Parameters.StartRow]==0 &&
				callArgs[Parameters.StartColumn]==0 &&
				OSF.OUtil.isArray(callArgs[Parameters.CellFormat]) && callArgs[Parameters.CellFormat].length===0 &&
				OSF.OUtil.isArray(callArgs[Parameters.TableOptions]) && callArgs[Parameters.TableOptions].length===0) {
				delete callArgs[Parameters.StartRow];
				delete callArgs[Parameters.StartColumn];
				delete callArgs[Parameters.CellFormat];
				delete callArgs[Parameters.TableOptions];
			}
			if (callArgs[Parameters.CoercionType] !=OSF.DDA.DataCoercion.getCoercionDefaultForBinding(caller.type) &&
				((callArgs[Parameters.StartRow] && callArgs[Parameters.StartRow] !=0) ||
					(callArgs[Parameters.StartColumn] && callArgs[Parameters.StartColumn] !=0) ||
					callArgs[Parameters.CellFormat] ||
					callArgs[Parameters.TableOptions])) {
				throw OSF.DDA.ErrorCodeManager.errorCodes.ooeCoercionTypeNotMatchBinding;
			}
			return callArgs;
		},
		privateStateCallbacks: [
			{
				name: Microsoft.Office.WebExtension.Parameters.Id,
				value: function (obj) { return obj.id; }
			}
		]
	});
	OSF.DDA.BindingPromise.prototype.setTableOptionsAsync=function OSF_DDA_BindingPromise$setTableOptionsAsync() {
		var args=arguments;
		this._fetch(function onComplete(binding) { binding.setTableOptionsAsync.apply(binding, args); });
		return this;
	},
		OSF.DDA.BindingPromise.prototype.setFormatsAsync=function OSF_DDA_BindingPromise$setFormatsAsync() {
			var args=arguments;
			this._fetch(function onComplete(binding) { binding.setFormatsAsync.apply(binding, args); });
			return this;
		},
		OSF.DDA.BindingPromise.prototype.clearFormatsAsync=function OSF_DDA_BindingPromise$clearFormatsAsync() {
			var args=arguments;
			this._fetch(function onComplete(binding) { binding.clearFormatsAsync.apply(binding, args); });
			return this;
		};
})();
(function () {
	function getObjectId(obj) { return obj.id; }
	OSF.DDA.AsyncMethodNames.addNames({
		ClearFormatsAsync: "clearFormatsAsync",
		SetTableOptionsAsync: "setTableOptionsAsync",
		SetFormatsAsync: "setFormatsAsync"
	});
	OSF.DDA.AsyncMethodCalls.define({
		method: OSF.DDA.AsyncMethodNames.ClearFormatsAsync,
		requiredArguments: [],
		supportedOptions: [],
		privateStateCallbacks: [
			{
				name: Microsoft.Office.WebExtension.Parameters.Id,
				value: getObjectId
			}
		]
	});
	OSF.DDA.AsyncMethodCalls.define({
		method: OSF.DDA.AsyncMethodNames.SetTableOptionsAsync,
		requiredArguments: [
			{
				"name": Microsoft.Office.WebExtension.Parameters.TableOptions,
				"defaultValue": []
			}
		],
		privateStateCallbacks: [
			{
				name: Microsoft.Office.WebExtension.Parameters.Id,
				value: getObjectId
			}
		]
	});
	OSF.DDA.AsyncMethodCalls.define({
		method: OSF.DDA.AsyncMethodNames.SetFormatsAsync,
		requiredArguments: [
			{
				"name": Microsoft.Office.WebExtension.Parameters.CellFormat,
				"defaultValue": []
			}
		],
		privateStateCallbacks: [
			{
				name: Microsoft.Office.WebExtension.Parameters.Id,
				value: getObjectId
			}
		]
	});
})();
Microsoft.Office.WebExtension.Table={
	All: 0,
	Data: 1,
	Headers: 2
};
(function () {
	OSF.DDA.SafeArray.Delegate.ParameterMap.define({
		type: OSF.DDA.MethodDispId.dispidClearFormatsMethod,
		toHost: [
			{ name: Microsoft.Office.WebExtension.Parameters.Id, value: 0 }
		]
	});
	OSF.DDA.SafeArray.Delegate.ParameterMap.define({
		type: OSF.DDA.MethodDispId.dispidSetTableOptionsMethod,
		toHost: [
			{ name: Microsoft.Office.WebExtension.Parameters.Id, value: 0 },
			{ name: Microsoft.Office.WebExtension.Parameters.TableOptions, value: 1 },
		]
	});
	OSF.DDA.SafeArray.Delegate.ParameterMap.define({
		type: OSF.DDA.MethodDispId.dispidSetFormatsMethod,
		toHost: [
			{ name: Microsoft.Office.WebExtension.Parameters.Id, value: 0 },
			{ name: Microsoft.Office.WebExtension.Parameters.CellFormat, value: 1 },
		]
	});
	OSF.DDA.SafeArray.Delegate.ParameterMap.define({
		type: OSF.DDA.MethodDispId.dispidSetSelectedDataMethod,
		toHost: [
			{ name: Microsoft.Office.WebExtension.Parameters.CoercionType, value: 0 },
			{ name: Microsoft.Office.WebExtension.Parameters.Data, value: 1 },
			{ name: Microsoft.Office.WebExtension.Parameters.CellFormat, value: 2 },
			{ name: Microsoft.Office.WebExtension.Parameters.TableOptions, value: 3 }
		]
	});
	OSF.DDA.SafeArray.Delegate.ParameterMap.define({
		type: OSF.DDA.MethodDispId.dispidSetBindingDataMethod,
		toHost: [
			{ name: Microsoft.Office.WebExtension.Parameters.Id, value: 0 },
			{ name: Microsoft.Office.WebExtension.Parameters.CoercionType, value: 1 },
			{ name: Microsoft.Office.WebExtension.Parameters.Data, value: 2 },
			{ name: OSF.DDA.SafeArray.UniqueArguments.Offset, value: 3 },
			{ name: Microsoft.Office.WebExtension.Parameters.CellFormat, value: 4 },
			{ name: Microsoft.Office.WebExtension.Parameters.TableOptions, value: 5 }
		]
	});
	var tableOptionProperties={
		headerRow: 0,
		bandedRows: 1,
		firstColumn: 2,
		lastColumn: 3,
		bandedColumns: 4,
		filterButton: 5,
		style: 6,
		totalRow: 7
	};
	var cellProperties={
		row: 0,
		column: 1
	};
	var formatProperties={
		alignHorizontal: { text: "alignHorizontal", type: 1 },
		alignVertical: { text: "alignVertical", type: 2 },
		backgroundColor: { text: "backgroundColor", type: 101 },
		borderStyle: { text: "borderStyle", type: 201 },
		borderColor: { text: "borderColor", type: 202 },
		borderTopStyle: { text: "borderTopStyle", type: 203 },
		borderTopColor: { text: "borderTopColor", type: 204 },
		borderBottomStyle: { text: "borderBottomStyle", type: 205 },
		borderBottomColor: { text: "borderBottomColor", type: 206 },
		borderLeftStyle: { text: "borderLeftStyle", type: 207 },
		borderLeftColor: { text: "borderLeftColor", type: 208 },
		borderRightStyle: { text: "borderRightStyle", type: 209 },
		borderRightColor: { text: "borderRightColor", type: 210 },
		borderOutlineStyle: { text: "borderOutlineStyle", type: 211 },
		borderOutlineColor: { text: "borderOutlineColor", type: 212 },
		borderInlineStyle: { text: "borderInlineStyle", type: 213 },
		borderInlineColor: { text: "borderInlineColor", type: 214 },
		fontFamily: { text: "fontFamily", type: 301 },
		fontStyle: { text: "fontStyle", type: 302 },
		fontSize: { text: "fontSize", type: 303 },
		fontUnderlineStyle: { text: "fontUnderlineStyle", type: 304 },
		fontColor: { text: "fontColor", type: 305 },
		fontDirection: { text: "fontDirection", type: 306 },
		fontStrikethrough: { text: "fontStrikethrough", type: 307 },
		fontSuperscript: { text: "fontSuperscript", type: 308 },
		fontSubscript: { text: "fontSubscript", type: 309 },
		fontNormal: { text: "fontNormal", type: 310 },
		indentLeft: { text: "indentLeft", type: 401 },
		indentRight: { text: "indentRight", type: 402 },
		numberFormat: { text: "numberFormat", type: 501 },
		width: { text: "width", type: 701 },
		height: { text: "height", type: 702 },
		wrapping: { text: "wrapping", type: 703 }
	};
	var borderStyleSet=[
		{ name: "none", value: 0 },
		{ name: "thin", value: 1 },
		{ name: "medium", value: 2 },
		{ name: "dashed", value: 3 },
		{ name: "dotted", value: 4 },
		{ name: "thick", value: 5 },
		{ name: "double", value: 6 },
		{ name: "hair", value: 7 },
		{ name: "medium dashed", value: 8 },
		{ name: "dash dot", value: 9 },
		{ name: "medium dash dot", value: 10 },
		{ name: "dash dot dot", value: 11 },
		{ name: "medium dash dot dot", value: 12 },
		{ name: "slant dash dot", value: 13 },
	];
	var colorSet=[
		{ name: "none", value: 0 },
		{ name: "black", value: 1 },
		{ name: "blue", value: 2 },
		{ name: "gray", value: 3 },
		{ name: "green", value: 4 },
		{ name: "orange", value: 5 },
		{ name: "pink", value: 6 },
		{ name: "purple", value: 7 },
		{ name: "red", value: 8 },
		{ name: "teal", value: 9 },
		{ name: "turquoise", value: 10 },
		{ name: "violet", value: 11 },
		{ name: "white", value: 12 },
		{ name: "yellow", value: 13 },
		{ name: "automatic", value: 14 },
	];
	var ns=OSF.DDA.SafeArray.Delegate.ParameterMap;
	ns.define({
		type: formatProperties.alignHorizontal.text,
		toHost: [
			{ name: "general", value: 0 },
			{ name: "left", value: 1 },
			{ name: "center", value: 2 },
			{ name: "right", value: 3 },
			{ name: "fill", value: 4 },
			{ name: "justify", value: 5 },
			{ name: "center across selection", value: 6 },
			{ name: "distributed", value: 7 },
		]
	});
	ns.define({
		type: formatProperties.alignVertical.text,
		toHost: [
			{ name: "top", value: 0 },
			{ name: "center", value: 1 },
			{ name: "bottom", value: 2 },
			{ name: "justify", value: 3 },
			{ name: "distributed", value: 4 },
		]
	});
	ns.define({
		type: formatProperties.backgroundColor.text,
		toHost: colorSet
	});
	ns.define({
		type: formatProperties.borderStyle.text,
		toHost: borderStyleSet
	});
	ns.define({
		type: formatProperties.borderColor.text,
		toHost: colorSet
	});
	ns.define({
		type: formatProperties.borderTopStyle.text,
		toHost: borderStyleSet
	});
	ns.define({
		type: formatProperties.borderTopColor.text,
		toHost: colorSet
	});
	ns.define({
		type: formatProperties.borderBottomStyle.text,
		toHost: borderStyleSet
	});
	ns.define({
		type: formatProperties.borderBottomColor.text,
		toHost: colorSet
	});
	ns.define({
		type: formatProperties.borderLeftStyle.text,
		toHost: borderStyleSet
	});
	ns.define({
		type: formatProperties.borderLeftColor.text,
		toHost: colorSet
	});
	ns.define({
		type: formatProperties.borderRightStyle.text,
		toHost: borderStyleSet
	});
	ns.define({
		type: formatProperties.borderRightColor.text,
		toHost: colorSet
	});
	ns.define({
		type: formatProperties.borderOutlineStyle.text,
		toHost: borderStyleSet
	});
	ns.define({
		type: formatProperties.borderOutlineColor.text,
		toHost: colorSet
	});
	ns.define({
		type: formatProperties.borderInlineStyle.text,
		toHost: borderStyleSet
	});
	ns.define({
		type: formatProperties.borderInlineColor.text,
		toHost: colorSet
	});
	ns.define({
		type: formatProperties.fontStyle.text,
		toHost: [
			{ name: "regular", value: 0 },
			{ name: "italic", value: 1 },
			{ name: "bold", value: 2 },
			{ name: "bold italic", value: 3 },
		]
	});
	ns.define({
		type: formatProperties.fontUnderlineStyle.text,
		toHost: [
			{ name: "none", value: 0 },
			{ name: "single", value: 1 },
			{ name: "double", value: 2 },
			{ name: "single accounting", value: 3 },
			{ name: "double accounting", value: 4 },
		]
	});
	ns.define({
		type: formatProperties.fontColor.text,
		toHost: colorSet
	});
	ns.define({
		type: formatProperties.fontDirection.text,
		toHost: [
			{ name: "context", value: 0 },
			{ name: "left-to-right", value: 1 },
			{ name: "right-to-left", value: 2 },
		]
	});
	ns.define({
		type: formatProperties.width.text,
		toHost: [
			{ name: "auto fit", value: -1 },
		]
	});
	ns.define({
		type: formatProperties.height.text,
		toHost: [
			{ name: "auto fit", value: -1 },
		]
	});
	ns.define({
		type: Microsoft.Office.WebExtension.Parameters.TableOptions,
		toHost: [
			{ name: "headerRow", value: 0 },
			{ name: "bandedRows", value: 1 },
			{ name: "firstColumn", value: 2 },
			{ name: "lastColumn", value: 3 },
			{ name: "bandedColumns", value: 4 },
			{ name: "filterButton", value: 5 },
			{ name: "style", value: 6 },
			{ name: "totalRow", value: 7 }
		]
	});
	ns.dynamicTypes[Microsoft.Office.WebExtension.Parameters.CellFormat]={
		toHost: function (data) {
			for (var entry in data) {
				if (data[entry].format) {
					data[entry].format=ns.doMapValues(data[entry].format, "toHost");
				}
			}
			return data;
		},
		fromHost: function (args) {
			return args;
		}
	};
	ns.setDynamicType(Microsoft.Office.WebExtension.Parameters.CellFormat, {
		toHost: function OSF_DDA_SafeArray_Delegate_SpecialProcessor_CellFormat$toHost(cellFormats) {
			var textCells="cells";
			var textFormat="format";
			var posCells=0;
			var posFormat=1;
			var ret=[];
			for (var index in cellFormats) {
				var cfOld=cellFormats[index];
				var cfNew=[];
				if (typeof (cfOld[textCells]) !=='undefined') {
					var cellsOld=cfOld[textCells];
					var cellsNew;
					if (typeof cfOld[textCells]==="object") {
						cellsNew=[];
						for (var entry in cellsOld) {
							if (typeof (cellProperties[entry]) !=='undefined') {
								cellsNew[cellProperties[entry]]=cellsOld[entry];
							}
						}
					}
					else {
						cellsNew=cellsOld;
					}
					cfNew[posCells]=cellsNew;
				}
				if (cfOld[textFormat]) {
					var formatOld=cfOld[textFormat];
					var formatNew=[];
					for (var entry2 in formatOld) {
						if (typeof (formatProperties[entry2]) !=='undefined') {
							formatNew.push([
								formatProperties[entry2].type,
								formatOld[entry2]
							]);
						}
					}
					cfNew[posFormat]=formatNew;
				}
				ret[index]=cfNew;
			}
			return ret;
		},
		fromHost: function OSF_DDA_SafeArray_Delegate_SpecialProcessor_CellFormat$fromHost(hostArgs) {
			return hostArgs;
		}
	});
	ns.setDynamicType(Microsoft.Office.WebExtension.Parameters.TableOptions, {
		toHost: function OSF_DDA_SafeArray_Delegate_SpecialProcessor_TableOptions$toHost(tableOptions) {
			var ret=[];
			for (var entry in tableOptions) {
				if (typeof (tableOptionProperties[entry]) !=='undefined') {
					ret[tableOptionProperties[entry]]=tableOptions[entry];
				}
			}
			return ret;
		},
		fromHost: function OSF_DDA_SafeArray_Delegate_SpecialProcessor_TableOptions$fromHost(hostArgs) {
			return hostArgs;
		}
	});
})();
OSF.OUtil.augmentList(Microsoft.Office.WebExtension.CoercionType, { Image: "image" });
OSF.OUtil.augmentList(Microsoft.Office.WebExtension.CoercionType, { XmlSvg: "xmlsvg" });
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: Microsoft.Office.WebExtension.Parameters.CoercionType,
	toHost: [
		{ name: Microsoft.Office.WebExtension.CoercionType.Image, value: 8 },
		{ name: Microsoft.Office.WebExtension.CoercionType.XmlSvg, value: 9 }
	]
});
var OfficeExt;
(function (OfficeExt) {
	var AppCommand;
	(function (AppCommand) {
		var AppCommandManager=(function () {
			function AppCommandManager() {
				var _this=this;
				this._pseudoDocument=null;
				this._eventDispatch=null;
				this._processAppCommandInvocation=function (args) {
					var verifyResult=_this._verifyManifestCallback(args.callbackName);
					if (verifyResult.errorCode !=OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess) {
						_this._invokeAppCommandCompletedMethod(args.appCommandId, verifyResult.errorCode, "");
						return;
					}
					var eventObj=_this._constructEventObjectForCallback(args);
					if (eventObj) {
						window.setTimeout(function () { verifyResult.callback(eventObj); }, 0);
					}
					else {
						_this._invokeAppCommandCompletedMethod(args.appCommandId, OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError, "");
					}
				};
			}
			AppCommandManager.initializeOsfDda=function () {
				OSF.DDA.AsyncMethodNames.addNames({
					AppCommandInvocationCompletedAsync: "appCommandInvocationCompletedAsync"
				});
				OSF.DDA.AsyncMethodCalls.define({
					method: OSF.DDA.AsyncMethodNames.AppCommandInvocationCompletedAsync,
					requiredArguments: [{
							"name": Microsoft.Office.WebExtension.Parameters.Id,
							"types": ["string"]
						},
						{
							"name": Microsoft.Office.WebExtension.Parameters.Status,
							"types": ["number"]
						},
						{
							"name": Microsoft.Office.WebExtension.Parameters.AppCommandInvocationCompletedData,
							"types": ["string"]
						}
					]
				});
				OSF.OUtil.augmentList(OSF.DDA.EventDescriptors, {
					AppCommandInvokedEvent: "AppCommandInvokedEvent"
				});
				OSF.OUtil.augmentList(Microsoft.Office.WebExtension.EventType, {
					AppCommandInvoked: "appCommandInvoked"
				});
				OSF.OUtil.setNamespace("AppCommand", OSF.DDA);
				OSF.DDA.AppCommand.AppCommandInvokedEventArgs=OfficeExt.AppCommand.AppCommandInvokedEventArgs;
			};
			AppCommandManager.prototype.initializeAndChangeOnce=function (callback) {
				AppCommand.registerDdaFacade();
				this._pseudoDocument={};
				OSF.DDA.DispIdHost.addAsyncMethods(this._pseudoDocument, [
					OSF.DDA.AsyncMethodNames.AppCommandInvocationCompletedAsync,
				]);
				this._eventDispatch=new OSF.EventDispatch([
					Microsoft.Office.WebExtension.EventType.AppCommandInvoked,
				]);
				var onRegisterCompleted=function (result) {
					if (callback) {
						if (result.status=="succeeded") {
							callback(OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess);
						}
						else {
							callback(OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError);
						}
					}
				};
				OSF.DDA.DispIdHost.addEventSupport(this._pseudoDocument, this._eventDispatch);
				this._pseudoDocument.addHandlerAsync(Microsoft.Office.WebExtension.EventType.AppCommandInvoked, this._processAppCommandInvocation, onRegisterCompleted);
			};
			AppCommandManager.prototype._verifyManifestCallback=function (callbackName) {
				var defaultResult={ callback: null, errorCode: OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidCallback };
				callbackName=callbackName.trim();
				try {
					var callList=callbackName.split(".");
					var parentObject=window;
					for (var i=0; i < callList.length - 1; i++) {
						if (parentObject[callList[i]] && (typeof parentObject[callList[i]]=="object" || typeof parentObject[callList[i]]=="function")) {
							parentObject=parentObject[callList[i]];
						}
						else {
							return defaultResult;
						}
					}
					var callbackFunc=parentObject[callList[callList.length - 1]];
					if (typeof callbackFunc !="function") {
						return defaultResult;
					}
				}
				catch (e) {
					return defaultResult;
				}
				return { callback: callbackFunc, errorCode: OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess };
			};
			AppCommandManager.prototype._invokeAppCommandCompletedMethod=function (appCommandId, resultCode, data) {
				this._pseudoDocument.appCommandInvocationCompletedAsync(appCommandId, resultCode, data);
			};
			AppCommandManager.prototype._constructEventObjectForCallback=function (args) {
				var _this=this;
				var eventObj=new AppCommandCallbackEventArgs();
				try {
					var jsonData=JSON.parse(args.eventObjStr);
					this._translateEventObjectInternal(jsonData, eventObj);
					Object.defineProperty(eventObj, 'completed', {
						value: function (completedContext) {
							eventObj.completedContext=completedContext;
							var jsonString=JSON.stringify(eventObj);
							_this._invokeAppCommandCompletedMethod(args.appCommandId, OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess, jsonString);
						},
						enumerable: true
					});
				}
				catch (e) {
					eventObj=null;
				}
				return eventObj;
			};
			AppCommandManager.prototype._translateEventObjectInternal=function (input, output) {
				for (var key in input) {
					if (!input.hasOwnProperty(key))
						continue;
					var inputChild=input[key];
					if (typeof inputChild=="object" && inputChild !=null) {
						OSF.OUtil.defineEnumerableProperty(output, key, {
							value: {}
						});
						this._translateEventObjectInternal(inputChild, output[key]);
					}
					else {
						Object.defineProperty(output, key, {
							value: inputChild,
							enumerable: true,
							writable: true
						});
					}
				}
			};
			AppCommandManager.prototype._constructObjectByTemplate=function (template, input) {
				var output={};
				if (!template || !input)
					return output;
				for (var key in template) {
					if (template.hasOwnProperty(key)) {
						output[key]=null;
						if (input[key] !=null) {
							var templateChild=template[key];
							var inputChild=input[key];
							var inputChildType=typeof inputChild;
							if (typeof templateChild=="object" && templateChild !=null) {
								output[key]=this._constructObjectByTemplate(templateChild, inputChild);
							}
							else if (inputChildType=="number" || inputChildType=="string" || inputChildType=="boolean") {
								output[key]=inputChild;
							}
						}
					}
				}
				return output;
			};
			AppCommandManager.instance=function () {
				if (AppCommandManager._instance==null) {
					AppCommandManager._instance=new AppCommandManager();
				}
				return AppCommandManager._instance;
			};
			AppCommandManager._instance=null;
			return AppCommandManager;
		})();
		AppCommand.AppCommandManager=AppCommandManager;
		var AppCommandInvokedEventArgs=(function () {
			function AppCommandInvokedEventArgs(appCommandId, callbackName, eventObjStr) {
				this.type=Microsoft.Office.WebExtension.EventType.AppCommandInvoked;
				this.appCommandId=appCommandId;
				this.callbackName=callbackName;
				this.eventObjStr=eventObjStr;
			}
			AppCommandInvokedEventArgs.create=function (eventProperties) {
				return new AppCommandInvokedEventArgs(eventProperties[AppCommand.AppCommandInvokedEventEnums.AppCommandId], eventProperties[AppCommand.AppCommandInvokedEventEnums.CallbackName], eventProperties[AppCommand.AppCommandInvokedEventEnums.EventObjStr]);
			};
			return AppCommandInvokedEventArgs;
		})();
		AppCommand.AppCommandInvokedEventArgs=AppCommandInvokedEventArgs;
		var AppCommandCallbackEventArgs=(function () {
			function AppCommandCallbackEventArgs() {
			}
			return AppCommandCallbackEventArgs;
		})();
		AppCommand.AppCommandCallbackEventArgs=AppCommandCallbackEventArgs;
		AppCommand.AppCommandInvokedEventEnums={
			AppCommandId: "appCommandId",
			CallbackName: "callbackName",
			EventObjStr: "eventObjStr"
		};
	})(AppCommand=OfficeExt.AppCommand || (OfficeExt.AppCommand={}));
})(OfficeExt || (OfficeExt={}));
OfficeExt.AppCommand.AppCommandManager.initializeOsfDda();
var OfficeExt;
(function (OfficeExt) {
	var AppCommand;
	(function (AppCommand) {
		function registerDdaFacade() {
			if (OSF.DDA.SafeArray) {
				var parameterMap=OSF.DDA.SafeArray.Delegate.ParameterMap;
				parameterMap.define({
					type: OSF.DDA.MethodDispId.dispidAppCommandInvocationCompletedMethod,
					toHost: [
						{ name: Microsoft.Office.WebExtension.Parameters.Id, value: 0 },
						{ name: Microsoft.Office.WebExtension.Parameters.Status, value: 1 },
						{ name: Microsoft.Office.WebExtension.Parameters.AppCommandInvocationCompletedData, value: 2 }
					]
				});
				parameterMap.define({
					type: OSF.DDA.EventDispId.dispidAppCommandInvokedEvent,
					fromHost: [
						{ name: OSF.DDA.EventDescriptors.AppCommandInvokedEvent, value: parameterMap.self }
					],
					isComplexType: true
				});
				parameterMap.define({
					type: OSF.DDA.EventDescriptors.AppCommandInvokedEvent,
					fromHost: [
						{ name: OfficeExt.AppCommand.AppCommandInvokedEventEnums.AppCommandId, value: 0 },
						{ name: OfficeExt.AppCommand.AppCommandInvokedEventEnums.CallbackName, value: 1 },
						{ name: OfficeExt.AppCommand.AppCommandInvokedEventEnums.EventObjStr, value: 2 },
					],
					isComplexType: true
				});
			}
		}
		AppCommand.registerDdaFacade=registerDdaFacade;
	})(AppCommand=OfficeExt.AppCommand || (OfficeExt.AppCommand={}));
})(OfficeExt || (OfficeExt={}));
OSF.DDA.AsyncMethodNames.addNames({ GetAccessTokenAsync: "getAccessTokenAsync" });
OSF.DDA.Auth=function OSF_DDA_Auth() {
};
OSF.DDA.AsyncMethodCalls.define({
	method: OSF.DDA.AsyncMethodNames.GetAccessTokenAsync,
	requiredArguments: [],
	supportedOptions: [
		{
			name: Microsoft.Office.WebExtension.Parameters.ForceConsent,
			value: {
				"types": ["boolean"],
				"defaultValue": false
			}
		},
		{
			name: Microsoft.Office.WebExtension.Parameters.ForceAddAccount,
			value: {
				"types": ["boolean"],
				"defaultValue": false
			}
		},
		{
			name: Microsoft.Office.WebExtension.Parameters.AuthChallenge,
			value: {
				"types": ["string"],
				"defaultValue": ""
			}
		}
	],
	onSucceeded: function (dataDescriptor, caller, callArgs) {
		var data=dataDescriptor[Microsoft.Office.WebExtension.Parameters.Data];
		return data;
	}
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.MethodDispId.dispidGetAccessTokenMethod,
	toHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.ForceConsent, value: 0 },
		{ name: Microsoft.Office.WebExtension.Parameters.ForceAddAccount, value: 1 },
		{ name: Microsoft.Office.WebExtension.Parameters.AuthChallenge, value: 2 }
	],
	fromHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.Data, value: OSF.DDA.SafeArray.Delegate.ParameterMap.self }
	]
});
OSF.DDA.AsyncMethodNames.addNames({
	OpenBrowserWindow: "openBrowserWindow"
});
OSF.DDA.OpenBrowser=function OSF_DDA_OpenBrowser() {
};
OSF.DDA.AsyncMethodCalls.define({
	method: OSF.DDA.AsyncMethodNames.OpenBrowserWindow,
	requiredArguments: [
		{
			"name": Microsoft.Office.WebExtension.Parameters.Url,
			"types": ["string"]
		}
	],
	supportedOptions: [
		{
			name: Microsoft.Office.WebExtension.Parameters.Reserved,
			value: {
				"types": ["number"],
				"defaultValue": 0
			}
		}
	],
	privateStateCallbacks: []
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.MethodDispId.dispidOpenBrowserWindow,
	toHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.Reserved, value: 0 },
		{ name: Microsoft.Office.WebExtension.Parameters.Url, value: 1 }
	]
});
OSF.DDA.AsyncMethodNames.addNames({
	ExecuteFeature: "executeFeatureAsync",
	QueryFeature: "queryFeatureAsync"
});
OSF.OUtil.augmentList(OSF.DDA.PropertyDescriptors, {
	FeatureProperties: "FeatureProperties",
	TcidEnabled: "TcidEnabled",
	TcidVisible: "TcidVisible"
});
OSF.DDA.ExecuteFeature=function OSF_DDA_ExecuteFeature() {
};
OSF.DDA.QueryFeature=function OSF_DDA_QueryFeature() {
};
OSF.DDA.AsyncMethodCalls.define({
	method: OSF.DDA.AsyncMethodNames.ExecuteFeature,
	requiredArguments: [
		{
			"name": Microsoft.Office.WebExtension.Parameters.Tcid,
			"types": ["number"]
		}
	],
	privateStateCallbacks: []
});
OSF.DDA.AsyncMethodCalls.define({
	method: OSF.DDA.AsyncMethodNames.QueryFeature,
	requiredArguments: [
		{
			"name": Microsoft.Office.WebExtension.Parameters.Tcid,
			"types": ["number"]
		}
	],
	privateStateCallbacks: []
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.PropertyDescriptors.FeatureProperties,
	fromHost: [
		{ name: OSF.DDA.PropertyDescriptors.TcidEnabled, value: 0 },
		{ name: OSF.DDA.PropertyDescriptors.TcidVisible, value: 1 }
	],
	isComplexType: true
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.MethodDispId.dispidExecuteFeature,
	toHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.Tcid, value: 0 }
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.MethodDispId.dispidQueryFeature,
	fromHost: [
		{ name: OSF.DDA.PropertyDescriptors.FeatureProperties, value: OSF.DDA.SafeArray.Delegate.ParameterMap.self }
	],
	toHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.Tcid, value: 0 }
	]
});
OSF.DDA.ExcelDocument=function OSF_DDA_ExcelDocument(officeAppContext, settings) {
	var bf=new OSF.DDA.BindingFacade(this);
	OSF.DDA.DispIdHost.addAsyncMethods(bf, [OSF.DDA.AsyncMethodNames.AddFromPromptAsync]);
	OSF.DDA.DispIdHost.addAsyncMethods(this, [OSF.DDA.AsyncMethodNames.GoToByIdAsync]);
	OSF.DDA.DispIdHost.addAsyncMethods(this, [OSF.DDA.AsyncMethodNames.GetDocumentCopyAsync]);
	OSF.DDA.DispIdHost.addAsyncMethods(this, [OSF.DDA.AsyncMethodNames.GetFilePropertiesAsync]);
	OSF.DDA.ExcelDocument.uber.constructor.call(this, officeAppContext, bf, settings);
	OSF.OUtil.finalizeProperties(this);
};
OSF.OUtil.extend(OSF.DDA.ExcelDocument, OSF.DDA.JsomDocument);
OSF.InitializationHelper.prototype.prepareRightAfterWebExtensionInitialize=function OSF_InitializationHelper$prepareRightAfterWebExtensionInitialize() {
	var appCommandHandler=OfficeExt.AppCommand.AppCommandManager.instance();
	appCommandHandler.initializeAndChangeOnce();
};
OSF.InitializationHelper.prototype.loadAppSpecificScriptAndCreateOM=function OSF_InitializationHelper$loadAppSpecificScriptAndCreateOM(appContext, appReady, basePath) {
	OSF.DDA.ErrorCodeManager.initializeErrorMessages(Strings.OfficeOM);
	appContext.doc=new OSF.DDA.ExcelDocument(appContext, this._initializeSettings(true));
	OSF.DDA.DispIdHost.addAsyncMethods(OSF.DDA.RichApi, [OSF.DDA.AsyncMethodNames.ExecuteRichApiRequestAsync]);
	OSF.DDA.RichApi.richApiMessageManager=new OfficeExt.RichApiMessageManager();
	appReady();
};
(function () {
	OSF.DDA.AsyncMethodCalls.define({
		method: OSF.DDA.AsyncMethodNames.SetSelectedDataAsync,
		requiredArguments: [
			{
				"name": Microsoft.Office.WebExtension.Parameters.Data,
				"types": ["string", "object", "number", "boolean"]
			}
		],
		supportedOptions: [{
				name: Microsoft.Office.WebExtension.Parameters.CoercionType,
				value: {
					"enum": Microsoft.Office.WebExtension.CoercionType,
					"calculate": function (requiredArgs) {
						return OSF.DDA.DataCoercion.determineCoercionType(requiredArgs[Microsoft.Office.WebExtension.Parameters.Data]);
					}
				}
			},
			{
				name: Microsoft.Office.WebExtension.Parameters.CellFormat,
				value: {
					"types": ["number", "object"],
					"defaultValue": []
				}
			},
			{
				name: Microsoft.Office.WebExtension.Parameters.TableOptions,
				value: {
					"types": ["number", "object"],
					"defaultValue": []
				}
			},
			{
				name: Microsoft.Office.WebExtension.Parameters.ImageWidth,
				value: {
					"types": ["number", "boolean"],
					"defaultValue": false
				}
			},
			{
				name: Microsoft.Office.WebExtension.Parameters.ImageHeight,
				value: {
					"types": ["number", "boolean"],
					"defaultValue": false
				}
			}
		],
		privateStateCallbacks: []
	});
})();
var __extends=(this && this.__extends) || (function () {
	var extendStatics=function (d, b) {
		extendStatics=Object.setPrototypeOf ||
			({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__=b; }) ||
			function (d, b) { for (var p in b)
				if (b.hasOwnProperty(p))
					d[p]=b[p]; };
		return extendStatics(d, b);
	};
	return function (d, b) {
		extendStatics(d, b);
		function __() { this.constructor=d; }
		d.prototype=b===null ? Object.create(b) : (__.prototype=b.prototype, new __());
	};
})();
var OfficeExtension;
(function (OfficeExtension) {
	var _Internal;
	(function (_Internal) {
		_Internal.OfficeRequire=function () {
			return null;
		}();
	})(_Internal=OfficeExtension._Internal || (OfficeExtension._Internal={}));
	(function (_Internal) {
		var PromiseImpl;
		(function (PromiseImpl) {
			function Init() {
				return (function () {
					"use strict";
					function lib$es6$promise$utils$$objectOrFunction(x) {
						return typeof x==='function' || (typeof x==='object' && x !==null);
					}
					function lib$es6$promise$utils$$isFunction(x) {
						return typeof x==='function';
					}
					function lib$es6$promise$utils$$isMaybeThenable(x) {
						return typeof x==='object' && x !==null;
					}
					var lib$es6$promise$utils$$_isArray;
					if (!Array.isArray) {
						lib$es6$promise$utils$$_isArray=function (x) {
							return Object.prototype.toString.call(x)==='[object Array]';
						};
					}
					else {
						lib$es6$promise$utils$$_isArray=Array.isArray;
					}
					var lib$es6$promise$utils$$isArray=lib$es6$promise$utils$$_isArray;
					var lib$es6$promise$asap$$len=0;
					var lib$es6$promise$asap$$toString={}.toString;
					var lib$es6$promise$asap$$vertxNext;
					var lib$es6$promise$asap$$customSchedulerFn;
					var lib$es6$promise$asap$$asap=function asap(callback, arg) {
						lib$es6$promise$asap$$queue[lib$es6$promise$asap$$len]=callback;
						lib$es6$promise$asap$$queue[lib$es6$promise$asap$$len+1]=arg;
						lib$es6$promise$asap$$len+=2;
						if (lib$es6$promise$asap$$len===2) {
							if (lib$es6$promise$asap$$customSchedulerFn) {
								lib$es6$promise$asap$$customSchedulerFn(lib$es6$promise$asap$$flush);
							}
							else {
								lib$es6$promise$asap$$scheduleFlush();
							}
						}
					};
					function lib$es6$promise$asap$$setScheduler(scheduleFn) {
						lib$es6$promise$asap$$customSchedulerFn=scheduleFn;
					}
					function lib$es6$promise$asap$$setAsap(asapFn) {
						lib$es6$promise$asap$$asap=asapFn;
					}
					var lib$es6$promise$asap$$browserWindow=(typeof window !=='undefined') ? window : undefined;
					var lib$es6$promise$asap$$browserGlobal=lib$es6$promise$asap$$browserWindow || {};
					var lib$es6$promise$asap$$BrowserMutationObserver=lib$es6$promise$asap$$browserGlobal.MutationObserver || lib$es6$promise$asap$$browserGlobal.WebKitMutationObserver;
					var lib$es6$promise$asap$$isNode=typeof process !=='undefined' && {}.toString.call(process)==='[object process]';
					var lib$es6$promise$asap$$isWorker=typeof Uint8ClampedArray !=='undefined' &&
						typeof importScripts !=='undefined' &&
						typeof MessageChannel !=='undefined';
					function lib$es6$promise$asap$$useNextTick() {
						var nextTick=process.nextTick;
						var version=process.versions.node.match(/^(?:(\d+)\.)?(?:(\d+)\.)?(\*|\d+)$/);
						if (Array.isArray(version) && version[1]==='0' && version[2]==='10') {
							nextTick=window.setImmediate;
						}
						return function () {
							nextTick(lib$es6$promise$asap$$flush);
						};
					}
					function lib$es6$promise$asap$$useVertxTimer() {
						return function () {
							lib$es6$promise$asap$$vertxNext(lib$es6$promise$asap$$flush);
						};
					}
					function lib$es6$promise$asap$$useMutationObserver() {
						var iterations=0;
						var observer=new lib$es6$promise$asap$$BrowserMutationObserver(lib$es6$promise$asap$$flush);
						var node=document.createTextNode('');
						observer.observe(node, { characterData: true });
						return function () {
							node.data=(iterations=++iterations % 2);
						};
					}
					function lib$es6$promise$asap$$useMessageChannel() {
						var channel=new MessageChannel();
						channel.port1.onmessage=lib$es6$promise$asap$$flush;
						return function () {
							channel.port2.postMessage(0);
						};
					}
					function lib$es6$promise$asap$$useSetTimeout() {
						return function () {
							setTimeout(lib$es6$promise$asap$$flush, 1);
						};
					}
					var lib$es6$promise$asap$$queue=new Array(1000);
					function lib$es6$promise$asap$$flush() {
						for (var i=0; i < lib$es6$promise$asap$$len; i+=2) {
							var callback=lib$es6$promise$asap$$queue[i];
							var arg=lib$es6$promise$asap$$queue[i+1];
							callback(arg);
							lib$es6$promise$asap$$queue[i]=undefined;
							lib$es6$promise$asap$$queue[i+1]=undefined;
						}
						lib$es6$promise$asap$$len=0;
					}
					var lib$es6$promise$asap$$scheduleFlush;
					if (lib$es6$promise$asap$$isNode) {
						lib$es6$promise$asap$$scheduleFlush=lib$es6$promise$asap$$useNextTick();
					}
					else if (lib$es6$promise$asap$$BrowserMutationObserver) {
						lib$es6$promise$asap$$scheduleFlush=lib$es6$promise$asap$$useMutationObserver();
					}
					else if (lib$es6$promise$asap$$isWorker) {
						lib$es6$promise$asap$$scheduleFlush=lib$es6$promise$asap$$useMessageChannel();
					}
					else {
						lib$es6$promise$asap$$scheduleFlush=lib$es6$promise$asap$$useSetTimeout();
					}
					function lib$es6$promise$$internal$$noop() { }
					var lib$es6$promise$$internal$$PENDING=void 0;
					var lib$es6$promise$$internal$$FULFILLED=1;
					var lib$es6$promise$$internal$$REJECTED=2;
					var lib$es6$promise$$internal$$GET_THEN_ERROR=new lib$es6$promise$$internal$$ErrorObject();
					function lib$es6$promise$$internal$$selfFullfillment() {
						return new TypeError("You cannot resolve a promise with itself");
					}
					function lib$es6$promise$$internal$$cannotReturnOwn() {
						return new TypeError('A promises callback cannot return that same promise.');
					}
					function lib$es6$promise$$internal$$getThen(promise) {
						try {
							return promise.then;
						}
						catch (error) {
							lib$es6$promise$$internal$$GET_THEN_ERROR.error=error;
							return lib$es6$promise$$internal$$GET_THEN_ERROR;
						}
					}
					function lib$es6$promise$$internal$$tryThen(then, value, fulfillmentHandler, rejectionHandler) {
						try {
							then.call(value, fulfillmentHandler, rejectionHandler);
						}
						catch (e) {
							return e;
						}
					}
					function lib$es6$promise$$internal$$handleForeignThenable(promise, thenable, then) {
						lib$es6$promise$asap$$asap(function (promise) {
							var sealed=false;
							var error=lib$es6$promise$$internal$$tryThen(then, thenable, function (value) {
								if (sealed) {
									return;
								}
								sealed=true;
								if (thenable !==value) {
									lib$es6$promise$$internal$$resolve(promise, value);
								}
								else {
									lib$es6$promise$$internal$$fulfill(promise, value);
								}
							}, function (reason) {
								if (sealed) {
									return;
								}
								sealed=true;
								lib$es6$promise$$internal$$reject(promise, reason);
							}, 'Settle: '+(promise._label || ' unknown promise'));
							if (!sealed && error) {
								sealed=true;
								lib$es6$promise$$internal$$reject(promise, error);
							}
						}, promise);
					}
					function lib$es6$promise$$internal$$handleOwnThenable(promise, thenable) {
						if (thenable._state===lib$es6$promise$$internal$$FULFILLED) {
							lib$es6$promise$$internal$$fulfill(promise, thenable._result);
						}
						else if (thenable._state===lib$es6$promise$$internal$$REJECTED) {
							lib$es6$promise$$internal$$reject(promise, thenable._result);
						}
						else {
							lib$es6$promise$$internal$$subscribe(thenable, undefined, function (value) {
								lib$es6$promise$$internal$$resolve(promise, value);
							}, function (reason) {
								lib$es6$promise$$internal$$reject(promise, reason);
							});
						}
					}
					function lib$es6$promise$$internal$$handleMaybeThenable(promise, maybeThenable) {
						if (maybeThenable.constructor===promise.constructor) {
							lib$es6$promise$$internal$$handleOwnThenable(promise, maybeThenable);
						}
						else {
							var then=lib$es6$promise$$internal$$getThen(maybeThenable);
							if (then===lib$es6$promise$$internal$$GET_THEN_ERROR) {
								lib$es6$promise$$internal$$reject(promise, lib$es6$promise$$internal$$GET_THEN_ERROR.error);
							}
							else if (then===undefined) {
								lib$es6$promise$$internal$$fulfill(promise, maybeThenable);
							}
							else if (lib$es6$promise$utils$$isFunction(then)) {
								lib$es6$promise$$internal$$handleForeignThenable(promise, maybeThenable, then);
							}
							else {
								lib$es6$promise$$internal$$fulfill(promise, maybeThenable);
							}
						}
					}
					function lib$es6$promise$$internal$$resolve(promise, value) {
						if (promise===value) {
							lib$es6$promise$$internal$$reject(promise, lib$es6$promise$$internal$$selfFullfillment());
						}
						else if (lib$es6$promise$utils$$objectOrFunction(value)) {
							lib$es6$promise$$internal$$handleMaybeThenable(promise, value);
						}
						else {
							lib$es6$promise$$internal$$fulfill(promise, value);
						}
					}
					function lib$es6$promise$$internal$$publishRejection(promise) {
						if (promise._onerror) {
							promise._onerror(promise._result);
						}
						lib$es6$promise$$internal$$publish(promise);
					}
					function lib$es6$promise$$internal$$fulfill(promise, value) {
						if (promise._state !==lib$es6$promise$$internal$$PENDING) {
							return;
						}
						promise._result=value;
						promise._state=lib$es6$promise$$internal$$FULFILLED;
						if (promise._subscribers.length !==0) {
							lib$es6$promise$asap$$asap(lib$es6$promise$$internal$$publish, promise);
						}
					}
					function lib$es6$promise$$internal$$reject(promise, reason) {
						if (promise._state !==lib$es6$promise$$internal$$PENDING) {
							return;
						}
						promise._state=lib$es6$promise$$internal$$REJECTED;
						promise._result=reason;
						lib$es6$promise$asap$$asap(lib$es6$promise$$internal$$publishRejection, promise);
					}
					function lib$es6$promise$$internal$$subscribe(parent, child, onFulfillment, onRejection) {
						var subscribers=parent._subscribers;
						var length=subscribers.length;
						parent._onerror=null;
						subscribers[length]=child;
						subscribers[length+lib$es6$promise$$internal$$FULFILLED]=onFulfillment;
						subscribers[length+lib$es6$promise$$internal$$REJECTED]=onRejection;
						if (length===0 && parent._state) {
							lib$es6$promise$asap$$asap(lib$es6$promise$$internal$$publish, parent);
						}
					}
					function lib$es6$promise$$internal$$publish(promise) {
						var subscribers=promise._subscribers;
						var settled=promise._state;
						if (subscribers.length===0) {
							return;
						}
						var child, callback, detail=promise._result;
						for (var i=0; i < subscribers.length; i+=3) {
							child=subscribers[i];
							callback=subscribers[i+settled];
							if (child) {
								lib$es6$promise$$internal$$invokeCallback(settled, child, callback, detail);
							}
							else {
								callback(detail);
							}
						}
						promise._subscribers.length=0;
					}
					function lib$es6$promise$$internal$$ErrorObject() {
						this.error=null;
					}
					var lib$es6$promise$$internal$$TRY_CATCH_ERROR=new lib$es6$promise$$internal$$ErrorObject();
					function lib$es6$promise$$internal$$tryCatch(callback, detail) {
						try {
							return callback(detail);
						}
						catch (e) {
							lib$es6$promise$$internal$$TRY_CATCH_ERROR.error=e;
							return lib$es6$promise$$internal$$TRY_CATCH_ERROR;
						}
					}
					function lib$es6$promise$$internal$$invokeCallback(settled, promise, callback, detail) {
						var hasCallback=lib$es6$promise$utils$$isFunction(callback), value, error, succeeded, failed;
						if (hasCallback) {
							value=lib$es6$promise$$internal$$tryCatch(callback, detail);
							if (value===lib$es6$promise$$internal$$TRY_CATCH_ERROR) {
								failed=true;
								error=value.error;
								value=null;
							}
							else {
								succeeded=true;
							}
							if (promise===value) {
								lib$es6$promise$$internal$$reject(promise, lib$es6$promise$$internal$$cannotReturnOwn());
								return;
							}
						}
						else {
							value=detail;
							succeeded=true;
						}
						if (promise._state !==lib$es6$promise$$internal$$PENDING) {
						}
						else if (hasCallback && succeeded) {
							lib$es6$promise$$internal$$resolve(promise, value);
						}
						else if (failed) {
							lib$es6$promise$$internal$$reject(promise, error);
						}
						else if (settled===lib$es6$promise$$internal$$FULFILLED) {
							lib$es6$promise$$internal$$fulfill(promise, value);
						}
						else if (settled===lib$es6$promise$$internal$$REJECTED) {
							lib$es6$promise$$internal$$reject(promise, value);
						}
					}
					function lib$es6$promise$$internal$$initializePromise(promise, resolver) {
						try {
							resolver(function resolvePromise(value) {
								lib$es6$promise$$internal$$resolve(promise, value);
							}, function rejectPromise(reason) {
								lib$es6$promise$$internal$$reject(promise, reason);
							});
						}
						catch (e) {
							lib$es6$promise$$internal$$reject(promise, e);
						}
					}
					function lib$es6$promise$enumerator$$Enumerator(Constructor, input) {
						var enumerator=this;
						enumerator._instanceConstructor=Constructor;
						enumerator.promise=new Constructor(lib$es6$promise$$internal$$noop);
						if (enumerator._validateInput(input)) {
							enumerator._input=input;
							enumerator.length=input.length;
							enumerator._remaining=input.length;
							enumerator._init();
							if (enumerator.length===0) {
								lib$es6$promise$$internal$$fulfill(enumerator.promise, enumerator._result);
							}
							else {
								enumerator.length=enumerator.length || 0;
								enumerator._enumerate();
								if (enumerator._remaining===0) {
									lib$es6$promise$$internal$$fulfill(enumerator.promise, enumerator._result);
								}
							}
						}
						else {
							lib$es6$promise$$internal$$reject(enumerator.promise, enumerator._validationError());
						}
					}
					lib$es6$promise$enumerator$$Enumerator.prototype._validateInput=function (input) {
						return lib$es6$promise$utils$$isArray(input);
					};
					lib$es6$promise$enumerator$$Enumerator.prototype._validationError=function () {
						return new _Internal.Error('Array Methods must be provided an Array');
					};
					lib$es6$promise$enumerator$$Enumerator.prototype._init=function () {
						this._result=new Array(this.length);
					};
					var lib$es6$promise$enumerator$$default=lib$es6$promise$enumerator$$Enumerator;
					lib$es6$promise$enumerator$$Enumerator.prototype._enumerate=function () {
						var enumerator=this;
						var length=enumerator.length;
						var promise=enumerator.promise;
						var input=enumerator._input;
						for (var i=0; promise._state===lib$es6$promise$$internal$$PENDING && i < length; i++) {
							enumerator._eachEntry(input[i], i);
						}
					};
					lib$es6$promise$enumerator$$Enumerator.prototype._eachEntry=function (entry, i) {
						var enumerator=this;
						var c=enumerator._instanceConstructor;
						if (lib$es6$promise$utils$$isMaybeThenable(entry)) {
							if (entry.constructor===c && entry._state !==lib$es6$promise$$internal$$PENDING) {
								entry._onerror=null;
								enumerator._settledAt(entry._state, i, entry._result);
							}
							else {
								enumerator._willSettleAt(c.resolve(entry), i);
							}
						}
						else {
							enumerator._remaining--;
							enumerator._result[i]=entry;
						}
					};
					lib$es6$promise$enumerator$$Enumerator.prototype._settledAt=function (state, i, value) {
						var enumerator=this;
						var promise=enumerator.promise;
						if (promise._state===lib$es6$promise$$internal$$PENDING) {
							enumerator._remaining--;
							if (state===lib$es6$promise$$internal$$REJECTED) {
								lib$es6$promise$$internal$$reject(promise, value);
							}
							else {
								enumerator._result[i]=value;
							}
						}
						if (enumerator._remaining===0) {
							lib$es6$promise$$internal$$fulfill(promise, enumerator._result);
						}
					};
					lib$es6$promise$enumerator$$Enumerator.prototype._willSettleAt=function (promise, i) {
						var enumerator=this;
						lib$es6$promise$$internal$$subscribe(promise, undefined, function (value) {
							enumerator._settledAt(lib$es6$promise$$internal$$FULFILLED, i, value);
						}, function (reason) {
							enumerator._settledAt(lib$es6$promise$$internal$$REJECTED, i, reason);
						});
					};
					function lib$es6$promise$promise$all$$all(entries) {
						return new lib$es6$promise$enumerator$$default(this, entries).promise;
					}
					var lib$es6$promise$promise$all$$default=lib$es6$promise$promise$all$$all;
					function lib$es6$promise$promise$race$$race(entries) {
						var Constructor=this;
						var promise=new Constructor(lib$es6$promise$$internal$$noop);
						if (!lib$es6$promise$utils$$isArray(entries)) {
							lib$es6$promise$$internal$$reject(promise, new TypeError('You must pass an array to race.'));
							return promise;
						}
						var length=entries.length;
						function onFulfillment(value) {
							lib$es6$promise$$internal$$resolve(promise, value);
						}
						function onRejection(reason) {
							lib$es6$promise$$internal$$reject(promise, reason);
						}
						for (var i=0; promise._state===lib$es6$promise$$internal$$PENDING && i < length; i++) {
							lib$es6$promise$$internal$$subscribe(Constructor.resolve(entries[i]), undefined, onFulfillment, onRejection);
						}
						return promise;
					}
					var lib$es6$promise$promise$race$$default=lib$es6$promise$promise$race$$race;
					function lib$es6$promise$promise$resolve$$resolve(object) {
						var Constructor=this;
						if (object && typeof object==='object' && object.constructor===Constructor) {
							return object;
						}
						var promise=new Constructor(lib$es6$promise$$internal$$noop);
						lib$es6$promise$$internal$$resolve(promise, object);
						return promise;
					}
					var lib$es6$promise$promise$resolve$$default=lib$es6$promise$promise$resolve$$resolve;
					function lib$es6$promise$promise$reject$$reject(reason) {
						var Constructor=this;
						var promise=new Constructor(lib$es6$promise$$internal$$noop);
						lib$es6$promise$$internal$$reject(promise, reason);
						return promise;
					}
					var lib$es6$promise$promise$reject$$default=lib$es6$promise$promise$reject$$reject;
					var lib$es6$promise$promise$$counter=0;
					function lib$es6$promise$promise$$needsResolver() {
						throw new TypeError('You must pass a resolver function as the first argument to the promise constructor');
					}
					function lib$es6$promise$promise$$needsNew() {
						throw new TypeError("Failed to construct 'Promise': Please use the 'new' operator, this object constructor cannot be called as a function.");
					}
					var lib$es6$promise$promise$$default=lib$es6$promise$promise$$Promise;
					function lib$es6$promise$promise$$Promise(resolver) {
						this._id=lib$es6$promise$promise$$counter++;
						this._state=undefined;
						this._result=undefined;
						this._subscribers=[];
						if (lib$es6$promise$$internal$$noop !==resolver) {
							if (!lib$es6$promise$utils$$isFunction(resolver)) {
								lib$es6$promise$promise$$needsResolver();
							}
							if (!(this instanceof lib$es6$promise$promise$$Promise)) {
								lib$es6$promise$promise$$needsNew();
							}
							lib$es6$promise$$internal$$initializePromise(this, resolver);
						}
					}
					lib$es6$promise$promise$$Promise.all=lib$es6$promise$promise$all$$default;
					lib$es6$promise$promise$$Promise.race=lib$es6$promise$promise$race$$default;
					lib$es6$promise$promise$$Promise.resolve=lib$es6$promise$promise$resolve$$default;
					lib$es6$promise$promise$$Promise.reject=lib$es6$promise$promise$reject$$default;
					lib$es6$promise$promise$$Promise._setScheduler=lib$es6$promise$asap$$setScheduler;
					lib$es6$promise$promise$$Promise._setAsap=lib$es6$promise$asap$$setAsap;
					lib$es6$promise$promise$$Promise._asap=lib$es6$promise$asap$$asap;
					lib$es6$promise$promise$$Promise.prototype={
						constructor: lib$es6$promise$promise$$Promise,
						then: function (onFulfillment, onRejection) {
							var parent=this;
							var state=parent._state;
							if (state===lib$es6$promise$$internal$$FULFILLED && !onFulfillment || state===lib$es6$promise$$internal$$REJECTED && !onRejection) {
								return this;
							}
							var child=new this.constructor(lib$es6$promise$$internal$$noop);
							var result=parent._result;
							if (state) {
								var callback=arguments[state - 1];
								lib$es6$promise$asap$$asap(function () {
									lib$es6$promise$$internal$$invokeCallback(state, child, callback, result);
								});
							}
							else {
								lib$es6$promise$$internal$$subscribe(parent, child, onFulfillment, onRejection);
							}
							return child;
						},
						'catch': function (onRejection) {
							return this.then(null, onRejection);
						}
					};
					return lib$es6$promise$promise$$default;
				}).call(this);
			}
			PromiseImpl.Init=Init;
		})(PromiseImpl=_Internal.PromiseImpl || (_Internal.PromiseImpl={}));
	})(_Internal=OfficeExtension._Internal || (OfficeExtension._Internal={}));
	(function (_Internal) {
		function isEdgeLessThan14() {
			var userAgent=window.navigator.userAgent;
			var versionIdx=userAgent.indexOf("Edge/");
			if (versionIdx >=0) {
				userAgent=userAgent.substring(versionIdx+5, userAgent.length);
				if (userAgent < "14.14393")
					return true;
				else
					return false;
			}
			return false;
		}
		function determinePromise() {
			if (typeof (window)==="undefined" && typeof (Promise)==="function") {
				return Promise;
			}
			if (typeof (window) !=="undefined" && window.Promise) {
				if (isEdgeLessThan14()) {
					return _Internal.PromiseImpl.Init();
				}
				else {
					return window.Promise;
				}
			}
			else {
				return _Internal.PromiseImpl.Init();
			}
		}
		_Internal.OfficePromise=determinePromise();
	})(_Internal=OfficeExtension._Internal || (OfficeExtension._Internal={}));
	var OfficePromise=_Internal.OfficePromise;
	OfficeExtension.Promise=OfficePromise;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension_1) {
	var SessionBase=(function () {
		function SessionBase() {
		}
		SessionBase.prototype._resolveRequestUrlAndHeaderInfo=function () {
			return CoreUtility._createPromiseFromResult(null);
		};
		SessionBase.prototype._createRequestExecutorOrNull=function () {
			return null;
		};
		Object.defineProperty(SessionBase.prototype, "eventRegistration", {
			get: function () {
				return null;
			},
			enumerable: true,
			configurable: true
		});
		return SessionBase;
	}());
	OfficeExtension_1.SessionBase=SessionBase;
	var HttpUtility=(function () {
		function HttpUtility() {
		}
		HttpUtility.setCustomSendRequestFunc=function (func) {
			HttpUtility.s_customSendRequestFunc=func;
		};
		HttpUtility.xhrSendRequestFunc=function (request) {
			return CoreUtility.createPromise(function (resolve, reject) {
				var xhr=new XMLHttpRequest();
				xhr.open(request.method, request.url);
				xhr.onload=function () {
					var resp={
						statusCode: xhr.status,
						headers: CoreUtility._parseHttpResponseHeaders(xhr.getAllResponseHeaders()),
						body: xhr.responseText
					};
					resolve(resp);
				};
				xhr.onerror=function () {
					reject(new _Internal.RuntimeError({
						code: CoreErrorCodes.connectionFailure,
						message: CoreUtility._getResourceString(CoreResourceStrings.connectionFailureWithStatus, xhr.statusText)
					}));
				};
				if (request.headers) {
					for (var key in request.headers) {
						xhr.setRequestHeader(key, request.headers[key]);
					}
				}
				xhr.send(CoreUtility._getRequestBodyText(request));
			});
		};
		HttpUtility.sendRequest=function (request) {
			HttpUtility.validateAndNormalizeRequest(request);
			var func=HttpUtility.s_customSendRequestFunc;
			if (!func) {
				func=HttpUtility.xhrSendRequestFunc;
			}
			return func(request);
		};
		HttpUtility.setCustomSendLocalDocumentRequestFunc=function (func) {
			HttpUtility.s_customSendLocalDocumentRequestFunc=func;
		};
		HttpUtility.sendLocalDocumentRequest=function (request) {
			HttpUtility.validateAndNormalizeRequest(request);
			var func;
			func=HttpUtility.s_customSendLocalDocumentRequestFunc || HttpUtility.officeJsSendLocalDocumentRequestFunc;
			return func(request);
		};
		HttpUtility.officeJsSendLocalDocumentRequestFunc=function (request) {
			request=CoreUtility._validateLocalDocumentRequest(request);
			var requestSafeArray=CoreUtility._buildRequestMessageSafeArray(request);
			return CoreUtility.createPromise(function (resolve, reject) {
				OSF.DDA.RichApi.executeRichApiRequestAsync(requestSafeArray, function (asyncResult) {
					var response;
					if (asyncResult.status=='succeeded') {
						response={
							statusCode: RichApiMessageUtility.getResponseStatusCode(asyncResult),
							headers: RichApiMessageUtility.getResponseHeaders(asyncResult),
							body: RichApiMessageUtility.getResponseBody(asyncResult)
						};
					}
					else {
						response=RichApiMessageUtility.buildHttpResponseFromOfficeJsError(asyncResult.error.code, asyncResult.error.message);
					}
					CoreUtility.log('Response:');
					CoreUtility.log(JSON.stringify(response));
					resolve(response);
				});
			});
		};
		HttpUtility.validateAndNormalizeRequest=function (request) {
			if (CoreUtility.isNullOrUndefined(request)) {
				throw _Internal.RuntimeError._createInvalidArgError({
					argumentName: 'request'
				});
			}
			if (CoreUtility.isNullOrEmptyString(request.method)) {
				request.method='GET';
			}
			request.method=request.method.toUpperCase();
		};
		HttpUtility.logRequest=function (request) {
			if (CoreUtility._logEnabled) {
				CoreUtility.log('---HTTP Request---');
				CoreUtility.log(request.method+' '+request.url);
				if (request.headers) {
					for (var key in request.headers) {
						CoreUtility.log(key+': '+request.headers[key]);
					}
				}
				if (HttpUtility._logBodyEnabled) {
					CoreUtility.log(CoreUtility._getRequestBodyText(request));
				}
			}
		};
		HttpUtility.logResponse=function (response) {
			if (CoreUtility._logEnabled) {
				CoreUtility.log('---HTTP Response---');
				CoreUtility.log(''+response.statusCode);
				if (response.headers) {
					for (var key in response.headers) {
						CoreUtility.log(key+': '+response.headers[key]);
					}
				}
				if (HttpUtility._logBodyEnabled) {
					CoreUtility.log(response.body);
				}
			}
		};
		HttpUtility._logBodyEnabled=false;
		return HttpUtility;
	}());
	OfficeExtension_1.HttpUtility=HttpUtility;
	var HostBridge=(function () {
		function HostBridge(m_bridge) {
			var _this=this;
			this.m_bridge=m_bridge;
			this.m_promiseResolver={};
			this.m_handlers=[];
			this.m_bridge.onMessageFromHost=function (messageText) {
				var message=JSON.parse(messageText);
				if (message.type==3) {
					var genericMessageBody=message.message;
					if (genericMessageBody && genericMessageBody.entries) {
						for (var i=0; i < genericMessageBody.entries.length; i++) {
							var entryObjectOrArray=genericMessageBody.entries[i];
							if (Array.isArray(entryObjectOrArray)) {
								var entry={
									messageCategory: entryObjectOrArray[0],
									messageType: entryObjectOrArray[1],
									targetId: entryObjectOrArray[2],
									message: entryObjectOrArray[3],
									id: entryObjectOrArray[4]
								};
								genericMessageBody.entries[i]=entry;
							}
						}
					}
				}
				_this.dispatchMessage(message);
			};
		}
		HostBridge.init=function (bridge) {
			if (typeof bridge !=='object' || !bridge) {
				return;
			}
			var instance=new HostBridge(bridge);
			HostBridge.s_instance=instance;
			HttpUtility.setCustomSendLocalDocumentRequestFunc(function (request) {
				request=CoreUtility._validateLocalDocumentRequest(request);
				var requestFlags=0;
				if (!CoreUtility.isReadonlyRestRequest(request.method)) {
					requestFlags=1;
				}
				var index=request.url.indexOf('?');
				if (index >=0) {
					var query=request.url.substr(index+1);
					var flagsInQueryString=CoreUtility._parseRequestFlagsFromQueryStringIfAny(query);
					if (flagsInQueryString >=0) {
						requestFlags=flagsInQueryString;
					}
				}
				var bridgeMessage={
					id: HostBridge.nextId(),
					type: 1,
					flags: requestFlags,
					message: request
				};
				return instance.sendMessageToHostAndExpectResponse(bridgeMessage).then(function (bridgeResponse) {
					var responseInfo=bridgeResponse.message;
					return responseInfo;
				});
			});
			for (var i=0; i < HostBridge.s_onInitedHandlers.length; i++) {
				HostBridge.s_onInitedHandlers[i](instance);
			}
		};
		Object.defineProperty(HostBridge, "instance", {
			get: function () {
				return HostBridge.s_instance;
			},
			enumerable: true,
			configurable: true
		});
		HostBridge.prototype.sendMessageToHost=function (message) {
			this.m_bridge.sendMessageToHost(JSON.stringify(message));
		};
		HostBridge.prototype.sendMessageToHostAndExpectResponse=function (message) {
			var _this=this;
			var ret=CoreUtility.createPromise(function (resolve, reject) {
				_this.m_promiseResolver[message.id]=resolve;
			});
			this.m_bridge.sendMessageToHost(JSON.stringify(message));
			return ret;
		};
		HostBridge.prototype.addHostMessageHandler=function (handler) {
			this.m_handlers.push(handler);
		};
		HostBridge.prototype.removeHostMessageHandler=function (handler) {
			var index=this.m_handlers.indexOf(handler);
			if (index >=0) {
				this.m_handlers.splice(index, 1);
			}
		};
		HostBridge.onInited=function (handler) {
			HostBridge.s_onInitedHandlers.push(handler);
			if (HostBridge.s_instance) {
				handler(HostBridge.s_instance);
			}
		};
		HostBridge.prototype.dispatchMessage=function (message) {
			if (typeof message.id==='number') {
				var resolve=this.m_promiseResolver[message.id];
				if (resolve) {
					resolve(message);
					delete this.m_promiseResolver[message.id];
					return;
				}
			}
			for (var i=0; i < this.m_handlers.length; i++) {
				this.m_handlers[i](message);
			}
		};
		HostBridge.nextId=function () {
			return HostBridge.s_nextId++;
		};
		HostBridge.s_onInitedHandlers=[];
		HostBridge.s_nextId=1;
		return HostBridge;
	}());
	OfficeExtension_1.HostBridge=HostBridge;
	if (typeof _richApiNativeBridge==='object' && _richApiNativeBridge) {
		HostBridge.init(_richApiNativeBridge);
	}
	var _Internal;
	(function (_Internal) {
		var RuntimeError=(function (_super) {
			__extends(RuntimeError, _super);
			function RuntimeError(error) {
				var _this=_super.call(this, typeof error==='string' ? error : error.message) || this;
				Object.setPrototypeOf(_this, RuntimeError.prototype);
				_this.name='RichApi.Error';
				if (typeof error==='string') {
					_this.message=error;
				}
				else {
					_this.code=error.code;
					_this.message=error.message;
					_this.traceMessages=error.traceMessages || [];
					_this.innerError=error.innerError || null;
					_this.debugInfo=_this._createDebugInfo(error.debugInfo || {});
				}
				return _this;
			}
			RuntimeError.prototype.toString=function () {
				return this.code+': '+this.message;
			};
			RuntimeError.prototype._createDebugInfo=function (partialDebugInfo) {
				var debugInfo={
					code: this.code,
					message: this.message
				};
				debugInfo.toString=function () {
					return JSON.stringify(this);
				};
				for (var key in partialDebugInfo) {
					debugInfo[key]=partialDebugInfo[key];
				}
				if (this.innerError) {
					if (this.innerError instanceof _Internal.RuntimeError) {
						debugInfo.innerError=this.innerError.debugInfo;
					}
					else {
						debugInfo.innerError=this.innerError;
					}
				}
				return debugInfo;
			};
			RuntimeError._createInvalidArgError=function (error) {
				return new _Internal.RuntimeError({
					code: CoreErrorCodes.invalidArgument,
					message: CoreUtility.isNullOrEmptyString(error.argumentName)
						? CoreUtility._getResourceString(CoreResourceStrings.invalidArgumentGeneric)
						: CoreUtility._getResourceString(CoreResourceStrings.invalidArgument, error.argumentName),
					debugInfo: error.errorLocation ? { errorLocation: error.errorLocation } : {},
					innerError: error.innerError
				});
			};
			return RuntimeError;
		}(Error));
		_Internal.RuntimeError=RuntimeError;
	})(_Internal=OfficeExtension_1._Internal || (OfficeExtension_1._Internal={}));
	OfficeExtension_1.Error=_Internal.RuntimeError;
	var CoreErrorCodes=(function () {
		function CoreErrorCodes() {
		}
		CoreErrorCodes.apiNotFound='ApiNotFound';
		CoreErrorCodes.accessDenied='AccessDenied';
		CoreErrorCodes.generalException='GeneralException';
		CoreErrorCodes.activityLimitReached='ActivityLimitReached';
		CoreErrorCodes.invalidArgument='InvalidArgument';
		CoreErrorCodes.connectionFailure='ConnectionFailure';
		CoreErrorCodes.timeout='Timeout';
		CoreErrorCodes.invalidOrTimedOutSession='InvalidOrTimedOutSession';
		CoreErrorCodes.invalidObjectPath='InvalidObjectPath';
		CoreErrorCodes.invalidRequestContext='InvalidRequestContext';
		CoreErrorCodes.valueNotLoaded='ValueNotLoaded';
		return CoreErrorCodes;
	}());
	OfficeExtension_1.CoreErrorCodes=CoreErrorCodes;
	var CoreResourceStrings=(function () {
		function CoreResourceStrings() {
		}
		CoreResourceStrings.apiNotFoundDetails='ApiNotFoundDetails';
		CoreResourceStrings.connectionFailureWithStatus='ConnectionFailureWithStatus';
		CoreResourceStrings.connectionFailureWithDetails='ConnectionFailureWithDetails';
		CoreResourceStrings.invalidArgument='InvalidArgument';
		CoreResourceStrings.invalidArgumentGeneric='InvalidArgumentGeneric';
		CoreResourceStrings.timeout='Timeout';
		CoreResourceStrings.invalidOrTimedOutSessionMessage='InvalidOrTimedOutSessionMessage';
		CoreResourceStrings.invalidObjectPath='InvalidObjectPath';
		CoreResourceStrings.invalidRequestContext='InvalidRequestContext';
		CoreResourceStrings.valueNotLoaded='ValueNotLoaded';
		return CoreResourceStrings;
	}());
	OfficeExtension_1.CoreResourceStrings=CoreResourceStrings;
	var CoreConstants=(function () {
		function CoreConstants() {
		}
		CoreConstants.flags='flags';
		CoreConstants.sourceLibHeader='SdkVersion';
		CoreConstants.processQuery='ProcessQuery';
		CoreConstants.localDocument='http://document.localhost/';
		CoreConstants.localDocumentApiPrefix='http://document.localhost/_api/';
		return CoreConstants;
	}());
	OfficeExtension_1.CoreConstants=CoreConstants;
	var RichApiMessageUtility=(function () {
		function RichApiMessageUtility() {
		}
		RichApiMessageUtility.buildMessageArrayForIRequestExecutor=function (customData, requestFlags, requestMessage, sourceLibHeaderValue) {
			var requestMessageText=JSON.stringify(requestMessage.Body);
			CoreUtility.log('Request:');
			CoreUtility.log(requestMessageText);
			var headers={};
			headers[CoreConstants.sourceLibHeader]=sourceLibHeaderValue;
			var messageSafearray=RichApiMessageUtility.buildRequestMessageSafeArray(customData, requestFlags, 'POST', CoreConstants.processQuery, headers, requestMessageText);
			return messageSafearray;
		};
		RichApiMessageUtility.buildResponseOnSuccess=function (responseBody, responseHeaders) {
			var response={ ErrorCode: '', ErrorMessage: '', Headers: null, Body: null };
			response.Body=JSON.parse(responseBody);
			response.Headers=responseHeaders;
			return response;
		};
		RichApiMessageUtility.buildResponseOnError=function (errorCode, message) {
			var response={ ErrorCode: '', ErrorMessage: '', Headers: null, Body: null };
			response.ErrorCode=CoreErrorCodes.generalException;
			response.ErrorMessage=message;
			if (errorCode==RichApiMessageUtility.OfficeJsErrorCode_ooeNoCapability) {
				response.ErrorCode=CoreErrorCodes.accessDenied;
			}
			else if (errorCode==RichApiMessageUtility.OfficeJsErrorCode_ooeActivityLimitReached) {
				response.ErrorCode=CoreErrorCodes.activityLimitReached;
			}
			else if (errorCode==RichApiMessageUtility.OfficeJsErrorCode_ooeInvalidOrTimedOutSession) {
				response.ErrorCode=CoreErrorCodes.invalidOrTimedOutSession;
				response.ErrorMessage=CoreUtility._getResourceString(CoreResourceStrings.invalidOrTimedOutSessionMessage);
			}
			return response;
		};
		RichApiMessageUtility.buildHttpResponseFromOfficeJsError=function (errorCode, message) {
			var statusCode=500;
			var errorBody={};
			errorBody['error']={};
			errorBody['error']['code']=CoreErrorCodes.generalException;
			errorBody['error']['message']=message;
			if (errorCode===RichApiMessageUtility.OfficeJsErrorCode_ooeNoCapability) {
				statusCode=403;
				errorBody['error']['code']=CoreErrorCodes.accessDenied;
			}
			else if (errorCode===RichApiMessageUtility.OfficeJsErrorCode_ooeActivityLimitReached) {
				statusCode=429;
				errorBody['error']['code']=CoreErrorCodes.activityLimitReached;
			}
			return { statusCode: statusCode, headers: {}, body: JSON.stringify(errorBody) };
		};
		RichApiMessageUtility.buildRequestMessageSafeArray=function (customData, requestFlags, method, path, headers, body) {
			var headerArray=[];
			if (headers) {
				for (var headerName in headers) {
					headerArray.push(headerName);
					headerArray.push(headers[headerName]);
				}
			}
			var appPermission=0;
			var solutionId='';
			var instanceId='';
			var marketplaceType='';
			return [
				customData,
				method,
				path,
				headerArray,
				body,
				appPermission,
				requestFlags,
				solutionId,
				instanceId,
				marketplaceType
			];
		};
		RichApiMessageUtility.getResponseBody=function (result) {
			return RichApiMessageUtility.getResponseBodyFromSafeArray(result.value.data);
		};
		RichApiMessageUtility.getResponseHeaders=function (result) {
			return RichApiMessageUtility.getResponseHeadersFromSafeArray(result.value.data);
		};
		RichApiMessageUtility.getResponseBodyFromSafeArray=function (data) {
			var ret=data[2];
			if (typeof ret==='string') {
				return ret;
			}
			var arr=ret;
			return arr.join('');
		};
		RichApiMessageUtility.getResponseHeadersFromSafeArray=function (data) {
			var arrayHeader=data[1];
			if (!arrayHeader) {
				return null;
			}
			var headers={};
			for (var i=0; i < arrayHeader.length - 1; i+=2) {
				headers[arrayHeader[i]]=arrayHeader[i+1];
			}
			return headers;
		};
		RichApiMessageUtility.getResponseStatusCode=function (result) {
			return RichApiMessageUtility.getResponseStatusCodeFromSafeArray(result.value.data);
		};
		RichApiMessageUtility.getResponseStatusCodeFromSafeArray=function (data) {
			return data[0];
		};
		RichApiMessageUtility.OfficeJsErrorCode_ooeInvalidOrTimedOutSession=5012;
		RichApiMessageUtility.OfficeJsErrorCode_ooeActivityLimitReached=5102;
		RichApiMessageUtility.OfficeJsErrorCode_ooeNoCapability=7000;
		return RichApiMessageUtility;
	}());
	OfficeExtension_1.RichApiMessageUtility=RichApiMessageUtility;
	(function (_Internal) {
		function getPromiseType() {
			if (typeof Promise !=='undefined') {
				return Promise;
			}
			if (typeof Office !=='undefined') {
				if (Office.Promise) {
					return Office.Promise;
				}
			}
			if (typeof OfficeExtension !=='undefined') {
				if (OfficeExtension.Promise) {
					return OfficeExtension.Promise;
				}
			}
			throw new _Internal.Error('No Promise implementation found');
		}
		_Internal.getPromiseType=getPromiseType;
	})(_Internal=OfficeExtension_1._Internal || (OfficeExtension_1._Internal={}));
	var CoreUtility=(function () {
		function CoreUtility() {
		}
		CoreUtility.log=function (message) {
			if (CoreUtility._logEnabled && typeof console !=='undefined' && console.log) {
				console.log(message);
			}
		};
		CoreUtility.checkArgumentNull=function (value, name) {
			if (CoreUtility.isNullOrUndefined(value)) {
				throw _Internal.RuntimeError._createInvalidArgError({ argumentName: name });
			}
		};
		CoreUtility.isNullOrUndefined=function (value) {
			if (value===null) {
				return true;
			}
			if (typeof value==='undefined') {
				return true;
			}
			return false;
		};
		CoreUtility.isUndefined=function (value) {
			if (typeof value==='undefined') {
				return true;
			}
			return false;
		};
		CoreUtility.isNullOrEmptyString=function (value) {
			if (value===null) {
				return true;
			}
			if (typeof value==='undefined') {
				return true;
			}
			if (value.length==0) {
				return true;
			}
			return false;
		};
		CoreUtility.isPlainJsonObject=function (value) {
			if (CoreUtility.isNullOrUndefined(value)) {
				return false;
			}
			if (typeof value !=='object') {
				return false;
			}
			if (Object.prototype.toString.apply(value) !=='[object Object]') {
				return false;
			}
			var prototype=value;
			do {
				prototype=Object.getPrototypeOf(prototype);
			} while (prototype !==null && Object.getPrototypeOf(prototype) !==null);
			return Object.getPrototypeOf(value)===prototype;
		};
		CoreUtility.trim=function (str) {
			return str.replace(new RegExp('^\\s+|\\s+$', 'g'), '');
		};
		CoreUtility.caseInsensitiveCompareString=function (str1, str2) {
			if (CoreUtility.isNullOrUndefined(str1)) {
				return CoreUtility.isNullOrUndefined(str2);
			}
			else {
				if (CoreUtility.isNullOrUndefined(str2)) {
					return false;
				}
				else {
					return str1.toUpperCase()==str2.toUpperCase();
				}
			}
		};
		CoreUtility.isReadonlyRestRequest=function (method) {
			return CoreUtility.caseInsensitiveCompareString(method, 'GET');
		};
		CoreUtility._getResourceString=function (resourceId, arg) {
			var ret;
			if (typeof window !=='undefined' && window.Strings && window.Strings.OfficeOM) {
				var stringName='L_'+resourceId;
				var stringValue=window.Strings.OfficeOM[stringName];
				if (stringValue) {
					ret=stringValue;
				}
			}
			if (!ret) {
				ret=CoreUtility.s_resourceStringValues[resourceId];
			}
			if (!ret) {
				ret=resourceId;
			}
			if (!CoreUtility.isNullOrUndefined(arg)) {
				if (Array.isArray(arg)) {
					var arrArg=arg;
					ret=CoreUtility._formatString(ret, arrArg);
				}
				else {
					ret=ret.replace('{0}', arg);
				}
			}
			return ret;
		};
		CoreUtility._formatString=function (format, arrArg) {
			return format.replace(/\{\d\}/g, function (v) {
				var position=parseInt(v.substr(1, v.length - 2));
				if (position < arrArg.length) {
					return arrArg[position];
				}
				else {
					throw _Internal.RuntimeError._createInvalidArgError({ argumentName: 'format' });
				}
			});
		};
		Object.defineProperty(CoreUtility, "Promise", {
			get: function () {
				return _Internal.getPromiseType();
			},
			enumerable: true,
			configurable: true
		});
		CoreUtility.createPromise=function (executor) {
			var ret=new CoreUtility.Promise(executor);
			return ret;
		};
		CoreUtility._createPromiseFromResult=function (value) {
			return CoreUtility.createPromise(function (resolve, reject) {
				resolve(value);
			});
		};
		CoreUtility._createPromiseFromException=function (reason) {
			return CoreUtility.createPromise(function (resolve, reject) {
				reject(reason);
			});
		};
		CoreUtility._createTimeoutPromise=function (timeout) {
			return CoreUtility.createPromise(function (resolve, reject) {
				setTimeout(function () {
					resolve(null);
				}, timeout);
			});
		};
		CoreUtility._createInvalidArgError=function (error) {
			return _Internal.RuntimeError._createInvalidArgError(error);
		};
		CoreUtility._isLocalDocumentUrl=function (url) {
			return CoreUtility._getLocalDocumentUrlPrefixLength(url) > 0;
		};
		CoreUtility._getLocalDocumentUrlPrefixLength=function (url) {
			var localDocumentPrefixes=[
				'http://document.localhost',
				'https://document.localhost',
				'//document.localhost'
			];
			var urlLower=url.toLowerCase().trim();
			for (var i=0; i < localDocumentPrefixes.length; i++) {
				if (urlLower===localDocumentPrefixes[i]) {
					return localDocumentPrefixes[i].length;
				}
				else if (urlLower.substr(0, localDocumentPrefixes[i].length+1)===localDocumentPrefixes[i]+'/') {
					return localDocumentPrefixes[i].length+1;
				}
			}
			return 0;
		};
		CoreUtility._validateLocalDocumentRequest=function (request) {
			var index=CoreUtility._getLocalDocumentUrlPrefixLength(request.url);
			if (index <=0) {
				throw _Internal.RuntimeError._createInvalidArgError({
					argumentName: 'request'
				});
			}
			var path=request.url.substr(index);
			var pathLower=path.toLowerCase();
			if (pathLower==='_api') {
				path='';
			}
			else if (pathLower.substr(0, '_api/'.length)==='_api/') {
				path=path.substr('_api/'.length);
			}
			return {
				method: request.method,
				url: path,
				headers: request.headers,
				body: request.body
			};
		};
		CoreUtility._parseRequestFlagsFromQueryStringIfAny=function (queryString) {
			var parts=queryString.split('&');
			for (var i=0; i < parts.length; i++) {
				var keyvalue=parts[i].split('=');
				if (keyvalue[0].toLowerCase()===CoreConstants.flags) {
					var flags=parseInt(keyvalue[1]);
					flags=flags & 255;
					return flags;
				}
			}
			return -1;
		};
		CoreUtility._getRequestBodyText=function (request) {
			var body='';
			if (typeof request.body==='string') {
				body=request.body;
			}
			else if (request.body && typeof request.body==='object') {
				body=JSON.stringify(request.body);
			}
			return body;
		};
		CoreUtility._parseResponseBody=function (response) {
			if (typeof response.body==='string') {
				var bodyText=CoreUtility.trim(response.body);
				return JSON.parse(bodyText);
			}
			else {
				return response.body;
			}
		};
		CoreUtility._buildRequestMessageSafeArray=function (request) {
			var requestFlags=0;
			if (!CoreUtility.isReadonlyRestRequest(request.method)) {
				requestFlags=1;
			}
			if (request.url.substr(0, CoreConstants.processQuery.length).toLowerCase()===				CoreConstants.processQuery.toLowerCase()) {
				var index=request.url.indexOf('?');
				if (index > 0) {
					var queryString=request.url.substr(index+1);
					var flagsInQueryString=CoreUtility._parseRequestFlagsFromQueryStringIfAny(queryString);
					if (flagsInQueryString >=0) {
						requestFlags=flagsInQueryString;
					}
				}
			}
			return RichApiMessageUtility.buildRequestMessageSafeArray('', requestFlags, request.method, request.url, request.headers, CoreUtility._getRequestBodyText(request));
		};
		CoreUtility._parseHttpResponseHeaders=function (allResponseHeaders) {
			var responseHeaders={};
			if (!CoreUtility.isNullOrEmptyString(allResponseHeaders)) {
				var regex=new RegExp('\r?\n');
				var entries=allResponseHeaders.split(regex);
				for (var i=0; i < entries.length; i++) {
					var entry=entries[i];
					if (entry !=null) {
						var index=entry.indexOf(':');
						if (index > 0) {
							var key=entry.substr(0, index);
							var value=entry.substr(index+1);
							key=CoreUtility.trim(key);
							value=CoreUtility.trim(value);
							responseHeaders[key.toUpperCase()]=value;
						}
					}
				}
			}
			return responseHeaders;
		};
		CoreUtility._parseErrorResponse=function (responseInfo) {
			var errorObj=null;
			if (CoreUtility.isPlainJsonObject(responseInfo.body)) {
				errorObj=responseInfo.body;
			}
			else if (!CoreUtility.isNullOrEmptyString(responseInfo.body)) {
				var errorResponseBody=CoreUtility.trim(responseInfo.body);
				try {
					errorObj=JSON.parse(errorResponseBody);
				}
				catch (e) {
					CoreUtility.log('Error when parse '+errorResponseBody);
				}
			}
			var errorMessage;
			var errorCode;
			if (!CoreUtility.isNullOrUndefined(errorObj) && typeof errorObj==='object' && errorObj.error) {
				errorCode=errorObj.error.code;
				errorMessage=CoreUtility._getResourceString(CoreResourceStrings.connectionFailureWithDetails, [
					responseInfo.statusCode.toString(),
					errorObj.error.code,
					errorObj.error.message
				]);
			}
			else {
				errorMessage=CoreUtility._getResourceString(CoreResourceStrings.connectionFailureWithStatus, responseInfo.statusCode.toString());
			}
			if (CoreUtility.isNullOrEmptyString(errorCode)) {
				errorCode=CoreErrorCodes.connectionFailure;
			}
			return { errorCode: errorCode, errorMessage: errorMessage };
		};
		CoreUtility._copyHeaders=function (src, dest) {
			if (src && dest) {
				for (var key in src) {
					dest[key]=src[key];
				}
			}
		};
		CoreUtility.addResourceStringValues=function (values) {
			for (var key in values) {
				CoreUtility.s_resourceStringValues[key]=values[key];
			}
		};
		CoreUtility._logEnabled=false;
		CoreUtility.s_resourceStringValues={
			ApiNotFoundDetails: 'The method or property {0} is part of the {1} requirement set, which is not available in your version of {2}.',
			ConnectionFailureWithStatus: 'The request failed with status code of {0}.',
			ConnectionFailureWithDetails: 'The request failed with status code of {0}, error code {1} and the following error message: {2}',
			InvalidArgument: "The argument '{0}' doesn't work for this situation, is missing, or isn't in the right format.",
			InvalidObjectPath: 'The object path \'{0}\' isn\'t working for what you\'re trying to do. If you\'re using the object across multiple "context.sync" calls and outside the sequential execution of a ".run" batch, please use the "context.trackedObjects.add()" and "context.trackedObjects.remove()" methods to manage the object\'s lifetime.',
			InvalidRequestContext: 'Cannot use the object across different request contexts.',
			Timeout: 'The operation has timed out.',
			ValueNotLoaded: 'The value of the result object has not been loaded yet. Before reading the value property, call "context.sync()" on the associated request context.'
		};
		return CoreUtility;
	}());
	OfficeExtension_1.CoreUtility=CoreUtility;
	OfficeExtension_1._internalConfig={
		showDisposeInfoInDebugInfo: false,
		showInternalApiInDebugInfo: false,
		enableEarlyDispose: true,
		alwaysPolyfillClientObjectUpdateMethod: false,
		alwaysPolyfillClientObjectRetrieveMethod: false,
		enableConcurrentFlag: true,
		enableUndoableFlag: true
	};
	OfficeExtension_1.config={
		extendedErrorLogging: false
	};
	var CommonActionFactory=(function () {
		function CommonActionFactory() {
		}
		CommonActionFactory.createSetPropertyAction=function (context, parent, propertyName, value, flags) {
			CommonUtility.validateObjectPath(parent);
			var actionInfo={
				Id: context._nextId(),
				ActionType: 4,
				Name: propertyName,
				ObjectPathId: parent._objectPath.objectPathInfo.Id,
				ArgumentInfo: {}
			};
			var args=[value];
			var referencedArgumentObjectPaths=CommonUtility.setMethodArguments(context, actionInfo.ArgumentInfo, args);
			CommonUtility.validateReferencedObjectPaths(referencedArgumentObjectPaths);
			var action=new Action(actionInfo, 0, flags);
			action.referencedObjectPath=parent._objectPath;
			action.referencedArgumentObjectPaths=referencedArgumentObjectPaths;
			return parent._addAction(action);
		};
		CommonActionFactory.createQueryAction=function (context, parent, queryOption, resultHandler) {
			CommonUtility.validateObjectPath(parent);
			var actionInfo={
				Id: context._nextId(),
				ActionType: 2,
				Name: '',
				ObjectPathId: parent._objectPath.objectPathInfo.Id,
				QueryInfo: queryOption
			};
			var action=new Action(actionInfo, 1, 4);
			action.referencedObjectPath=parent._objectPath;
			return parent._addAction(action, resultHandler);
		};
		CommonActionFactory.createQueryAsJsonAction=function (context, parent, queryOption, resultHandler) {
			CommonUtility.validateObjectPath(parent);
			var actionInfo={
				Id: context._nextId(),
				ActionType: 7,
				Name: '',
				ObjectPathId: parent._objectPath.objectPathInfo.Id,
				QueryInfo: queryOption
			};
			var action=new Action(actionInfo, 1, 4);
			action.referencedObjectPath=parent._objectPath;
			return parent._addAction(action, resultHandler);
		};
		CommonActionFactory.createUpdateAction=function (context, parent, objectState) {
			CommonUtility.validateObjectPath(parent);
			var actionInfo={
				Id: context._nextId(),
				ActionType: 9,
				Name: '',
				ObjectPathId: parent._objectPath.objectPathInfo.Id,
				ObjectState: objectState
			};
			var action=new Action(actionInfo, 0, 0);
			action.referencedObjectPath=parent._objectPath;
			return parent._addAction(action);
		};
		return CommonActionFactory;
	}());
	OfficeExtension_1.CommonActionFactory=CommonActionFactory;
	var ClientObjectBase=(function () {
		function ClientObjectBase(contextBase, objectPath) {
			this.m_contextBase=contextBase;
			this.m_objectPath=objectPath;
		}
		Object.defineProperty(ClientObjectBase.prototype, "_objectPath", {
			get: function () {
				return this.m_objectPath;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ClientObjectBase.prototype, "_context", {
			get: function () {
				return this.m_contextBase;
			},
			enumerable: true,
			configurable: true
		});
		ClientObjectBase.prototype._addAction=function (action, resultHandler) {
			var _this=this;
			if (resultHandler===void 0) {
				resultHandler=null;
			}
			return CoreUtility.createPromise(function (resolve, reject) {
				_this._context._addServiceApiAction(action, resultHandler, resolve, reject);
			});
		};
		ClientObjectBase.prototype._retrieve=function (option, resultHandler) {
			var shouldPolyfill=OfficeExtension_1._internalConfig.alwaysPolyfillClientObjectRetrieveMethod;
			if (!shouldPolyfill) {
				shouldPolyfill=!CommonUtility.isSetSupported('RichApiRuntime', '1.1');
			}
			var queryOption=ClientRequestContextBase._parseQueryOption(option);
			if (shouldPolyfill) {
				return CommonActionFactory.createQueryAction(this._context, this, queryOption, resultHandler);
			}
			return CommonActionFactory.createQueryAsJsonAction(this._context, this, queryOption, resultHandler);
		};
		ClientObjectBase.prototype._recursivelyUpdate=function (properties) {
			var shouldPolyfill=OfficeExtension_1._internalConfig.alwaysPolyfillClientObjectUpdateMethod;
			if (!shouldPolyfill) {
				shouldPolyfill=!CommonUtility.isSetSupported('RichApiRuntime', '1.2');
			}
			try {
				var scalarPropNames=this[CommonConstants.scalarPropertyNames];
				if (!scalarPropNames) {
					scalarPropNames=[];
				}
				var scalarPropUpdatable=this[CommonConstants.scalarPropertyUpdateable];
				if (!scalarPropUpdatable) {
					scalarPropUpdatable=[];
					for (var i=0; i < scalarPropNames.length; i++) {
						scalarPropUpdatable.push(false);
					}
				}
				var navigationPropNames=this[CommonConstants.navigationPropertyNames];
				if (!navigationPropNames) {
					navigationPropNames=[];
				}
				var scalarProps={};
				var navigationProps={};
				var scalarPropCount=0;
				for (var propName in properties) {
					var index=scalarPropNames.indexOf(propName);
					if (index >=0) {
						if (!scalarPropUpdatable[index]) {
							throw new _Internal.RuntimeError({
								code: CoreErrorCodes.invalidArgument,
								message: CoreUtility._getResourceString(CommonResourceStrings.attemptingToSetReadOnlyProperty, propName),
								debugInfo: {
									errorLocation: propName
								}
							});
						}
						scalarProps[propName]=properties[propName];
++scalarPropCount;
					}
					else if (navigationPropNames.indexOf(propName) >=0) {
						navigationProps[propName]=properties[propName];
					}
					else {
						throw new _Internal.RuntimeError({
							code: CoreErrorCodes.invalidArgument,
							message: CoreUtility._getResourceString(CommonResourceStrings.propertyDoesNotExist, propName),
							debugInfo: {
								errorLocation: propName
							}
						});
					}
				}
				if (scalarPropCount > 0) {
					if (shouldPolyfill) {
						for (var i=0; i < scalarPropNames.length; i++) {
							var propName=scalarPropNames[i];
							var propValue=scalarProps[propName];
							if (!CommonUtility.isUndefined(propValue)) {
								CommonActionFactory.createSetPropertyAction(this._context, this, propName, propValue);
							}
						}
					}
					else {
						CommonActionFactory.createUpdateAction(this._context, this, scalarProps);
					}
				}
				for (var propName in navigationProps) {
					var navigationPropProxy=this[propName];
					var navigationPropValue=navigationProps[propName];
					navigationPropProxy._recursivelyUpdate(navigationPropValue);
				}
			}
			catch (innerError) {
				throw new _Internal.RuntimeError({
					code: CoreErrorCodes.invalidArgument,
					message: CoreUtility._getResourceString(CoreResourceStrings.invalidArgument, 'properties'),
					debugInfo: {
						errorLocation: this._className+'.update'
					},
					innerError: innerError
				});
			}
		};
		return ClientObjectBase;
	}());
	OfficeExtension_1.ClientObjectBase=ClientObjectBase;
	var Action=(function () {
		function Action(actionInfo, operationType, flags) {
			this.m_actionInfo=actionInfo;
			this.m_operationType=operationType;
			this.m_flags=flags;
		}
		Object.defineProperty(Action.prototype, "actionInfo", {
			get: function () {
				return this.m_actionInfo;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Action.prototype, "operationType", {
			get: function () {
				return this.m_operationType;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Action.prototype, "flags", {
			get: function () {
				return this.m_flags;
			},
			enumerable: true,
			configurable: true
		});
		return Action;
	}());
	OfficeExtension_1.Action=Action;
	var ObjectPath=(function () {
		function ObjectPath(objectPathInfo, parentObjectPath, isCollection, isInvalidAfterRequest, operationType, flags) {
			this.m_objectPathInfo=objectPathInfo;
			this.m_parentObjectPath=parentObjectPath;
			this.m_isCollection=isCollection;
			this.m_isInvalidAfterRequest=isInvalidAfterRequest;
			this.m_isValid=true;
			this.m_operationType=operationType;
			this.m_flags=flags;
		}
		Object.defineProperty(ObjectPath.prototype, "objectPathInfo", {
			get: function () {
				return this.m_objectPathInfo;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ObjectPath.prototype, "operationType", {
			get: function () {
				return this.m_operationType;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ObjectPath.prototype, "flags", {
			get: function () {
				return this.m_flags;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ObjectPath.prototype, "isCollection", {
			get: function () {
				return this.m_isCollection;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ObjectPath.prototype, "isInvalidAfterRequest", {
			get: function () {
				return this.m_isInvalidAfterRequest;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ObjectPath.prototype, "parentObjectPath", {
			get: function () {
				return this.m_parentObjectPath;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ObjectPath.prototype, "argumentObjectPaths", {
			get: function () {
				return this.m_argumentObjectPaths;
			},
			set: function (value) {
				this.m_argumentObjectPaths=value;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ObjectPath.prototype, "isValid", {
			get: function () {
				return this.m_isValid;
			},
			set: function (value) {
				this.m_isValid=value;
				if (!value &&
					this.m_objectPathInfo.ObjectPathType===6 &&
					this.m_savedObjectPathInfo) {
					ObjectPath.copyObjectPathInfo(this.m_savedObjectPathInfo.pathInfo, this.m_objectPathInfo);
					this.m_parentObjectPath=this.m_savedObjectPathInfo.parent;
					this.m_isValid=true;
					this.m_savedObjectPathInfo=null;
				}
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ObjectPath.prototype, "originalObjectPathInfo", {
			get: function () {
				return this.m_originalObjectPathInfo;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ObjectPath.prototype, "getByIdMethodName", {
			get: function () {
				return this.m_getByIdMethodName;
			},
			set: function (value) {
				this.m_getByIdMethodName=value;
			},
			enumerable: true,
			configurable: true
		});
		ObjectPath.prototype._updateAsNullObject=function () {
			this.resetForUpdateUsingObjectData();
			this.m_objectPathInfo.ObjectPathType=7;
			this.m_objectPathInfo.Name='';
			this.m_parentObjectPath=null;
		};
		ObjectPath.prototype.saveOriginalObjectPathInfo=function () {
			if (OfficeExtension_1.config.extendedErrorLogging && !this.m_originalObjectPathInfo) {
				this.m_originalObjectPathInfo={};
				ObjectPath.copyObjectPathInfo(this.m_objectPathInfo, this.m_originalObjectPathInfo);
			}
		};
		ObjectPath.prototype.updateUsingObjectData=function (value, clientObject) {
			var referenceId=value[CommonConstants.referenceId];
			if (!CoreUtility.isNullOrEmptyString(referenceId)) {
				if (!this.m_savedObjectPathInfo &&
					!this.isInvalidAfterRequest &&
					ObjectPath.isRestorableObjectPath(this.m_objectPathInfo.ObjectPathType)) {
					var pathInfo={};
					ObjectPath.copyObjectPathInfo(this.m_objectPathInfo, pathInfo);
					this.m_savedObjectPathInfo={
						pathInfo: pathInfo,
						parent: this.m_parentObjectPath
					};
				}
				this.saveOriginalObjectPathInfo();
				this.resetForUpdateUsingObjectData();
				this.m_objectPathInfo.ObjectPathType=6;
				this.m_objectPathInfo.Name=referenceId;
				delete this.m_objectPathInfo.ParentObjectPathId;
				this.m_parentObjectPath=null;
				return;
			}
			if (clientObject) {
				var collectionPropertyPath=clientObject[CommonConstants.collectionPropertyPath];
				if (!CoreUtility.isNullOrEmptyString(collectionPropertyPath) && clientObject.context) {
					var id=CommonUtility.tryGetObjectIdFromLoadOrRetrieveResult(value);
					if (!CoreUtility.isNullOrUndefined(id)) {
						var propNames=collectionPropertyPath.split('.');
						var parent_1=clientObject.context[propNames[0]];
						for (var i=1; i < propNames.length; i++) {
							parent_1=parent_1[propNames[i]];
						}
						this.saveOriginalObjectPathInfo();
						this.resetForUpdateUsingObjectData();
						this.m_parentObjectPath=parent_1._objectPath;
						this.m_objectPathInfo.ParentObjectPathId=this.m_parentObjectPath.objectPathInfo.Id;
						this.m_objectPathInfo.ObjectPathType=5;
						this.m_objectPathInfo.Name='';
						this.m_objectPathInfo.ArgumentInfo.Arguments=[id];
						return;
					}
				}
			}
			var parentIsCollection=this.parentObjectPath && this.parentObjectPath.isCollection;
			var getByIdMethodName=this.getByIdMethodName;
			if (parentIsCollection || !CoreUtility.isNullOrEmptyString(getByIdMethodName)) {
				var id=CommonUtility.tryGetObjectIdFromLoadOrRetrieveResult(value);
				if (!CoreUtility.isNullOrUndefined(id)) {
					this.saveOriginalObjectPathInfo();
					this.resetForUpdateUsingObjectData();
					if (!CoreUtility.isNullOrEmptyString(getByIdMethodName)) {
						this.m_objectPathInfo.ObjectPathType=3;
						this.m_objectPathInfo.Name=getByIdMethodName;
						this.m_getByIdMethodName=null;
					}
					else {
						this.m_objectPathInfo.ObjectPathType=5;
						this.m_objectPathInfo.Name='';
					}
					this.m_objectPathInfo.ArgumentInfo.Arguments=[id];
					return;
				}
			}
		};
		ObjectPath.prototype.resetForUpdateUsingObjectData=function () {
			this.m_isInvalidAfterRequest=false;
			this.m_isValid=true;
			this.m_operationType=1;
			this.m_flags=4;
			this.m_objectPathInfo.ArgumentInfo={};
			this.m_argumentObjectPaths=null;
		};
		ObjectPath.isRestorableObjectPath=function (objectPathType) {
			return (objectPathType===1 ||
				objectPathType===5 ||
				objectPathType===3 ||
				objectPathType===4);
		};
		ObjectPath.copyObjectPathInfo=function (src, dest) {
			dest.Id=src.Id;
			dest.ArgumentInfo=src.ArgumentInfo;
			dest.Name=src.Name;
			dest.ObjectPathType=src.ObjectPathType;
			dest.ParentObjectPathId=src.ParentObjectPathId;
		};
		return ObjectPath;
	}());
	OfficeExtension_1.ObjectPath=ObjectPath;
	var ClientRequestContextBase=(function () {
		function ClientRequestContextBase() {
			this.m_nextId=0;
		}
		ClientRequestContextBase.prototype._nextId=function () {
			return++this.m_nextId;
		};
		ClientRequestContextBase.prototype._addServiceApiAction=function (action, resultHandler, resolve, reject) {
			if (!this.m_serviceApiQueue) {
				this.m_serviceApiQueue=new ServiceApiQueue(this);
			}
			this.m_serviceApiQueue.add(action, resultHandler, resolve, reject);
		};
		ClientRequestContextBase._parseQueryOption=function (option) {
			var queryOption={};
			if (typeof option==='string') {
				var select=option;
				queryOption.Select=CommonUtility._parseSelectExpand(select);
			}
			else if (Array.isArray(option)) {
				queryOption.Select=option;
			}
			else if (typeof option==='object') {
				var loadOption=option;
				if (ClientRequestContextBase.isLoadOption(loadOption)) {
					if (typeof loadOption.select==='string') {
						queryOption.Select=CommonUtility._parseSelectExpand(loadOption.select);
					}
					else if (Array.isArray(loadOption.select)) {
						queryOption.Select=loadOption.select;
					}
					else if (!CommonUtility.isNullOrUndefined(loadOption.select)) {
						throw _Internal.RuntimeError._createInvalidArgError({ argumentName: 'option.select' });
					}
					if (typeof loadOption.expand==='string') {
						queryOption.Expand=CommonUtility._parseSelectExpand(loadOption.expand);
					}
					else if (Array.isArray(loadOption.expand)) {
						queryOption.Expand=loadOption.expand;
					}
					else if (!CommonUtility.isNullOrUndefined(loadOption.expand)) {
						throw _Internal.RuntimeError._createInvalidArgError({ argumentName: 'option.expand' });
					}
					if (typeof loadOption.top==='number') {
						queryOption.Top=loadOption.top;
					}
					else if (!CommonUtility.isNullOrUndefined(loadOption.top)) {
						throw _Internal.RuntimeError._createInvalidArgError({ argumentName: 'option.top' });
					}
					if (typeof loadOption.skip==='number') {
						queryOption.Skip=loadOption.skip;
					}
					else if (!CommonUtility.isNullOrUndefined(loadOption.skip)) {
						throw _Internal.RuntimeError._createInvalidArgError({ argumentName: 'option.skip' });
					}
				}
				else {
					queryOption=ClientRequestContextBase.parseStrictLoadOption(option);
				}
			}
			else if (!CommonUtility.isNullOrUndefined(option)) {
				throw _Internal.RuntimeError._createInvalidArgError({ argumentName: 'option' });
			}
			return queryOption;
		};
		ClientRequestContextBase.isLoadOption=function (loadOption) {
			if (!CommonUtility.isUndefined(loadOption.select) &&
				(typeof loadOption.select==='string' || Array.isArray(loadOption.select)))
				return true;
			if (!CommonUtility.isUndefined(loadOption.expand) &&
				(typeof loadOption.expand==='string' || Array.isArray(loadOption.expand)))
				return true;
			if (!CommonUtility.isUndefined(loadOption.top) && typeof loadOption.top==='number')
				return true;
			if (!CommonUtility.isUndefined(loadOption.skip) && typeof loadOption.skip==='number')
				return true;
			for (var i in loadOption) {
				return false;
			}
			return true;
		};
		ClientRequestContextBase.parseStrictLoadOption=function (option) {
			var ret={ Select: [] };
			ClientRequestContextBase.parseStrictLoadOptionHelper(ret, '', 'option', option);
			return ret;
		};
		ClientRequestContextBase.combineQueryPath=function (pathPrefix, key, separator) {
			if (pathPrefix.length===0) {
				return key;
			}
			else {
				return pathPrefix+separator+key;
			}
		};
		ClientRequestContextBase.parseStrictLoadOptionHelper=function (queryInfo, pathPrefix, argPrefix, option) {
			for (var key in option) {
				var value=option[key];
				if (key==='$all') {
					if (typeof value !=='boolean') {
						throw _Internal.RuntimeError._createInvalidArgError({
							argumentName: ClientRequestContextBase.combineQueryPath(argPrefix, key, '.')
						});
					}
					if (value) {
						queryInfo.Select.push(ClientRequestContextBase.combineQueryPath(pathPrefix, '*', '/'));
					}
				}
				else if (key==='$top') {
					if (typeof value !=='number' || pathPrefix.length > 0) {
						throw _Internal.RuntimeError._createInvalidArgError({
							argumentName: ClientRequestContextBase.combineQueryPath(argPrefix, key, '.')
						});
					}
					queryInfo.Top=value;
				}
				else if (key==='$skip') {
					if (typeof value !=='number' || pathPrefix.length > 0) {
						throw _Internal.RuntimeError._createInvalidArgError({
							argumentName: ClientRequestContextBase.combineQueryPath(argPrefix, key, '.')
						});
					}
					queryInfo.Skip=value;
				}
				else {
					if (typeof value==='boolean') {
						if (value) {
							queryInfo.Select.push(ClientRequestContextBase.combineQueryPath(pathPrefix, key, '/'));
						}
					}
					else if (typeof value==='object') {
						ClientRequestContextBase.parseStrictLoadOptionHelper(queryInfo, ClientRequestContextBase.combineQueryPath(pathPrefix, key, '/'), ClientRequestContextBase.combineQueryPath(argPrefix, key, '.'), value);
					}
					else {
						throw _Internal.RuntimeError._createInvalidArgError({
							argumentName: ClientRequestContextBase.combineQueryPath(argPrefix, key, '.')
						});
					}
				}
			}
		};
		return ClientRequestContextBase;
	}());
	OfficeExtension_1.ClientRequestContextBase=ClientRequestContextBase;
	var InstantiateActionUpdateObjectPathHandler=(function () {
		function InstantiateActionUpdateObjectPathHandler(m_objectPath) {
			this.m_objectPath=m_objectPath;
		}
		InstantiateActionUpdateObjectPathHandler.prototype._handleResult=function (value) {
			if (CoreUtility.isNullOrUndefined(value)) {
				this.m_objectPath._updateAsNullObject();
			}
			else {
				this.m_objectPath.updateUsingObjectData(value, null);
			}
		};
		return InstantiateActionUpdateObjectPathHandler;
	}());
	var ClientRequestBase=(function () {
		function ClientRequestBase(context) {
			this.m_contextBase=context;
			this.m_actions=[];
			this.m_actionResultHandler={};
			this.m_referencedObjectPaths={};
			this.m_instantiatedObjectPaths={};
			this.m_preSyncPromises=[];
		}
		ClientRequestBase.prototype.addAction=function (action) {
			this.m_actions.push(action);
			if (action.actionInfo.ActionType==1) {
				this.m_instantiatedObjectPaths[action.actionInfo.ObjectPathId]=action;
			}
		};
		Object.defineProperty(ClientRequestBase.prototype, "hasActions", {
			get: function () {
				return this.m_actions.length > 0;
			},
			enumerable: true,
			configurable: true
		});
		ClientRequestBase.prototype._getLastAction=function () {
			return this.m_actions[this.m_actions.length - 1];
		};
		ClientRequestBase.prototype.ensureInstantiateObjectPath=function (objectPath) {
			if (objectPath) {
				if (this.m_instantiatedObjectPaths[objectPath.objectPathInfo.Id]) {
					return;
				}
				this.ensureInstantiateObjectPath(objectPath.parentObjectPath);
				this.ensureInstantiateObjectPaths(objectPath.argumentObjectPaths);
				if (!this.m_instantiatedObjectPaths[objectPath.objectPathInfo.Id]) {
					var actionInfo={
						Id: this.m_contextBase._nextId(),
						ActionType: 1,
						Name: '',
						ObjectPathId: objectPath.objectPathInfo.Id
					};
					var instantiateAction=new Action(actionInfo, 1, 4);
					instantiateAction.referencedObjectPath=objectPath;
					this.addReferencedObjectPath(objectPath);
					this.addAction(instantiateAction);
					var resultHandler=new InstantiateActionUpdateObjectPathHandler(objectPath);
					this.addActionResultHandler(instantiateAction, resultHandler);
				}
			}
		};
		ClientRequestBase.prototype.ensureInstantiateObjectPaths=function (objectPaths) {
			if (objectPaths) {
				for (var i=0; i < objectPaths.length; i++) {
					this.ensureInstantiateObjectPath(objectPaths[i]);
				}
			}
		};
		ClientRequestBase.prototype.addReferencedObjectPath=function (objectPath) {
			if (!objectPath || this.m_referencedObjectPaths[objectPath.objectPathInfo.Id]) {
				return;
			}
			if (!objectPath.isValid) {
				throw new _Internal.RuntimeError({
					code: CoreErrorCodes.invalidObjectPath,
					message: CoreUtility._getResourceString(CoreResourceStrings.invalidObjectPath, CommonUtility.getObjectPathExpression(objectPath)),
					debugInfo: {
						errorLocation: CommonUtility.getObjectPathExpression(objectPath)
					}
				});
			}
			while (objectPath) {
				this.m_referencedObjectPaths[objectPath.objectPathInfo.Id]=objectPath;
				if (objectPath.objectPathInfo.ObjectPathType==3) {
					this.addReferencedObjectPaths(objectPath.argumentObjectPaths);
				}
				objectPath=objectPath.parentObjectPath;
			}
		};
		ClientRequestBase.prototype.addReferencedObjectPaths=function (objectPaths) {
			if (objectPaths) {
				for (var i=0; i < objectPaths.length; i++) {
					this.addReferencedObjectPath(objectPaths[i]);
				}
			}
		};
		ClientRequestBase.prototype.addActionResultHandler=function (action, resultHandler) {
			this.m_actionResultHandler[action.actionInfo.Id]=resultHandler;
		};
		ClientRequestBase.prototype.aggregrateRequestFlags=function (requestFlags, operationType, flags) {
			if (operationType===0) {
				requestFlags=requestFlags | 1;
				if ((flags & 2)===0) {
					requestFlags=requestFlags & ~16;
				}
				requestFlags=requestFlags & ~4;
			}
			if (flags & 1) {
				requestFlags=requestFlags | 2;
			}
			if ((flags & 4)===0) {
				requestFlags=requestFlags & ~4;
			}
			return requestFlags;
		};
		ClientRequestBase.prototype.finallyNormalizeFlags=function (requestFlags) {
			if ((requestFlags & 1)===0) {
				requestFlags=requestFlags & ~16;
			}
			if (!OfficeExtension_1._internalConfig.enableConcurrentFlag) {
				requestFlags=requestFlags & ~4;
			}
			if (!OfficeExtension_1._internalConfig.enableUndoableFlag) {
				requestFlags=requestFlags & ~16;
			}
			if (!CommonUtility.isSetSupported('RichApiRuntimeFlag', '1.1')) {
				requestFlags=requestFlags & ~4;
				requestFlags=requestFlags & ~16;
			}
			if (typeof this.m_flagsForTesting==='number') {
				requestFlags=this.m_flagsForTesting;
			}
			return requestFlags;
		};
		ClientRequestBase.prototype.buildRequestMessageBodyAndRequestFlags=function () {
			if (OfficeExtension_1._internalConfig.enableEarlyDispose) {
				ClientRequestBase._calculateLastUsedObjectPathIds(this.m_actions);
			}
			var requestFlags=4 | 16;
			var objectPaths={};
			for (var i in this.m_referencedObjectPaths) {
				requestFlags=this.aggregrateRequestFlags(requestFlags, this.m_referencedObjectPaths[i].operationType, this.m_referencedObjectPaths[i].flags);
				objectPaths[i]=this.m_referencedObjectPaths[i].objectPathInfo;
			}
			var actions=[];
			var hasKeepReference=false;
			for (var index=0; index < this.m_actions.length; index++) {
				var action=this.m_actions[index];
				if (action.actionInfo.ActionType===3 &&
					action.actionInfo.Name===CommonConstants.keepReference) {
					hasKeepReference=true;
				}
				requestFlags=this.aggregrateRequestFlags(requestFlags, action.operationType, action.flags);
				actions.push(action.actionInfo);
			}
			requestFlags=this.finallyNormalizeFlags(requestFlags);
			var body={
				AutoKeepReference: this.m_contextBase._autoCleanup && hasKeepReference,
				Actions: actions,
				ObjectPaths: objectPaths
			};
			return {
				body: body,
				flags: requestFlags
			};
		};
		ClientRequestBase.prototype.processResponse=function (actionResults) {
			if (actionResults) {
				for (var i=0; i < actionResults.length; i++) {
					var actionResult=actionResults[i];
					var handler=this.m_actionResultHandler[actionResult.ActionId];
					if (handler) {
						handler._handleResult(actionResult.Value);
					}
				}
			}
		};
		ClientRequestBase.prototype.invalidatePendingInvalidObjectPaths=function () {
			for (var i in this.m_referencedObjectPaths) {
				if (this.m_referencedObjectPaths[i].isInvalidAfterRequest) {
					this.m_referencedObjectPaths[i].isValid=false;
				}
			}
		};
		ClientRequestBase.prototype._addPreSyncPromise=function (value) {
			this.m_preSyncPromises.push(value);
		};
		Object.defineProperty(ClientRequestBase.prototype, "_preSyncPromises", {
			get: function () {
				return this.m_preSyncPromises;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ClientRequestBase.prototype, "_actions", {
			get: function () {
				return this.m_actions;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ClientRequestBase.prototype, "_objectPaths", {
			get: function () {
				return this.m_referencedObjectPaths;
			},
			enumerable: true,
			configurable: true
		});
		ClientRequestBase.prototype._removeKeepReferenceAction=function (objectPathId) {
			for (var i=this.m_actions.length - 1; i >=0; i--) {
				var actionInfo=this.m_actions[i].actionInfo;
				if (actionInfo.ObjectPathId===objectPathId &&
					actionInfo.ActionType===3 &&
					actionInfo.Name===CommonConstants.keepReference) {
					this.m_actions.splice(i, 1);
					break;
				}
			}
		};
		ClientRequestBase._updateLastUsedActionIdOfObjectPathId=function (lastUsedActionIdOfObjectPathId, objectPath, actionId) {
			while (objectPath) {
				if (lastUsedActionIdOfObjectPathId[objectPath.objectPathInfo.Id]) {
					return;
				}
				lastUsedActionIdOfObjectPathId[objectPath.objectPathInfo.Id]=actionId;
				var argumentObjectPaths=objectPath.argumentObjectPaths;
				if (argumentObjectPaths) {
					var argumentObjectPathsLength=argumentObjectPaths.length;
					for (var i=0; i < argumentObjectPathsLength; i++) {
						ClientRequestBase._updateLastUsedActionIdOfObjectPathId(lastUsedActionIdOfObjectPathId, argumentObjectPaths[i], actionId);
					}
				}
				objectPath=objectPath.parentObjectPath;
			}
		};
		ClientRequestBase._calculateLastUsedObjectPathIds=function (actions) {
			var lastUsedActionIdOfObjectPathId={};
			var actionsLength=actions.length;
			for (var index=actionsLength - 1; index >=0; --index) {
				var action=actions[index];
				var actionId=action.actionInfo.Id;
				if (action.referencedObjectPath) {
					ClientRequestBase._updateLastUsedActionIdOfObjectPathId(lastUsedActionIdOfObjectPathId, action.referencedObjectPath, actionId);
				}
				var referencedObjectPaths=action.referencedArgumentObjectPaths;
				if (referencedObjectPaths) {
					var referencedObjectPathsLength=referencedObjectPaths.length;
					for (var refIndex=0; refIndex < referencedObjectPathsLength; refIndex++) {
						ClientRequestBase._updateLastUsedActionIdOfObjectPathId(lastUsedActionIdOfObjectPathId, referencedObjectPaths[refIndex], actionId);
					}
				}
			}
			var lastUsedObjectPathIdsOfAction={};
			for (var key in lastUsedActionIdOfObjectPathId) {
				var actionId=lastUsedActionIdOfObjectPathId[key];
				var objectPathIds=lastUsedObjectPathIdsOfAction[actionId];
				if (!objectPathIds) {
					objectPathIds=[];
					lastUsedObjectPathIdsOfAction[actionId]=objectPathIds;
				}
				objectPathIds.push(parseInt(key));
			}
			for (var index=0; index < actionsLength; index++) {
				var action=actions[index];
				var lastUsedObjectPathIds=lastUsedObjectPathIdsOfAction[action.actionInfo.Id];
				if (lastUsedObjectPathIds && lastUsedObjectPathIds.length > 0) {
					action.actionInfo.L=lastUsedObjectPathIds;
				}
				else if (action.actionInfo.L) {
					delete action.actionInfo.L;
				}
			}
		};
		return ClientRequestBase;
	}());
	OfficeExtension_1.ClientRequestBase=ClientRequestBase;
	var ClientResult=(function () {
		function ClientResult(m_type) {
			this.m_type=m_type;
		}
		Object.defineProperty(ClientResult.prototype, "value", {
			get: function () {
				if (!this.m_isLoaded) {
					throw new _Internal.RuntimeError({
						code: CoreErrorCodes.valueNotLoaded,
						message: CoreUtility._getResourceString(CoreResourceStrings.valueNotLoaded),
						debugInfo: {
							errorLocation: 'clientResult.value'
						}
					});
				}
				return this.m_value;
			},
			enumerable: true,
			configurable: true
		});
		ClientResult.prototype._handleResult=function (value) {
			this.m_isLoaded=true;
			if (typeof value==='object' && value && value._IsNull) {
				return;
			}
			if (this.m_type===1) {
				this.m_value=CommonUtility.adjustToDateTime(value);
			}
			else {
				this.m_value=value;
			}
		};
		return ClientResult;
	}());
	OfficeExtension_1.ClientResult=ClientResult;
	var ServiceApiQueue=(function () {
		function ServiceApiQueue(m_context) {
			this.m_context=m_context;
			this.m_actions=[];
		}
		ServiceApiQueue.prototype.add=function (action, resultHandler, resolve, reject) {
			var _this=this;
			this.m_actions.push({ action: action, resultHandler: resultHandler, resolve: resolve, reject: reject });
			if (this.m_actions.length===1) {
				setTimeout(function () { return _this.processActions(); }, 0);
			}
		};
		ServiceApiQueue.prototype.processActions=function () {
			var _this=this;
			if (this.m_actions.length===0) {
				return;
			}
			var actions=this.m_actions;
			this.m_actions=[];
			var request=new ClientRequestBase(this.m_context);
			for (var i=0; i < actions.length; i++) {
				var action=actions[i];
				request.ensureInstantiateObjectPath(action.action.referencedObjectPath);
				request.ensureInstantiateObjectPaths(action.action.referencedArgumentObjectPaths);
				request.addAction(action.action);
				request.addReferencedObjectPath(action.action.referencedObjectPath);
				request.addReferencedObjectPaths(action.action.referencedArgumentObjectPaths);
			}
			var _a=request.buildRequestMessageBodyAndRequestFlags(), body=_a.body, flags=_a.flags;
			var requestMessage={
				Url: CoreConstants.localDocumentApiPrefix,
				Headers: null,
				Body: body
			};
			CoreUtility.log('Request:');
			CoreUtility.log(JSON.stringify(body));
			var executor=new HttpRequestExecutor();
			executor
				.executeAsync(this.m_context._customData, flags, requestMessage)
				.then(function (response) {
				_this.processResponse(request, actions, response);
			})["catch"](function (ex) {
				for (var i=0; i < actions.length; i++) {
					var action=actions[i];
					action.reject(ex);
				}
			});
		};
		ServiceApiQueue.prototype.processResponse=function (request, actions, response) {
			var error=this.getErrorFromResponse(response);
			var actionResults=null;
			if (response.Body.Results) {
				actionResults=response.Body.Results;
			}
			else if (response.Body.ProcessedResults && response.Body.ProcessedResults.Results) {
				actionResults=response.Body.ProcessedResults.Results;
			}
			if (!actionResults) {
				actionResults=[];
			}
			this.processActionResults(request, actions, actionResults, error);
		};
		ServiceApiQueue.prototype.getErrorFromResponse=function (response) {
			if (!CoreUtility.isNullOrEmptyString(response.ErrorCode)) {
				return new _Internal.RuntimeError({
					code: response.ErrorCode,
					message: response.ErrorMessage
				});
			}
			if (response.Body && response.Body.Error) {
				return new _Internal.RuntimeError({
					code: response.Body.Error.Code,
					message: response.Body.Error.Message
				});
			}
			return null;
		};
		ServiceApiQueue.prototype.processActionResults=function (request, actions, actionResults, err) {
			request.processResponse(actionResults);
			for (var i=0; i < actions.length; i++) {
				var action=actions[i];
				var actionId=action.action.actionInfo.Id;
				var hasResult=false;
				for (var j=0; j < actionResults.length; j++) {
					if (actionId==actionResults[j].ActionId) {
						var resultValue=actionResults[j].Value;
						if (action.resultHandler) {
							action.resultHandler._handleResult(resultValue);
							resultValue=action.resultHandler.value;
						}
						if (action.resolve) {
							action.resolve(resultValue);
						}
						hasResult=true;
						break;
					}
				}
				if (!hasResult && action.reject) {
					if (err) {
						action.reject(err);
					}
					else {
						action.reject('No response for the action.');
					}
				}
			}
		};
		return ServiceApiQueue;
	}());
	var HttpRequestExecutor=(function () {
		function HttpRequestExecutor() {
		}
		HttpRequestExecutor.prototype.executeAsync=function (customData, requestFlags, requestMessage) {
			var url=requestMessage.Url;
			if (url.charAt(url.length - 1) !='/') {
				url=url+'/';
			}
			url=url+CoreConstants.processQuery;
			url=url+'?'+CoreConstants.flags+'='+requestFlags.toString();
			var requestInfo={
				method: 'POST',
				url: url,
				headers: {},
				body: requestMessage.Body
			};
			requestInfo.headers[CoreConstants.sourceLibHeader]=HttpRequestExecutor.SourceLibHeaderValue;
			requestInfo.headers['CONTENT-TYPE']='application/json';
			if (requestMessage.Headers) {
				for (var key in requestMessage.Headers) {
					requestInfo.headers[key]=requestMessage.Headers[key];
				}
			}
			var sendRequestFunc=CoreUtility._isLocalDocumentUrl(requestInfo.url)
				? HttpUtility.sendLocalDocumentRequest
				: HttpUtility.sendRequest;
			return sendRequestFunc(requestInfo).then(function (responseInfo) {
				var response;
				if (responseInfo.statusCode===200) {
					response={
						ErrorCode: null,
						ErrorMessage: null,
						Headers: responseInfo.headers,
						Body: CoreUtility._parseResponseBody(responseInfo)
					};
				}
				else {
					CoreUtility.log('Error Response:'+responseInfo.body);
					var error=CoreUtility._parseErrorResponse(responseInfo);
					response={
						ErrorCode: error.errorCode,
						ErrorMessage: error.errorMessage,
						Headers: responseInfo.headers,
						Body: null
					};
				}
				return response;
			});
		};
		HttpRequestExecutor.SourceLibHeaderValue='officejs-rest';
		return HttpRequestExecutor;
	}());
	OfficeExtension_1.HttpRequestExecutor=HttpRequestExecutor;
	var CommonConstants=(function (_super) {
		__extends(CommonConstants, _super);
		function CommonConstants() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		CommonConstants.collectionPropertyPath='_collectionPropertyPath';
		CommonConstants.id='Id';
		CommonConstants.idLowerCase='id';
		CommonConstants.idPrivate='_Id';
		CommonConstants.keepReference='_KeepReference';
		CommonConstants.objectPathIdPrivate='_ObjectPathId';
		CommonConstants.referenceId='_ReferenceId';
		CommonConstants.items='_Items';
		CommonConstants.itemsLowerCase='items';
		CommonConstants.scalarPropertyNames='_scalarPropertyNames';
		CommonConstants.navigationPropertyNames='_navigationPropertyNames';
		CommonConstants.scalarPropertyUpdateable='_scalarPropertyUpdateable';
		return CommonConstants;
	}(CoreConstants));
	OfficeExtension_1.CommonConstants=CommonConstants;
	var CommonUtility=(function (_super) {
		__extends(CommonUtility, _super);
		function CommonUtility() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		CommonUtility.validateObjectPath=function (clientObject) {
			var objectPath=clientObject._objectPath;
			while (objectPath) {
				if (!objectPath.isValid) {
					throw new _Internal.RuntimeError({
						code: CoreErrorCodes.invalidObjectPath,
						message: CoreUtility._getResourceString(CoreResourceStrings.invalidObjectPath, CommonUtility.getObjectPathExpression(objectPath)),
						debugInfo: {
							errorLocation: CommonUtility.getObjectPathExpression(objectPath)
						}
					});
				}
				objectPath=objectPath.parentObjectPath;
			}
		};
		CommonUtility.validateReferencedObjectPaths=function (objectPaths) {
			if (objectPaths) {
				for (var i=0; i < objectPaths.length; i++) {
					var objectPath=objectPaths[i];
					while (objectPath) {
						if (!objectPath.isValid) {
							throw new _Internal.RuntimeError({
								code: CoreErrorCodes.invalidObjectPath,
								message: CoreUtility._getResourceString(CoreResourceStrings.invalidObjectPath, CommonUtility.getObjectPathExpression(objectPath))
							});
						}
						objectPath=objectPath.parentObjectPath;
					}
				}
			}
		};
		CommonUtility._toCamelLowerCase=function (name) {
			if (CoreUtility.isNullOrEmptyString(name)) {
				return name;
			}
			var index=0;
			while (index < name.length && name.charCodeAt(index) >=65 && name.charCodeAt(index) <=90) {
				index++;
			}
			if (index < name.length) {
				return name.substr(0, index).toLowerCase()+name.substr(index);
			}
			else {
				return name.toLowerCase();
			}
		};
		CommonUtility.adjustToDateTime=function (value) {
			if (CoreUtility.isNullOrUndefined(value)) {
				return null;
			}
			if (typeof value==='string') {
				return new Date(value);
			}
			if (Array.isArray(value)) {
				var arr=value;
				for (var i=0; i < arr.length; i++) {
					arr[i]=CommonUtility.adjustToDateTime(arr[i]);
				}
				return arr;
			}
			throw CoreUtility._createInvalidArgError({ argumentName: 'date' });
		};
		CommonUtility.tryGetObjectIdFromLoadOrRetrieveResult=function (value) {
			var id=value[CommonConstants.id];
			if (CoreUtility.isNullOrUndefined(id)) {
				id=value[CommonConstants.idLowerCase];
			}
			if (CoreUtility.isNullOrUndefined(id)) {
				id=value[CommonConstants.idPrivate];
			}
			return id;
		};
		CommonUtility.getObjectPathExpression=function (objectPath) {
			var ret='';
			while (objectPath) {
				switch (objectPath.objectPathInfo.ObjectPathType) {
					case 1:
						ret=ret;
						break;
					case 2:
						ret='new()'+(ret.length > 0 ? '.' : '')+ret;
						break;
					case 3:
						ret=CommonUtility.normalizeName(objectPath.objectPathInfo.Name)+'()'+(ret.length > 0 ? '.' : '')+ret;
						break;
					case 4:
						ret=CommonUtility.normalizeName(objectPath.objectPathInfo.Name)+(ret.length > 0 ? '.' : '')+ret;
						break;
					case 5:
						ret='getItem()'+(ret.length > 0 ? '.' : '')+ret;
						break;
					case 6:
						ret='_reference()'+(ret.length > 0 ? '.' : '')+ret;
						break;
				}
				objectPath=objectPath.parentObjectPath;
			}
			return ret;
		};
		CommonUtility.setMethodArguments=function (context, argumentInfo, args) {
			if (CoreUtility.isNullOrUndefined(args)) {
				return null;
			}
			var referencedObjectPaths=new Array();
			var referencedObjectPathIds=new Array();
			var hasOne=CommonUtility.collectObjectPathInfos(context, args, referencedObjectPaths, referencedObjectPathIds);
			argumentInfo.Arguments=args;
			if (hasOne) {
				argumentInfo.ReferencedObjectPathIds=referencedObjectPathIds;
			}
			return referencedObjectPaths;
		};
		CommonUtility.validateContext=function (context, obj) {
			if (context && obj && obj._context !==context) {
				throw new _Internal.RuntimeError({
					code: CoreErrorCodes.invalidRequestContext,
					message: CoreUtility._getResourceString(CoreResourceStrings.invalidRequestContext)
				});
			}
		};
		CommonUtility.isSetSupported=function (apiSetName, apiSetVersion) {
			if (typeof window !=='undefined' &&
				window.Office &&
				window.Office.context &&
				window.Office.context.requirements) {
				return window.Office.context.requirements.isSetSupported(apiSetName, apiSetVersion);
			}
			return true;
		};
		CommonUtility.throwIfApiNotSupported=function (apiFullName, apiSetName, apiSetVersion, hostName) {
			if (!CommonUtility._doApiNotSupportedCheck) {
				return;
			}
			if (!CommonUtility.isSetSupported(apiSetName, apiSetVersion)) {
				var message=CoreUtility._getResourceString(CoreResourceStrings.apiNotFoundDetails, [
					apiFullName,
					apiSetName+' '+apiSetVersion,
					hostName
				]);
				throw new _Internal.RuntimeError({
					code: CoreErrorCodes.apiNotFound,
					message: message,
					debugInfo: { errorLocation: apiFullName }
				});
			}
		};
		CommonUtility._parseSelectExpand=function (select) {
			var args=[];
			if (!CoreUtility.isNullOrEmptyString(select)) {
				var propertyNames=select.split(',');
				for (var i=0; i < propertyNames.length; i++) {
					var propertyName=propertyNames[i];
					propertyName=sanitizeForAnyItemsSlash(propertyName.trim());
					if (propertyName.length > 0) {
						args.push(propertyName);
					}
				}
			}
			return args;
			function sanitizeForAnyItemsSlash(propertyName) {
				var propertyNameLower=propertyName.toLowerCase();
				if (propertyNameLower==='items' || propertyNameLower==='items/') {
					return '*';
				}
				var itemsSlashLength=6;
				var isItemsSlashOrItemsDot=propertyNameLower.substr(0, itemsSlashLength)==='items/' ||
					propertyNameLower.substr(0, itemsSlashLength)==='items.';
				if (isItemsSlashOrItemsDot) {
					propertyName=propertyName.substr(itemsSlashLength);
				}
				return propertyName.replace(new RegExp('[/.]items[/.]', 'gi'), '/');
			}
		};
		CommonUtility.changePropertyNameToCamelLowerCase=function (value) {
			var charCodeUnderscore=95;
			if (Array.isArray(value)) {
				var ret=[];
				for (var i=0; i < value.length; i++) {
					ret.push(this.changePropertyNameToCamelLowerCase(value[i]));
				}
				return ret;
			}
			else if (typeof value==='object' && value !==null) {
				var ret={};
				for (var key in value) {
					var propValue=value[key];
					if (key===CommonConstants.items) {
						ret={};
						ret[CommonConstants.itemsLowerCase]=this.changePropertyNameToCamelLowerCase(propValue);
						break;
					}
					else {
						var propName=CommonUtility._toCamelLowerCase(key);
						ret[propName]=this.changePropertyNameToCamelLowerCase(propValue);
					}
				}
				return ret;
			}
			else {
				return value;
			}
		};
		CommonUtility.purifyJson=function (value) {
			var charCodeUnderscore=95;
			if (Array.isArray(value)) {
				var ret=[];
				for (var i=0; i < value.length; i++) {
					ret.push(this.purifyJson(value[i]));
				}
				return ret;
			}
			else if (typeof value==='object' && value !==null) {
				var ret={};
				for (var key in value) {
					if (key.charCodeAt(0) !==charCodeUnderscore) {
						var propValue=value[key];
						if (typeof propValue==='object' && propValue !==null && Array.isArray(propValue['items'])) {
							propValue=propValue['items'];
						}
						ret[key]=this.purifyJson(propValue);
					}
				}
				return ret;
			}
			else {
				return value;
			}
		};
		CommonUtility.collectObjectPathInfos=function (context, args, referencedObjectPaths, referencedObjectPathIds) {
			var hasOne=false;
			for (var i=0; i < args.length; i++) {
				if (args[i] instanceof ClientObjectBase) {
					var clientObject=args[i];
					CommonUtility.validateContext(context, clientObject);
					args[i]=clientObject._objectPath.objectPathInfo.Id;
					referencedObjectPathIds.push(clientObject._objectPath.objectPathInfo.Id);
					referencedObjectPaths.push(clientObject._objectPath);
					hasOne=true;
				}
				else if (Array.isArray(args[i])) {
					var childArrayObjectPathIds=new Array();
					var childArrayHasOne=CommonUtility.collectObjectPathInfos(context, args[i], referencedObjectPaths, childArrayObjectPathIds);
					if (childArrayHasOne) {
						referencedObjectPathIds.push(childArrayObjectPathIds);
						hasOne=true;
					}
					else {
						referencedObjectPathIds.push(0);
					}
				}
				else if (CoreUtility.isPlainJsonObject(args[i])) {
					referencedObjectPathIds.push(0);
					CommonUtility.replaceClientObjectPropertiesWithObjectPathIds(args[i], referencedObjectPaths);
				}
				else {
					referencedObjectPathIds.push(0);
				}
			}
			return hasOne;
		};
		CommonUtility.replaceClientObjectPropertiesWithObjectPathIds=function (value, referencedObjectPaths) {
			var _a, _b;
			for (var key in value) {
				var propValue=value[key];
				if (propValue instanceof ClientObjectBase) {
					referencedObjectPaths.push(propValue._objectPath);
					value[key]=(_a={}, _a[CommonConstants.objectPathIdPrivate]=propValue._objectPath.objectPathInfo.Id, _a);
				}
				else if (Array.isArray(propValue)) {
					for (var i=0; i < propValue.length; i++) {
						if (propValue[i] instanceof ClientObjectBase) {
							var elem=propValue[i];
							referencedObjectPaths.push(elem._objectPath);
							propValue[i]=(_b={}, _b[CommonConstants.objectPathIdPrivate]=elem._objectPath.objectPathInfo.Id, _b);
						}
						else if (CoreUtility.isPlainJsonObject(propValue[i])) {
							CommonUtility.replaceClientObjectPropertiesWithObjectPathIds(propValue[i], referencedObjectPaths);
						}
					}
				}
				else if (CoreUtility.isPlainJsonObject(propValue)) {
					CommonUtility.replaceClientObjectPropertiesWithObjectPathIds(propValue, referencedObjectPaths);
				}
				else {
				}
			}
		};
		CommonUtility.normalizeName=function (name) {
			return name.substr(0, 1).toLowerCase()+name.substr(1);
		};
		CommonUtility._doApiNotSupportedCheck=false;
		return CommonUtility;
	}(CoreUtility));
	OfficeExtension_1.CommonUtility=CommonUtility;
	var CommonResourceStrings=(function (_super) {
		__extends(CommonResourceStrings, _super);
		function CommonResourceStrings() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		CommonResourceStrings.propertyDoesNotExist='PropertyDoesNotExist';
		CommonResourceStrings.attemptingToSetReadOnlyProperty='AttemptingToSetReadOnlyProperty';
		return CommonResourceStrings;
	}(CoreResourceStrings));
	OfficeExtension_1.CommonResourceStrings=CommonResourceStrings;
	var ClientRetrieveResult=(function (_super) {
		__extends(ClientRetrieveResult, _super);
		function ClientRetrieveResult(m_shouldPolyfill) {
			var _this=_super.call(this) || this;
			_this.m_shouldPolyfill=m_shouldPolyfill;
			return _this;
		}
		ClientRetrieveResult.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (this.m_shouldPolyfill) {
				this.m_value=CommonUtility.changePropertyNameToCamelLowerCase(this.m_value);
			}
			this.m_value=this.removeItemNodes(this.m_value);
		};
		ClientRetrieveResult.prototype.removeItemNodes=function (value) {
			if (typeof value==='object' && value !==null && value[CommonConstants.itemsLowerCase]) {
				value=value[CommonConstants.itemsLowerCase];
			}
			return CommonUtility.purifyJson(value);
		};
		return ClientRetrieveResult;
	}(ClientResult));
	OfficeExtension_1.ClientRetrieveResult=ClientRetrieveResult;
	var OperationalApiHelper=(function () {
		function OperationalApiHelper() {
		}
		OperationalApiHelper.invokeMethod=function (obj, methodName, operationType, args, flags, resultProcessType) {
			if (operationType===void 0) {
				operationType=0;
			}
			if (args===void 0) {
				args=[];
			}
			if (flags===void 0) {
				flags=0;
			}
			if (resultProcessType===void 0) {
				resultProcessType=0;
			}
			return CoreUtility.createPromise(function (resolve, reject) {
				var result=new ClientResult();
				var actionInfo={
					Id: obj._context._nextId(),
					ActionType: 3,
					Name: methodName,
					ObjectPathId: obj._objectPath.objectPathInfo.Id,
					ArgumentInfo: {}
				};
				var referencedArgumentObjectPaths=CommonUtility.setMethodArguments(obj._context, actionInfo.ArgumentInfo, args);
				var action=new Action(actionInfo, operationType, flags);
				action.referencedObjectPath=obj._objectPath;
				action.referencedArgumentObjectPaths=referencedArgumentObjectPaths;
				obj._context._addServiceApiAction(action, result, resolve, reject);
			});
		};
		OperationalApiHelper.invokeRetrieve=function (obj, select) {
			var shouldPolyfill=OfficeExtension_1._internalConfig.alwaysPolyfillClientObjectRetrieveMethod;
			if (!shouldPolyfill) {
				shouldPolyfill=!CommonUtility.isSetSupported('RichApiRuntime', '1.1');
			}
			var option;
			if (typeof select[0]==='object' && select[0].hasOwnProperty('$all')) {
				if (!select[0]['$all']) {
					throw _Internal.RuntimeError._createInvalidArgError({});
				}
				option=select[0];
			}
			else {
				option=OperationalApiHelper._parseSelectOption(select);
			}
			return obj._retrieve(option, new ClientRetrieveResult(shouldPolyfill));
		};
		OperationalApiHelper._parseSelectOption=function (select) {
			if (!select || !select[0]) {
				throw _Internal.RuntimeError._createInvalidArgError({});
			}
			var parsedSelect=select[0] && typeof select[0] !=='string' ? select[0] : select;
			return Array.isArray(parsedSelect) ? parsedSelect : OperationalApiHelper.parseRecursiveSelect(parsedSelect);
		};
		OperationalApiHelper.parseRecursiveSelect=function (select) {
			var deconstruct=function (selectObj) {
				return Object.keys(selectObj).reduce(function (scalars, name) {
					var value=selectObj[name];
					if (typeof value==='object') {
						return scalars.concat(deconstruct(value).map(function (postfix) { return name+"/"+postfix; }));
					}
					if (value) {
						return scalars.concat(name);
					}
					return scalars;
				}, []);
			};
			return deconstruct(select);
		};
		OperationalApiHelper.invokeRecursiveUpdate=function (obj, properties) {
			return CoreUtility.createPromise(function (resolve, reject) {
				obj._recursivelyUpdate(properties);
				var actionInfo={
					Id: obj._context._nextId(),
					ActionType: 5,
					Name: 'Trace',
					ObjectPathId: 0
				};
				var action=new Action(actionInfo, 1, 4);
				obj._context._addServiceApiAction(action, null, resolve, reject);
			});
		};
		OperationalApiHelper.createRootServiceObject=function (type, context) {
			var objectPathInfo={
				Id: context._nextId(),
				ObjectPathType: 1,
				Name: ''
			};
			var objectPath=new ObjectPath(objectPathInfo, null, false, false, 1, 4);
			return new type(context, objectPath);
		};
		OperationalApiHelper.createTopLevelServiceObject=function (type, context, typeName, isCollection, flags) {
			var objectPathInfo={
				Id: context._nextId(),
				ObjectPathType: 2,
				Name: typeName
			};
			var objectPath=new ObjectPath(objectPathInfo, null, isCollection, false, 1, flags | 4);
			return new type(context, objectPath);
		};
		OperationalApiHelper.createPropertyObject=function (type, parent, propertyName, isCollection, flags) {
			var objectPathInfo={
				Id: parent._context._nextId(),
				ObjectPathType: 4,
				Name: propertyName,
				ParentObjectPathId: parent._objectPath.objectPathInfo.Id
			};
			var objectPath=new ObjectPath(objectPathInfo, parent._objectPath, isCollection, false, 1, flags | 4);
			return new type(parent._context, objectPath);
		};
		OperationalApiHelper.createIndexerObject=function (type, parent, args) {
			var objectPathInfo={
				Id: parent._context._nextId(),
				ObjectPathType: 5,
				Name: '',
				ParentObjectPathId: parent._objectPath.objectPathInfo.Id,
				ArgumentInfo: {}
			};
			objectPathInfo.ArgumentInfo.Arguments=args;
			var objectPath=new ObjectPath(objectPathInfo, parent._objectPath, false, false, 1, 4);
			return new type(parent._context, objectPath);
		};
		OperationalApiHelper.createMethodObject=function (type, parent, methodName, operationType, args, isCollection, isInvalidAfterRequest, getByIdMethodName, flags) {
			var objectPathInfo={
				Id: parent._context._nextId(),
				ObjectPathType: 3,
				Name: methodName,
				ParentObjectPathId: parent._objectPath.objectPathInfo.Id,
				ArgumentInfo: {}
			};
			var argumentObjectPaths=CommonUtility.setMethodArguments(parent._context, objectPathInfo.ArgumentInfo, args);
			var objectPath=new ObjectPath(objectPathInfo, parent._objectPath, isCollection, isInvalidAfterRequest, operationType, flags);
			objectPath.argumentObjectPaths=argumentObjectPaths;
			objectPath.getByIdMethodName=getByIdMethodName;
			return new type(parent._context, objectPath);
		};
		OperationalApiHelper.createAndInstantiateMethodObject=function (type, parent, methodName, operationType, args, isCollection, isInvalidAfterRequest, getByIdMethodName, flags) {
			return CoreUtility.createPromise(function (resolve, reject) {
				var objectPathInfo={
					Id: parent._context._nextId(),
					ObjectPathType: 3,
					Name: methodName,
					ParentObjectPathId: parent._objectPath.objectPathInfo.Id,
					ArgumentInfo: {}
				};
				var argumentObjectPaths=CommonUtility.setMethodArguments(parent._context, objectPathInfo.ArgumentInfo, args);
				var objectPath=new ObjectPath(objectPathInfo, parent._objectPath, isCollection, isInvalidAfterRequest, operationType, flags);
				objectPath.argumentObjectPaths=argumentObjectPaths;
				objectPath.getByIdMethodName=getByIdMethodName;
				var result=new ClientResult();
				var actionInfo={
					Id: parent._context._nextId(),
					ActionType: 1,
					Name: '',
					ObjectPathId: objectPath.objectPathInfo.Id,
					QueryInfo: {}
				};
				var action=new Action(actionInfo, 1, 4);
				action.referencedObjectPath=objectPath;
				parent._context._addServiceApiAction(action, result, function () { return resolve(new type(parent._context, objectPath)); }, reject);
			});
		};
		OperationalApiHelper.localDocumentContext=new ClientRequestContextBase();
		return OperationalApiHelper;
	}());
	OfficeExtension_1.OperationalApiHelper=OperationalApiHelper;
	var ErrorCodes=(function (_super) {
		__extends(ErrorCodes, _super);
		function ErrorCodes() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		ErrorCodes.propertyNotLoaded='PropertyNotLoaded';
		ErrorCodes.runMustReturnPromise='RunMustReturnPromise';
		ErrorCodes.cannotRegisterEvent='CannotRegisterEvent';
		ErrorCodes.invalidOrTimedOutSession='InvalidOrTimedOutSession';
		ErrorCodes.cannotUpdateReadOnlyProperty='CannotUpdateReadOnlyProperty';
		return ErrorCodes;
	}(CoreErrorCodes));
	OfficeExtension_1.ErrorCodes=ErrorCodes;
	var TraceMarkerActionResultHandler=(function () {
		function TraceMarkerActionResultHandler(callback) {
			this.m_callback=callback;
		}
		TraceMarkerActionResultHandler.prototype._handleResult=function (value) {
			if (this.m_callback) {
				this.m_callback();
			}
		};
		return TraceMarkerActionResultHandler;
	}());
	var ActionFactory=(function (_super) {
		__extends(ActionFactory, _super);
		function ActionFactory() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		ActionFactory.createMethodAction=function (context, parent, methodName, operationType, args, flags) {
			Utility.validateObjectPath(parent);
			var actionInfo={
				Id: context._nextId(),
				ActionType: 3,
				Name: methodName,
				ObjectPathId: parent._objectPath.objectPathInfo.Id,
				ArgumentInfo: {}
			};
			var referencedArgumentObjectPaths=Utility.setMethodArguments(context, actionInfo.ArgumentInfo, args);
			Utility.validateReferencedObjectPaths(referencedArgumentObjectPaths);
			var action=new Action(actionInfo, operationType, Utility._fixupApiFlags(flags));
			action.referencedObjectPath=parent._objectPath;
			action.referencedArgumentObjectPaths=referencedArgumentObjectPaths;
			parent._addAction(action);
			return action;
		};
		ActionFactory.createRecursiveQueryAction=function (context, parent, query) {
			Utility.validateObjectPath(parent);
			var actionInfo={
				Id: context._nextId(),
				ActionType: 6,
				Name: '',
				ObjectPathId: parent._objectPath.objectPathInfo.Id,
				RecursiveQueryInfo: query
			};
			var action=new Action(actionInfo, 1, 4);
			action.referencedObjectPath=parent._objectPath;
			parent._addAction(action);
			return action;
		};
		ActionFactory.createEnsureUnchangedAction=function (context, parent, objectState) {
			Utility.validateObjectPath(parent);
			var actionInfo={
				Id: context._nextId(),
				ActionType: 8,
				Name: '',
				ObjectPathId: parent._objectPath.objectPathInfo.Id,
				ObjectState: objectState
			};
			var action=new Action(actionInfo, 1, 4);
			action.referencedObjectPath=parent._objectPath;
			parent._addAction(action);
			return action;
		};
		ActionFactory.createInstantiateAction=function (context, obj) {
			Utility.validateObjectPath(obj);
			context._pendingRequest.ensureInstantiateObjectPath(obj._objectPath.parentObjectPath);
			context._pendingRequest.ensureInstantiateObjectPaths(obj._objectPath.argumentObjectPaths);
			var actionInfo={
				Id: context._nextId(),
				ActionType: 1,
				Name: '',
				ObjectPathId: obj._objectPath.objectPathInfo.Id
			};
			var action=new Action(actionInfo, 1, 4);
			action.referencedObjectPath=obj._objectPath;
			obj._addAction(action, new InstantiateActionResultHandler(obj), true);
			return action;
		};
		ActionFactory.createTraceAction=function (context, message, addTraceMessage) {
			var actionInfo={
				Id: context._nextId(),
				ActionType: 5,
				Name: 'Trace',
				ObjectPathId: 0
			};
			var ret=new Action(actionInfo, 1, 4);
			context._pendingRequest.addAction(ret);
			if (addTraceMessage) {
				context._pendingRequest.addTrace(actionInfo.Id, message);
			}
			return ret;
		};
		ActionFactory.createTraceMarkerForCallback=function (context, callback) {
			var action=ActionFactory.createTraceAction(context, null, false);
			context._pendingRequest.addActionResultHandler(action, new TraceMarkerActionResultHandler(callback));
		};
		return ActionFactory;
	}(CommonActionFactory));
	OfficeExtension_1.ActionFactory=ActionFactory;
	var ClientObject=(function (_super) {
		__extends(ClientObject, _super);
		function ClientObject(context, objectPath) {
			var _this=_super.call(this, context, objectPath) || this;
			Utility.checkArgumentNull(context, 'context');
			_this.m_context=context;
			if (_this._objectPath) {
				if (!context._processingResult && context._pendingRequest) {
					ActionFactory.createInstantiateAction(context, _this);
					if (context._autoCleanup && _this._KeepReference) {
						context.trackedObjects._autoAdd(_this);
					}
				}
			}
			return _this;
		}
		Object.defineProperty(ClientObject.prototype, "context", {
			get: function () {
				return this.m_context;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ClientObject.prototype, "isNull", {
			get: function () {
				Utility.throwIfNotLoaded('isNull', this._isNull, null, this._isNull);
				return this._isNull;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ClientObject.prototype, "isNullObject", {
			get: function () {
				Utility.throwIfNotLoaded('isNullObject', this._isNull, null, this._isNull);
				return this._isNull;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ClientObject.prototype, "_isNull", {
			get: function () {
				return this.m_isNull;
			},
			set: function (value) {
				this.m_isNull=value;
				if (value && this._objectPath) {
					this._objectPath._updateAsNullObject();
				}
			},
			enumerable: true,
			configurable: true
		});
		ClientObject.prototype._addAction=function (action, resultHandler, isInstantiationEnsured) {
			if (resultHandler===void 0) {
				resultHandler=null;
			}
			if (!isInstantiationEnsured) {
				this.context._pendingRequest.ensureInstantiateObjectPath(this._objectPath);
				this.context._pendingRequest.ensureInstantiateObjectPaths(action.referencedArgumentObjectPaths);
			}
			this.context._pendingRequest.addAction(action);
			this.context._pendingRequest.addReferencedObjectPath(this._objectPath);
			this.context._pendingRequest.addReferencedObjectPaths(action.referencedArgumentObjectPaths);
			this.context._pendingRequest.addActionResultHandler(action, resultHandler);
			return CoreUtility._createPromiseFromResult(null);
		};
		ClientObject.prototype._handleResult=function (value) {
			this._isNull=Utility.isNullOrUndefined(value);
			this.context.trackedObjects._autoTrackIfNecessaryWhenHandleObjectResultValue(this, value);
		};
		ClientObject.prototype._handleIdResult=function (value) {
			this._isNull=Utility.isNullOrUndefined(value);
			Utility.fixObjectPathIfNecessary(this, value);
			this.context.trackedObjects._autoTrackIfNecessaryWhenHandleObjectResultValue(this, value);
		};
		ClientObject.prototype._handleRetrieveResult=function (value, result) {
			this._handleIdResult(value);
		};
		ClientObject.prototype._recursivelySet=function (input, options, scalarWriteablePropertyNames, objectPropertyNames, notAllowedToBeSetPropertyNames) {
			var isClientObject=input instanceof ClientObject;
			var originalInput=input;
			if (isClientObject) {
				if (Object.getPrototypeOf(this)===Object.getPrototypeOf(input)) {
					input=JSON.parse(JSON.stringify(input));
				}
				else {
					throw _Internal.RuntimeError._createInvalidArgError({
						argumentName: 'properties',
						errorLocation: this._className+'.set'
					});
				}
			}
			try {
				var prop;
				for (var i=0; i < scalarWriteablePropertyNames.length; i++) {
					prop=scalarWriteablePropertyNames[i];
					if (input.hasOwnProperty(prop)) {
						if (typeof input[prop] !=='undefined') {
							this[prop]=input[prop];
						}
					}
				}
				for (var i=0; i < objectPropertyNames.length; i++) {
					prop=objectPropertyNames[i];
					if (input.hasOwnProperty(prop)) {
						if (typeof input[prop] !=='undefined') {
							var dataToPassToSet=isClientObject ? originalInput[prop] : input[prop];
							this[prop].set(dataToPassToSet, options);
						}
					}
				}
				var throwOnReadOnly=!isClientObject;
				if (options && !Utility.isNullOrUndefined(throwOnReadOnly)) {
					throwOnReadOnly=options.throwOnReadOnly;
				}
				for (var i=0; i < notAllowedToBeSetPropertyNames.length; i++) {
					prop=notAllowedToBeSetPropertyNames[i];
					if (input.hasOwnProperty(prop)) {
						if (typeof input[prop] !=='undefined' && throwOnReadOnly) {
							throw new _Internal.RuntimeError({
								code: CoreErrorCodes.invalidArgument,
								message: CoreUtility._getResourceString(ResourceStrings.cannotApplyPropertyThroughSetMethod, prop),
								debugInfo: {
									errorLocation: prop
								}
							});
						}
					}
				}
				for (prop in input) {
					if (scalarWriteablePropertyNames.indexOf(prop) < 0 && objectPropertyNames.indexOf(prop) < 0) {
						var propertyDescriptor=Object.getOwnPropertyDescriptor(Object.getPrototypeOf(this), prop);
						if (!propertyDescriptor) {
							throw new _Internal.RuntimeError({
								code: CoreErrorCodes.invalidArgument,
								message: CoreUtility._getResourceString(CommonResourceStrings.propertyDoesNotExist, prop),
								debugInfo: {
									errorLocation: prop
								}
							});
						}
						if (throwOnReadOnly && !propertyDescriptor.set) {
							throw new _Internal.RuntimeError({
								code: CoreErrorCodes.invalidArgument,
								message: CoreUtility._getResourceString(CommonResourceStrings.attemptingToSetReadOnlyProperty, prop),
								debugInfo: {
									errorLocation: prop
								}
							});
						}
					}
				}
			}
			catch (innerError) {
				throw new _Internal.RuntimeError({
					code: CoreErrorCodes.invalidArgument,
					message: CoreUtility._getResourceString(CoreResourceStrings.invalidArgument, 'properties'),
					debugInfo: {
						errorLocation: this._className+'.set'
					},
					innerError: innerError
				});
			}
		};
		return ClientObject;
	}(ClientObjectBase));
	OfficeExtension_1.ClientObject=ClientObject;
	var HostBridgeRequestExecutor=(function () {
		function HostBridgeRequestExecutor(session) {
			this.m_session=session;
		}
		HostBridgeRequestExecutor.prototype.executeAsync=function (customData, requestFlags, requestMessage) {
			var httpRequestInfo={
				url: CoreConstants.processQuery,
				method: 'POST',
				headers: requestMessage.Headers,
				body: requestMessage.Body
			};
			var message={
				id: HostBridge.nextId(),
				type: 1,
				flags: requestFlags,
				message: httpRequestInfo
			};
			CoreUtility.log(JSON.stringify(message));
			return this.m_session.sendMessageToHost(message).then(function (nativeBridgeResponse) {
				CoreUtility.log('Received response: '+JSON.stringify(nativeBridgeResponse));
				var responseInfo=nativeBridgeResponse.message;
				var response;
				if (responseInfo.statusCode===200) {
					response={
						ErrorCode: null,
						ErrorMessage: null,
						Headers: responseInfo.headers,
						Body: CoreUtility._parseResponseBody(responseInfo)
					};
				}
				else {
					CoreUtility.log('Error Response:'+responseInfo.body);
					var error=CoreUtility._parseErrorResponse(responseInfo);
					response={
						ErrorCode: error.errorCode,
						ErrorMessage: error.errorMessage,
						Headers: responseInfo.headers,
						Body: null
					};
				}
				return response;
			});
		};
		return HostBridgeRequestExecutor;
	}());
	var HostBridgeSession=(function (_super) {
		__extends(HostBridgeSession, _super);
		function HostBridgeSession(m_bridge) {
			var _this=_super.call(this) || this;
			_this.m_bridge=m_bridge;
			_this.m_bridge.addHostMessageHandler(function (message) {
				if (message.type===3) {
					GenericEventRegistration.getGenericEventRegistration()._handleRichApiMessage(message.message);
				}
			});
			return _this;
		}
		HostBridgeSession.getInstanceIfHostBridgeInited=function () {
			if (HostBridge.instance) {
				if (CoreUtility.isNullOrUndefined(HostBridgeSession.s_instance) ||
					HostBridgeSession.s_instance.m_bridge !==HostBridge.instance) {
					HostBridgeSession.s_instance=new HostBridgeSession(HostBridge.instance);
				}
				return HostBridgeSession.s_instance;
			}
			return null;
		};
		HostBridgeSession.prototype._resolveRequestUrlAndHeaderInfo=function () {
			return CoreUtility._createPromiseFromResult(null);
		};
		HostBridgeSession.prototype._createRequestExecutorOrNull=function () {
			CoreUtility.log('NativeBridgeSession::CreateRequestExecutor');
			return new HostBridgeRequestExecutor(this);
		};
		Object.defineProperty(HostBridgeSession.prototype, "eventRegistration", {
			get: function () {
				return GenericEventRegistration.getGenericEventRegistration();
			},
			enumerable: true,
			configurable: true
		});
		HostBridgeSession.prototype.sendMessageToHost=function (message) {
			return this.m_bridge.sendMessageToHostAndExpectResponse(message);
		};
		return HostBridgeSession;
	}(SessionBase));
	var ClientRequestContext=(function (_super) {
		__extends(ClientRequestContext, _super);
		function ClientRequestContext(url) {
			var _this=_super.call(this) || this;
			_this.m_customRequestHeaders={};
			_this.m_batchMode=0;
			_this._onRunFinishedNotifiers=[];
			if (SessionBase._overrideSession) {
				_this.m_requestUrlAndHeaderInfoResolver=SessionBase._overrideSession;
			}
			else {
				if (Utility.isNullOrUndefined(url) || (typeof url==='string' && url.length===0)) {
					url=ClientRequestContext.defaultRequestUrlAndHeaders;
					if (!url) {
						url={ url: CoreConstants.localDocument, headers: {} };
					}
				}
				if (typeof url==='string') {
					_this.m_requestUrlAndHeaderInfo={ url: url, headers: {} };
				}
				else if (ClientRequestContext.isRequestUrlAndHeaderInfoResolver(url)) {
					_this.m_requestUrlAndHeaderInfoResolver=url;
				}
				else if (ClientRequestContext.isRequestUrlAndHeaderInfo(url)) {
					var requestInfo=url;
					_this.m_requestUrlAndHeaderInfo={ url: requestInfo.url, headers: {} };
					CoreUtility._copyHeaders(requestInfo.headers, _this.m_requestUrlAndHeaderInfo.headers);
				}
				else {
					throw _Internal.RuntimeError._createInvalidArgError({ argumentName: 'url' });
				}
			}
			if (!_this.m_requestUrlAndHeaderInfoResolver &&
				_this.m_requestUrlAndHeaderInfo &&
				CoreUtility._isLocalDocumentUrl(_this.m_requestUrlAndHeaderInfo.url) &&
				HostBridgeSession.getInstanceIfHostBridgeInited()) {
				_this.m_requestUrlAndHeaderInfo=null;
				_this.m_requestUrlAndHeaderInfoResolver=HostBridgeSession.getInstanceIfHostBridgeInited();
			}
			if (_this.m_requestUrlAndHeaderInfoResolver instanceof SessionBase) {
				_this.m_session=_this.m_requestUrlAndHeaderInfoResolver;
			}
			_this._processingResult=false;
			_this._customData=Constants.iterativeExecutor;
			_this.sync=_this.sync.bind(_this);
			return _this;
		}
		Object.defineProperty(ClientRequestContext.prototype, "session", {
			get: function () {
				return this.m_session;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ClientRequestContext.prototype, "eventRegistration", {
			get: function () {
				if (this.m_session) {
					return this.m_session.eventRegistration;
				}
				return _Internal.officeJsEventRegistration;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ClientRequestContext.prototype, "_url", {
			get: function () {
				if (this.m_requestUrlAndHeaderInfo) {
					return this.m_requestUrlAndHeaderInfo.url;
				}
				return null;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ClientRequestContext.prototype, "_pendingRequest", {
			get: function () {
				if (this.m_pendingRequest==null) {
					this.m_pendingRequest=new ClientRequest(this);
				}
				return this.m_pendingRequest;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ClientRequestContext.prototype, "debugInfo", {
			get: function () {
				var prettyPrinter=new RequestPrettyPrinter(this._rootObjectPropertyName, this._pendingRequest._objectPaths, this._pendingRequest._actions, OfficeExtension_1._internalConfig.showDisposeInfoInDebugInfo);
				var statements=prettyPrinter.process();
				return { pendingStatements: statements };
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ClientRequestContext.prototype, "trackedObjects", {
			get: function () {
				if (!this.m_trackedObjects) {
					this.m_trackedObjects=new TrackedObjects(this);
				}
				return this.m_trackedObjects;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ClientRequestContext.prototype, "requestHeaders", {
			get: function () {
				return this.m_customRequestHeaders;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ClientRequestContext.prototype, "batchMode", {
			get: function () {
				return this.m_batchMode;
			},
			enumerable: true,
			configurable: true
		});
		ClientRequestContext.prototype.ensureInProgressBatchIfBatchMode=function () {
			if (this.m_batchMode===1 && !this.m_explicitBatchInProgress) {
				throw Utility.createRuntimeError(CoreErrorCodes.generalException, CoreUtility._getResourceString(ResourceStrings.notInsideBatch), null);
			}
		};
		ClientRequestContext.prototype.load=function (clientObj, option) {
			Utility.validateContext(this, clientObj);
			var queryOption=ClientRequestContext._parseQueryOption(option);
			CommonActionFactory.createQueryAction(this, clientObj, queryOption, clientObj);
		};
		ClientRequestContext.prototype.loadRecursive=function (clientObj, options, maxDepth) {
			if (!Utility.isPlainJsonObject(options)) {
				throw _Internal.RuntimeError._createInvalidArgError({ argumentName: 'options' });
			}
			var quries={};
			for (var key in options) {
				quries[key]=ClientRequestContext._parseQueryOption(options[key]);
			}
			var action=ActionFactory.createRecursiveQueryAction(this, clientObj, { Queries: quries, MaxDepth: maxDepth });
			this._pendingRequest.addActionResultHandler(action, clientObj);
		};
		ClientRequestContext.prototype.trace=function (message) {
			ActionFactory.createTraceAction(this, message, true);
		};
		ClientRequestContext.prototype._processOfficeJsErrorResponse=function (officeJsErrorCode, response) { };
		ClientRequestContext.prototype.ensureRequestUrlAndHeaderInfo=function () {
			var _this=this;
			return Utility._createPromiseFromResult(null).then(function () {
				if (!_this.m_requestUrlAndHeaderInfo) {
					return _this.m_requestUrlAndHeaderInfoResolver._resolveRequestUrlAndHeaderInfo().then(function (value) {
						_this.m_requestUrlAndHeaderInfo=value;
						if (!_this.m_requestUrlAndHeaderInfo) {
							_this.m_requestUrlAndHeaderInfo={ url: CoreConstants.localDocument, headers: {} };
						}
						if (Utility.isNullOrEmptyString(_this.m_requestUrlAndHeaderInfo.url)) {
							_this.m_requestUrlAndHeaderInfo.url=CoreConstants.localDocument;
						}
						if (!_this.m_requestUrlAndHeaderInfo.headers) {
							_this.m_requestUrlAndHeaderInfo.headers={};
						}
						if (typeof _this.m_requestUrlAndHeaderInfoResolver._createRequestExecutorOrNull==='function') {
							var executor=_this.m_requestUrlAndHeaderInfoResolver._createRequestExecutorOrNull();
							if (executor) {
								_this._requestExecutor=executor;
							}
						}
					});
				}
			});
		};
		ClientRequestContext.prototype.syncPrivateMain=function () {
			var _this=this;
			return this.ensureRequestUrlAndHeaderInfo().then(function () {
				var req=_this._pendingRequest;
				_this.m_pendingRequest=null;
				return _this.processPreSyncPromises(req).then(function () { return _this.syncPrivate(req); });
			});
		};
		ClientRequestContext.prototype.syncPrivate=function (req) {
			var _this=this;
			if (!req.hasActions) {
				return this.processPendingEventHandlers(req);
			}
			var _a=req.buildRequestMessageBodyAndRequestFlags(), msgBody=_a.body, requestFlags=_a.flags;
			if (this._requestFlagModifier) {
				requestFlags |=this._requestFlagModifier;
			}
			if (!this._requestExecutor) {
				if (CoreUtility._isLocalDocumentUrl(this.m_requestUrlAndHeaderInfo.url)) {
					this._requestExecutor=new OfficeJsRequestExecutor(this);
				}
				else {
					this._requestExecutor=new HttpRequestExecutor();
				}
			}
			var requestExecutor=this._requestExecutor;
			var headers={};
			CoreUtility._copyHeaders(this.m_requestUrlAndHeaderInfo.headers, headers);
			CoreUtility._copyHeaders(this.m_customRequestHeaders, headers);
			var requestExecutorRequestMessage={
				Url: this.m_requestUrlAndHeaderInfo.url,
				Headers: headers,
				Body: msgBody
			};
			req.invalidatePendingInvalidObjectPaths();
			var errorFromResponse=null;
			var errorFromProcessEventHandlers=null;
			this._lastSyncStart=typeof performance==='undefined' ? 0 : performance.now();
			this._lastRequestFlags=requestFlags;
			return requestExecutor
				.executeAsync(this._customData, requestFlags, requestExecutorRequestMessage)
				.then(function (response) {
				_this._lastSyncEnd=typeof performance==='undefined' ? 0 : performance.now();
				errorFromResponse=_this.processRequestExecutorResponseMessage(req, response);
				return _this.processPendingEventHandlers(req)["catch"](function (ex) {
					CoreUtility.log('Error in processPendingEventHandlers');
					CoreUtility.log(JSON.stringify(ex));
					errorFromProcessEventHandlers=ex;
				});
			})
				.then(function () {
				if (errorFromResponse) {
					CoreUtility.log('Throw error from response: '+JSON.stringify(errorFromResponse));
					throw errorFromResponse;
				}
				if (errorFromProcessEventHandlers) {
					CoreUtility.log('Throw error from ProcessEventHandler: '+JSON.stringify(errorFromProcessEventHandlers));
					var transformedError=null;
					if (errorFromProcessEventHandlers instanceof _Internal.RuntimeError) {
						transformedError=errorFromProcessEventHandlers;
						transformedError.traceMessages=req._responseTraceMessages;
					}
					else {
						var message=null;
						if (typeof errorFromProcessEventHandlers==='string') {
							message=errorFromProcessEventHandlers;
						}
						else {
							message=errorFromProcessEventHandlers.message;
						}
						if (Utility.isNullOrEmptyString(message)) {
							message=CoreUtility._getResourceString(ResourceStrings.cannotRegisterEvent);
						}
						transformedError=new _Internal.RuntimeError({
							code: ErrorCodes.cannotRegisterEvent,
							message: message,
							traceMessages: req._responseTraceMessages
						});
					}
					throw transformedError;
				}
			});
		};
		ClientRequestContext.prototype.processRequestExecutorResponseMessage=function (req, response) {
			if (response.Body && response.Body.TraceIds) {
				req._setResponseTraceIds(response.Body.TraceIds);
			}
			var traceMessages=req._responseTraceMessages;
			var errorStatementInfo=null;
			if (response.Body) {
				if (response.Body.Error && response.Body.Error.ActionIndex >=0) {
					var prettyPrinter=new RequestPrettyPrinter(this._rootObjectPropertyName, req._objectPaths, req._actions, false, true);
					var debugInfoStatementInfo=prettyPrinter.processForDebugStatementInfo(response.Body.Error.ActionIndex);
					errorStatementInfo={
						statement: debugInfoStatementInfo.statement,
						surroundingStatements: debugInfoStatementInfo.surroundingStatements,
						fullStatements: ['Please enable config.extendedErrorLogging to see full statements.']
					};
					if (OfficeExtension_1.config.extendedErrorLogging) {
						prettyPrinter=new RequestPrettyPrinter(this._rootObjectPropertyName, req._objectPaths, req._actions, false, false);
						errorStatementInfo.fullStatements=prettyPrinter.process();
					}
				}
				var actionResults=null;
				if (response.Body.Results) {
					actionResults=response.Body.Results;
				}
				else if (response.Body.ProcessedResults && response.Body.ProcessedResults.Results) {
					actionResults=response.Body.ProcessedResults.Results;
				}
				if (actionResults) {
					this._processingResult=true;
					try {
						req.processResponse(actionResults);
					}
					finally {
						this._processingResult=false;
					}
				}
			}
			if (!Utility.isNullOrEmptyString(response.ErrorCode)) {
				return new _Internal.RuntimeError({
					code: response.ErrorCode,
					message: response.ErrorMessage,
					traceMessages: traceMessages
				});
			}
			else if (response.Body && response.Body.Error) {
				var debugInfo={
					errorLocation: response.Body.Error.Location
				};
				if (errorStatementInfo) {
					debugInfo.statement=errorStatementInfo.statement;
					debugInfo.surroundingStatements=errorStatementInfo.surroundingStatements;
					debugInfo.fullStatements=errorStatementInfo.fullStatements;
				}
				return new _Internal.RuntimeError({
					code: response.Body.Error.Code,
					message: response.Body.Error.Message,
					traceMessages: traceMessages,
					debugInfo: debugInfo
				});
			}
			return null;
		};
		ClientRequestContext.prototype.processPendingEventHandlers=function (req) {
			var ret=Utility._createPromiseFromResult(null);
			for (var i=0; i < req._pendingProcessEventHandlers.length; i++) {
				var eventHandlers=req._pendingProcessEventHandlers[i];
				ret=ret.then(this.createProcessOneEventHandlersFunc(eventHandlers, req));
			}
			return ret;
		};
		ClientRequestContext.prototype.createProcessOneEventHandlersFunc=function (eventHandlers, req) {
			return function () { return eventHandlers._processRegistration(req); };
		};
		ClientRequestContext.prototype.processPreSyncPromises=function (req) {
			var ret=Utility._createPromiseFromResult(null);
			for (var i=0; i < req._preSyncPromises.length; i++) {
				var p=req._preSyncPromises[i];
				ret=ret.then(this.createProcessOneProSyncFunc(p));
			}
			return ret;
		};
		ClientRequestContext.prototype.createProcessOneProSyncFunc=function (p) {
			return function () { return p; };
		};
		ClientRequestContext.prototype.sync=function (passThroughValue) {
			return this.syncPrivateMain().then(function () { return passThroughValue; });
		};
		ClientRequestContext.prototype.batch=function (batchBody) {
			var _this=this;
			if (this.m_batchMode !==1) {
				return CoreUtility._createPromiseFromException(Utility.createRuntimeError(CoreErrorCodes.generalException, null, null));
			}
			if (this.m_explicitBatchInProgress) {
				return CoreUtility._createPromiseFromException(Utility.createRuntimeError(CoreErrorCodes.generalException, CoreUtility._getResourceString(ResourceStrings.pendingBatchInProgress), null));
			}
			if (Utility.isNullOrUndefined(batchBody)) {
				return Utility._createPromiseFromResult(null);
			}
			this.m_explicitBatchInProgress=true;
			var previousRequest=this.m_pendingRequest;
			this.m_pendingRequest=new ClientRequest(this);
			var batchBodyResult;
			try {
				batchBodyResult=batchBody(this._rootObject, this);
			}
			catch (ex) {
				this.m_explicitBatchInProgress=false;
				this.m_pendingRequest=previousRequest;
				return CoreUtility._createPromiseFromException(ex);
			}
			var request;
			var batchBodyResultPromise;
			if (typeof batchBodyResult==='object' && batchBodyResult && typeof batchBodyResult.then==='function') {
				batchBodyResultPromise=Utility._createPromiseFromResult(null)
					.then(function () {
					return batchBodyResult;
				})
					.then(function (result) {
					_this.m_explicitBatchInProgress=false;
					request=_this.m_pendingRequest;
					_this.m_pendingRequest=previousRequest;
					return result;
				})["catch"](function (ex) {
					_this.m_explicitBatchInProgress=false;
					request=_this.m_pendingRequest;
					_this.m_pendingRequest=previousRequest;
					return CoreUtility._createPromiseFromException(ex);
				});
			}
			else {
				this.m_explicitBatchInProgress=false;
				request=this.m_pendingRequest;
				this.m_pendingRequest=previousRequest;
				batchBodyResultPromise=Utility._createPromiseFromResult(batchBodyResult);
			}
			return batchBodyResultPromise.then(function (result) {
				return _this.ensureRequestUrlAndHeaderInfo()
					.then(function () {
					return _this.syncPrivate(request);
				})
					.then(function () {
					return result;
				});
			});
		};
		ClientRequestContext._run=function (ctxInitializer, runBody, numCleanupAttempts, retryDelay, onCleanupSuccess, onCleanupFailure) {
			if (numCleanupAttempts===void 0) {
				numCleanupAttempts=3;
			}
			if (retryDelay===void 0) {
				retryDelay=5000;
			}
			return ClientRequestContext._runCommon('run', null, ctxInitializer, 0, runBody, numCleanupAttempts, retryDelay, null, onCleanupSuccess, onCleanupFailure);
		};
		ClientRequestContext.isValidRequestInfo=function (value) {
			return (typeof value==='string' ||
				ClientRequestContext.isRequestUrlAndHeaderInfo(value) ||
				ClientRequestContext.isRequestUrlAndHeaderInfoResolver(value));
		};
		ClientRequestContext.isRequestUrlAndHeaderInfo=function (value) {
			return (typeof value==='object' &&
				value !==null &&
				Object.getPrototypeOf(value)===Object.getPrototypeOf({}) &&
				!Utility.isNullOrUndefined(value.url));
		};
		ClientRequestContext.isRequestUrlAndHeaderInfoResolver=function (value) {
			return typeof value==='object' && value !==null && typeof value._resolveRequestUrlAndHeaderInfo==='function';
		};
		ClientRequestContext._runBatch=function (functionName, receivedRunArgs, ctxInitializer, onBeforeRun, numCleanupAttempts, retryDelay, onCleanupSuccess, onCleanupFailure) {
			if (numCleanupAttempts===void 0) {
				numCleanupAttempts=3;
			}
			if (retryDelay===void 0) {
				retryDelay=5000;
			}
			return ClientRequestContext._runBatchCommon(0, functionName, receivedRunArgs, ctxInitializer, numCleanupAttempts, retryDelay, onBeforeRun, onCleanupSuccess, onCleanupFailure);
		};
		ClientRequestContext._runExplicitBatch=function (functionName, receivedRunArgs, ctxInitializer, onBeforeRun, numCleanupAttempts, retryDelay, onCleanupSuccess, onCleanupFailure) {
			if (numCleanupAttempts===void 0) {
				numCleanupAttempts=3;
			}
			if (retryDelay===void 0) {
				retryDelay=5000;
			}
			return ClientRequestContext._runBatchCommon(1, functionName, receivedRunArgs, ctxInitializer, numCleanupAttempts, retryDelay, onBeforeRun, onCleanupSuccess, onCleanupFailure);
		};
		ClientRequestContext._runBatchCommon=function (batchMode, functionName, receivedRunArgs, ctxInitializer, numCleanupAttempts, retryDelay, onBeforeRun, onCleanupSuccess, onCleanupFailure) {
			if (numCleanupAttempts===void 0) {
				numCleanupAttempts=3;
			}
			if (retryDelay===void 0) {
				retryDelay=5000;
			}
			var ctxRetriever;
			var batch;
			var requestInfo=null;
			var previousObjects=null;
			var argOffset=0;
			var options=null;
			if (receivedRunArgs.length > 0) {
				if (ClientRequestContext.isValidRequestInfo(receivedRunArgs[0])) {
					requestInfo=receivedRunArgs[0];
					argOffset=1;
				}
				else if (Utility.isPlainJsonObject(receivedRunArgs[0])) {
					options=receivedRunArgs[0];
					requestInfo=options.session;
					if (requestInfo !=null && !ClientRequestContext.isValidRequestInfo(requestInfo)) {
						return ClientRequestContext.createErrorPromise(functionName);
					}
					previousObjects=options.previousObjects;
					argOffset=1;
				}
			}
			if (receivedRunArgs.length==argOffset+1) {
				batch=receivedRunArgs[argOffset+0];
			}
			else if (options==null && receivedRunArgs.length==argOffset+2) {
				previousObjects=receivedRunArgs[argOffset+0];
				batch=receivedRunArgs[argOffset+1];
			}
			else {
				return ClientRequestContext.createErrorPromise(functionName);
			}
			if (previousObjects !=null) {
				if (previousObjects instanceof ClientObject) {
					ctxRetriever=function () { return previousObjects.context; };
				}
				else if (previousObjects instanceof ClientRequestContext) {
					ctxRetriever=function () { return previousObjects; };
				}
				else if (Array.isArray(previousObjects)) {
					var array=previousObjects;
					if (array.length==0) {
						return ClientRequestContext.createErrorPromise(functionName);
					}
					for (var i=0; i < array.length; i++) {
						if (!(array[i] instanceof ClientObject)) {
							return ClientRequestContext.createErrorPromise(functionName);
						}
						if (array[i].context !=array[0].context) {
							return ClientRequestContext.createErrorPromise(functionName, ResourceStrings.invalidRequestContext);
						}
					}
					ctxRetriever=function () { return array[0].context; };
				}
				else {
					return ClientRequestContext.createErrorPromise(functionName);
				}
			}
			else {
				ctxRetriever=ctxInitializer;
			}
			var onBeforeRunWithOptions=null;
			if (onBeforeRun) {
				onBeforeRunWithOptions=function (context) { return onBeforeRun(options || {}, context); };
			}
			return ClientRequestContext._runCommon(functionName, requestInfo, ctxRetriever, batchMode, batch, numCleanupAttempts, retryDelay, onBeforeRunWithOptions, onCleanupSuccess, onCleanupFailure);
		};
		ClientRequestContext.createErrorPromise=function (functionName, code) {
			if (code===void 0) {
				code=CoreResourceStrings.invalidArgument;
			}
			return CoreUtility._createPromiseFromException(Utility.createRuntimeError(code, CoreUtility._getResourceString(code), functionName));
		};
		ClientRequestContext._runCommon=function (functionName, requestInfo, ctxRetriever, batchMode, runBody, numCleanupAttempts, retryDelay, onBeforeRun, onCleanupSuccess, onCleanupFailure) {
			if (SessionBase._overrideSession) {
				requestInfo=SessionBase._overrideSession;
			}
			var starterPromise=CoreUtility.createPromise(function (resolve, reject) {
				resolve();
			});
			var ctx;
			var succeeded=false;
			var resultOrError;
			var previousBatchMode;
			return starterPromise
				.then(function () {
				ctx=ctxRetriever(requestInfo);
				if (ctx._autoCleanup) {
					return new OfficeExtension_1.Promise(function (resolve, reject) {
						ctx._onRunFinishedNotifiers.push(function () {
							ctx._autoCleanup=true;
							resolve();
						});
					});
				}
				else {
					ctx._autoCleanup=true;
				}
			})
				.then(function () {
				if (typeof runBody !=='function') {
					return ClientRequestContext.createErrorPromise(functionName);
				}
				previousBatchMode=ctx.m_batchMode;
				ctx.m_batchMode=batchMode;
				if (onBeforeRun) {
					onBeforeRun(ctx);
				}
				var runBodyResult;
				if (batchMode==1) {
					runBodyResult=runBody(ctx.batch.bind(ctx));
				}
				else {
					runBodyResult=runBody(ctx);
				}
				if (Utility.isNullOrUndefined(runBodyResult) || typeof runBodyResult.then !=='function') {
					Utility.throwError(ResourceStrings.runMustReturnPromise);
				}
				return runBodyResult;
			})
				.then(function (runBodyResult) {
				if (batchMode===1) {
					return runBodyResult;
				}
				else {
					return ctx.sync(runBodyResult);
				}
			})
				.then(function (result) {
				succeeded=true;
				resultOrError=result;
			})["catch"](function (error) {
				resultOrError=error;
			})
				.then(function () {
				var itemsToRemove=ctx.trackedObjects._retrieveAndClearAutoCleanupList();
				ctx._autoCleanup=false;
				ctx.m_batchMode=previousBatchMode;
				for (var key in itemsToRemove) {
					itemsToRemove[key]._objectPath.isValid=false;
				}
				var cleanupCounter=0;
				if (Utility._synchronousCleanup || ClientRequestContext.isRequestUrlAndHeaderInfoResolver(requestInfo)) {
					return attemptCleanup();
				}
				else {
					attemptCleanup();
				}
				function attemptCleanup() {
					cleanupCounter++;
					var savedPendingRequest=ctx.m_pendingRequest;
					var savedBatchMode=ctx.m_batchMode;
					var request=new ClientRequest(ctx);
					ctx.m_pendingRequest=request;
					ctx.m_batchMode=0;
					try {
						for (var key in itemsToRemove) {
							ctx.trackedObjects.remove(itemsToRemove[key]);
						}
					}
					finally {
						ctx.m_batchMode=savedBatchMode;
						ctx.m_pendingRequest=savedPendingRequest;
					}
					return ctx
						.syncPrivate(request)
						.then(function () {
						if (onCleanupSuccess) {
							onCleanupSuccess(cleanupCounter);
						}
					})["catch"](function () {
						if (onCleanupFailure) {
							onCleanupFailure(cleanupCounter);
						}
						if (cleanupCounter < numCleanupAttempts) {
							setTimeout(function () {
								attemptCleanup();
							}, retryDelay);
						}
					});
				}
			})
				.then(function () {
				if (ctx._onRunFinishedNotifiers && ctx._onRunFinishedNotifiers.length > 0) {
					var func=ctx._onRunFinishedNotifiers.shift();
					func();
				}
				if (succeeded) {
					return resultOrError;
				}
				else {
					throw resultOrError;
				}
			});
		};
		return ClientRequestContext;
	}(ClientRequestContextBase));
	OfficeExtension_1.ClientRequestContext=ClientRequestContext;
	var RetrieveResultImpl=(function () {
		function RetrieveResultImpl(m_proxy, m_shouldPolyfill) {
			this.m_proxy=m_proxy;
			this.m_shouldPolyfill=m_shouldPolyfill;
			var scalarPropertyNames=m_proxy[Constants.scalarPropertyNames];
			var navigationPropertyNames=m_proxy[Constants.navigationPropertyNames];
			var typeName=m_proxy[Constants.className];
			var isCollection=m_proxy[Constants.isCollection];
			if (scalarPropertyNames) {
				for (var i=0; i < scalarPropertyNames.length; i++) {
					Utility.definePropertyThrowUnloadedException(this, typeName, scalarPropertyNames[i]);
				}
			}
			if (navigationPropertyNames) {
				for (var i=0; i < navigationPropertyNames.length; i++) {
					Utility.definePropertyThrowUnloadedException(this, typeName, navigationPropertyNames[i]);
				}
			}
			if (isCollection) {
				Utility.definePropertyThrowUnloadedException(this, typeName, Constants.itemsLowerCase);
			}
		}
		Object.defineProperty(RetrieveResultImpl.prototype, "$proxy", {
			get: function () {
				return this.m_proxy;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(RetrieveResultImpl.prototype, "$isNullObject", {
			get: function () {
				if (!this.m_isLoaded) {
					throw new _Internal.RuntimeError({
						code: ErrorCodes.valueNotLoaded,
						message: CoreUtility._getResourceString(ResourceStrings.valueNotLoaded),
						debugInfo: {
							errorLocation: 'retrieveResult.$isNullObject'
						}
					});
				}
				return this.m_isNullObject;
			},
			enumerable: true,
			configurable: true
		});
		RetrieveResultImpl.prototype.toJSON=function () {
			if (!this.m_isLoaded) {
				return undefined;
			}
			if (this.m_isNullObject) {
				return null;
			}
			if (Utility.isUndefined(this.m_json)) {
				this.m_json=Utility.purifyJson(this.m_value);
			}
			return this.m_json;
		};
		RetrieveResultImpl.prototype.toString=function () {
			return JSON.stringify(this.toJSON());
		};
		RetrieveResultImpl.prototype._handleResult=function (value) {
			this.m_isLoaded=true;
			if (value===null || (typeof value==='object' && value && value._IsNull)) {
				this.m_isNullObject=true;
				value=null;
			}
			else {
				this.m_isNullObject=false;
			}
			if (this.m_shouldPolyfill) {
				value=Utility.changePropertyNameToCamelLowerCase(value);
			}
			this.m_value=value;
			this.m_proxy._handleRetrieveResult(value, this);
		};
		return RetrieveResultImpl;
	}());
	var Constants=(function (_super) {
		__extends(Constants, _super);
		function Constants() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Constants.getItemAt='GetItemAt';
		Constants.index='_Index';
		Constants.iterativeExecutor='IterativeExecutor';
		Constants.isTracked='_IsTracked';
		Constants.eventMessageCategory=65536;
		Constants.eventWorkbookId='Workbook';
		Constants.eventSourceRemote='Remote';
		Constants.proxy='$proxy';
		Constants.className='_className';
		Constants.isCollection='_isCollection';
		Constants.collectionPropertyPath='_collectionPropertyPath';
		Constants.objectPathInfoDoNotKeepReferenceFieldName='D';
		return Constants;
	}(CommonConstants));
	OfficeExtension_1.Constants=Constants;
	var ClientRequest=(function (_super) {
		__extends(ClientRequest, _super);
		function ClientRequest(context) {
			var _this=_super.call(this, context) || this;
			_this.m_context=context;
			_this.m_pendingProcessEventHandlers=[];
			_this.m_pendingEventHandlerActions={};
			_this.m_traceInfos={};
			_this.m_responseTraceIds={};
			_this.m_responseTraceMessages=[];
			return _this;
		}
		Object.defineProperty(ClientRequest.prototype, "traceInfos", {
			get: function () {
				return this.m_traceInfos;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ClientRequest.prototype, "_responseTraceMessages", {
			get: function () {
				return this.m_responseTraceMessages;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ClientRequest.prototype, "_responseTraceIds", {
			get: function () {
				return this.m_responseTraceIds;
			},
			enumerable: true,
			configurable: true
		});
		ClientRequest.prototype._setResponseTraceIds=function (value) {
			if (value) {
				for (var i=0; i < value.length; i++) {
					var traceId=value[i];
					this.m_responseTraceIds[traceId]=traceId;
					var message=this.m_traceInfos[traceId];
					if (!CoreUtility.isNullOrUndefined(message)) {
						this.m_responseTraceMessages.push(message);
					}
				}
			}
		};
		ClientRequest.prototype.addTrace=function (actionId, message) {
			this.m_traceInfos[actionId]=message;
		};
		ClientRequest.prototype._addPendingEventHandlerAction=function (eventHandlers, action) {
			if (!this.m_pendingEventHandlerActions[eventHandlers._id]) {
				this.m_pendingEventHandlerActions[eventHandlers._id]=[];
				this.m_pendingProcessEventHandlers.push(eventHandlers);
			}
			this.m_pendingEventHandlerActions[eventHandlers._id].push(action);
		};
		Object.defineProperty(ClientRequest.prototype, "_pendingProcessEventHandlers", {
			get: function () {
				return this.m_pendingProcessEventHandlers;
			},
			enumerable: true,
			configurable: true
		});
		ClientRequest.prototype._getPendingEventHandlerActions=function (eventHandlers) {
			return this.m_pendingEventHandlerActions[eventHandlers._id];
		};
		return ClientRequest;
	}(ClientRequestBase));
	OfficeExtension_1.ClientRequest=ClientRequest;
	var EventHandlers=(function () {
		function EventHandlers(context, parentObject, name, eventInfo) {
			var _this=this;
			this.m_id=context._nextId();
			this.m_context=context;
			this.m_name=name;
			this.m_handlers=[];
			this.m_registered=false;
			this.m_eventInfo=eventInfo;
			this.m_callback=function (args) {
				_this.m_eventInfo.eventArgsTransformFunc(args).then(function (newArgs) { return _this.fireEvent(newArgs); });
			};
		}
		Object.defineProperty(EventHandlers.prototype, "_registered", {
			get: function () {
				return this.m_registered;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(EventHandlers.prototype, "_id", {
			get: function () {
				return this.m_id;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(EventHandlers.prototype, "_handlers", {
			get: function () {
				return this.m_handlers;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(EventHandlers.prototype, "_context", {
			get: function () {
				return this.m_context;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(EventHandlers.prototype, "_callback", {
			get: function () {
				return this.m_callback;
			},
			enumerable: true,
			configurable: true
		});
		EventHandlers.prototype.add=function (handler) {
			var action=ActionFactory.createTraceAction(this.m_context, null, false);
			this.m_context._pendingRequest._addPendingEventHandlerAction(this, {
				id: action.actionInfo.Id,
				handler: handler,
				operation: 0
			});
			return new EventHandlerResult(this.m_context, this, handler);
		};
		EventHandlers.prototype.remove=function (handler) {
			var action=ActionFactory.createTraceAction(this.m_context, null, false);
			this.m_context._pendingRequest._addPendingEventHandlerAction(this, {
				id: action.actionInfo.Id,
				handler: handler,
				operation: 1
			});
		};
		EventHandlers.prototype.removeAll=function () {
			var action=ActionFactory.createTraceAction(this.m_context, null, false);
			this.m_context._pendingRequest._addPendingEventHandlerAction(this, {
				id: action.actionInfo.Id,
				handler: null,
				operation: 2
			});
		};
		EventHandlers.prototype._processRegistration=function (req) {
			var _this=this;
			var ret=CoreUtility._createPromiseFromResult(null);
			var actions=req._getPendingEventHandlerActions(this);
			if (!actions) {
				return ret;
			}
			var handlersResult=[];
			for (var i=0; i < this.m_handlers.length; i++) {
				handlersResult.push(this.m_handlers[i]);
			}
			var hasChange=false;
			for (var i=0; i < actions.length; i++) {
				if (req._responseTraceIds[actions[i].id]) {
					hasChange=true;
					switch (actions[i].operation) {
						case 0:
							handlersResult.push(actions[i].handler);
							break;
						case 1:
							for (var index=handlersResult.length - 1; index >=0; index--) {
								if (handlersResult[index]===actions[i].handler) {
									handlersResult.splice(index, 1);
									break;
								}
							}
							break;
						case 2:
							handlersResult=[];
							break;
					}
				}
			}
			if (hasChange) {
				if (!this.m_registered && handlersResult.length > 0) {
					ret=ret.then(function () { return _this.m_eventInfo.registerFunc(_this.m_callback); }).then(function () { return (_this.m_registered=true); });
				}
				else if (this.m_registered && handlersResult.length==0) {
					ret=ret
						.then(function () { return _this.m_eventInfo.unregisterFunc(_this.m_callback); })["catch"](function (ex) {
						CoreUtility.log('Error when unregister event: '+JSON.stringify(ex));
					})
						.then(function () { return (_this.m_registered=false); });
				}
				ret=ret.then(function () { return (_this.m_handlers=handlersResult); });
			}
			return ret;
		};
		EventHandlers.prototype.fireEvent=function (args) {
			var promises=[];
			for (var i=0; i < this.m_handlers.length; i++) {
				var handler=this.m_handlers[i];
				var p=CoreUtility._createPromiseFromResult(null)
					.then(this.createFireOneEventHandlerFunc(handler, args))["catch"](function (ex) {
					CoreUtility.log('Error when invoke handler: '+JSON.stringify(ex));
				});
				promises.push(p);
			}
			CoreUtility.Promise.all(promises);
		};
		EventHandlers.prototype.createFireOneEventHandlerFunc=function (handler, args) {
			return function () { return handler(args); };
		};
		return EventHandlers;
	}());
	OfficeExtension_1.EventHandlers=EventHandlers;
	var EventHandlerResult=(function () {
		function EventHandlerResult(context, handlers, handler) {
			this.m_context=context;
			this.m_allHandlers=handlers;
			this.m_handler=handler;
		}
		Object.defineProperty(EventHandlerResult.prototype, "context", {
			get: function () {
				return this.m_context;
			},
			enumerable: true,
			configurable: true
		});
		EventHandlerResult.prototype.remove=function () {
			if (this.m_allHandlers && this.m_handler) {
				this.m_allHandlers.remove(this.m_handler);
				this.m_allHandlers=null;
				this.m_handler=null;
			}
		};
		return EventHandlerResult;
	}());
	OfficeExtension_1.EventHandlerResult=EventHandlerResult;
	(function (_Internal) {
		var OfficeJsEventRegistration=(function () {
			function OfficeJsEventRegistration() {
			}
			OfficeJsEventRegistration.prototype.register=function (eventId, targetId, handler) {
				switch (eventId) {
					case 4:
						return Utility.promisify(function (callback) { return Office.context.document.bindings.getByIdAsync(targetId, callback); }).then(function (officeBinding) {
							return Utility.promisify(function (callback) {
								return officeBinding.addHandlerAsync(Office.EventType.BindingDataChanged, handler, callback);
							});
						});
					case 3:
						return Utility.promisify(function (callback) { return Office.context.document.bindings.getByIdAsync(targetId, callback); }).then(function (officeBinding) {
							return Utility.promisify(function (callback) {
								return officeBinding.addHandlerAsync(Office.EventType.BindingSelectionChanged, handler, callback);
							});
						});
					case 2:
						return Utility.promisify(function (callback) {
							return Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, handler, callback);
						});
					case 1:
						return Utility.promisify(function (callback) {
							return Office.context.document.settings.addHandlerAsync(Office.EventType.SettingsChanged, handler, callback);
						});
					case 5:
						return Utility.promisify(function (callback) {
							return OSF.DDA.RichApi.richApiMessageManager.addHandlerAsync('richApiMessage', handler, callback);
						});
					case 13:
						return Utility.promisify(function (callback) {
							return Office.context.document.addHandlerAsync(Office.EventType.ObjectDeleted, handler, { id: targetId }, callback);
						});
					case 14:
						return Utility.promisify(function (callback) {
							return Office.context.document.addHandlerAsync(Office.EventType.ObjectSelectionChanged, handler, { id: targetId }, callback);
						});
					case 15:
						return Utility.promisify(function (callback) {
							return Office.context.document.addHandlerAsync(Office.EventType.ObjectDataChanged, handler, { id: targetId }, callback);
						});
					case 16:
						return Utility.promisify(function (callback) {
							return Office.context.document.addHandlerAsync(Office.EventType.ContentControlAdded, handler, { id: targetId }, callback);
						});
					default:
						throw _Internal.RuntimeError._createInvalidArgError({ argumentName: 'eventId' });
				}
			};
			OfficeJsEventRegistration.prototype.unregister=function (eventId, targetId, handler) {
				switch (eventId) {
					case 4:
						return Utility.promisify(function (callback) { return Office.context.document.bindings.getByIdAsync(targetId, callback); }).then(function (officeBinding) {
							return Utility.promisify(function (callback) {
								return officeBinding.removeHandlerAsync(Office.EventType.BindingDataChanged, { handler: handler }, callback);
							});
						});
					case 3:
						return Utility.promisify(function (callback) { return Office.context.document.bindings.getByIdAsync(targetId, callback); }).then(function (officeBinding) {
							return Utility.promisify(function (callback) {
								return officeBinding.removeHandlerAsync(Office.EventType.BindingSelectionChanged, { handler: handler }, callback);
							});
						});
					case 2:
						return Utility.promisify(function (callback) {
							return Office.context.document.removeHandlerAsync(Office.EventType.DocumentSelectionChanged, { handler: handler }, callback);
						});
					case 1:
						return Utility.promisify(function (callback) {
							return Office.context.document.settings.removeHandlerAsync(Office.EventType.SettingsChanged, { handler: handler }, callback);
						});
					case 5:
						return Utility.promisify(function (callback) {
							return OSF.DDA.RichApi.richApiMessageManager.removeHandlerAsync('richApiMessage', { handler: handler }, callback);
						});
					case 13:
						return Utility.promisify(function (callback) {
							return Office.context.document.removeHandlerAsync(Office.EventType.ObjectDeleted, { id: targetId, handler: handler }, callback);
						});
					case 14:
						return Utility.promisify(function (callback) {
							return Office.context.document.removeHandlerAsync(Office.EventType.ObjectSelectionChanged, { id: targetId, handler: handler }, callback);
						});
					case 15:
						return Utility.promisify(function (callback) {
							return Office.context.document.removeHandlerAsync(Office.EventType.ObjectDataChanged, { id: targetId, handler: handler }, callback);
						});
					case 16:
						return Utility.promisify(function (callback) {
							return Office.context.document.removeHandlerAsync(Office.EventType.ContentControlAdded, { id: targetId, handler: handler }, callback);
						});
					default:
						throw _Internal.RuntimeError._createInvalidArgError({ argumentName: 'eventId' });
				}
			};
			return OfficeJsEventRegistration;
		}());
		_Internal.officeJsEventRegistration=new OfficeJsEventRegistration();
	})(_Internal=OfficeExtension_1._Internal || (OfficeExtension_1._Internal={}));
	var EventRegistration=(function () {
		function EventRegistration(registerEventImpl, unregisterEventImpl) {
			this.m_handlersByEventByTarget={};
			this.m_registerEventImpl=registerEventImpl;
			this.m_unregisterEventImpl=unregisterEventImpl;
		}
		EventRegistration.prototype.getHandlers=function (eventId, targetId) {
			if (Utility.isNullOrUndefined(targetId)) {
				targetId='';
			}
			var handlersById=this.m_handlersByEventByTarget[eventId];
			if (!handlersById) {
				handlersById={};
				this.m_handlersByEventByTarget[eventId]=handlersById;
			}
			var handlers=handlersById[targetId];
			if (!handlers) {
				handlers=[];
				handlersById[targetId]=handlers;
			}
			return handlers;
		};
		EventRegistration.prototype.register=function (eventId, targetId, handler) {
			if (!handler) {
				throw _Internal.RuntimeError._createInvalidArgError({ argumentName: 'handler' });
			}
			var handlers=this.getHandlers(eventId, targetId);
			handlers.push(handler);
			if (handlers.length===1) {
				return this.m_registerEventImpl(eventId, targetId);
			}
			return Utility._createPromiseFromResult(null);
		};
		EventRegistration.prototype.unregister=function (eventId, targetId, handler) {
			if (!handler) {
				throw _Internal.RuntimeError._createInvalidArgError({ argumentName: 'handler' });
			}
			var handlers=this.getHandlers(eventId, targetId);
			for (var index=handlers.length - 1; index >=0; index--) {
				if (handlers[index]===handler) {
					handlers.splice(index, 1);
					break;
				}
			}
			if (handlers.length===0) {
				return this.m_unregisterEventImpl(eventId, targetId);
			}
			return Utility._createPromiseFromResult(null);
		};
		return EventRegistration;
	}());
	OfficeExtension_1.EventRegistration=EventRegistration;
	var GenericEventRegistration=(function () {
		function GenericEventRegistration() {
			this.m_eventRegistration=new EventRegistration(this._registerEventImpl.bind(this), this._unregisterEventImpl.bind(this));
			this.m_richApiMessageHandler=this._handleRichApiMessage.bind(this);
		}
		GenericEventRegistration.prototype.ready=function () {
			var _this=this;
			if (!this.m_ready) {
				if (GenericEventRegistration._testReadyImpl) {
					this.m_ready=GenericEventRegistration._testReadyImpl().then(function () {
						_this.m_isReady=true;
					});
				}
				else if (HostBridge.instance) {
					this.m_ready=Utility._createPromiseFromResult(null).then(function () {
						_this.m_isReady=true;
					});
				}
				else {
					this.m_ready=_Internal.officeJsEventRegistration
						.register(5, '', this.m_richApiMessageHandler)
						.then(function () {
						_this.m_isReady=true;
					});
				}
			}
			return this.m_ready;
		};
		Object.defineProperty(GenericEventRegistration.prototype, "isReady", {
			get: function () {
				return this.m_isReady;
			},
			enumerable: true,
			configurable: true
		});
		GenericEventRegistration.prototype.register=function (eventId, targetId, handler) {
			var _this=this;
			return this.ready().then(function () { return _this.m_eventRegistration.register(eventId, targetId, handler); });
		};
		GenericEventRegistration.prototype.unregister=function (eventId, targetId, handler) {
			var _this=this;
			return this.ready().then(function () { return _this.m_eventRegistration.unregister(eventId, targetId, handler); });
		};
		GenericEventRegistration.prototype._registerEventImpl=function (eventId, targetId) {
			return Utility._createPromiseFromResult(null);
		};
		GenericEventRegistration.prototype._unregisterEventImpl=function (eventId, targetId) {
			return Utility._createPromiseFromResult(null);
		};
		GenericEventRegistration.prototype._handleRichApiMessage=function (msg) {
			if (msg && msg.entries) {
				for (var entryIndex=0; entryIndex < msg.entries.length; entryIndex++) {
					var entry=msg.entries[entryIndex];
					if (entry.messageCategory==Constants.eventMessageCategory) {
						if (CoreUtility._logEnabled) {
							CoreUtility.log(JSON.stringify(entry));
						}
						var funcs=this.m_eventRegistration.getHandlers(entry.messageType, entry.targetId);
						if (funcs.length > 0) {
							var arg=JSON.parse(entry.message);
							if (entry.isRemoteOverride) {
								arg.source=Constants.eventSourceRemote;
							}
							for (var i=0; i < funcs.length; i++) {
								funcs[i](arg);
							}
						}
					}
				}
			}
		};
		GenericEventRegistration.getGenericEventRegistration=function () {
			if (!GenericEventRegistration.s_genericEventRegistration) {
				GenericEventRegistration.s_genericEventRegistration=new GenericEventRegistration();
			}
			return GenericEventRegistration.s_genericEventRegistration;
		};
		GenericEventRegistration.richApiMessageEventCategory=65536;
		return GenericEventRegistration;
	}());
	function _testSetRichApiMessageReadyImpl(impl) {
		GenericEventRegistration._testReadyImpl=impl;
	}
	OfficeExtension_1._testSetRichApiMessageReadyImpl=_testSetRichApiMessageReadyImpl;
	function _testTriggerRichApiMessageEvent(msg) {
		GenericEventRegistration.getGenericEventRegistration()._handleRichApiMessage(msg);
	}
	OfficeExtension_1._testTriggerRichApiMessageEvent=_testTriggerRichApiMessageEvent;
	var GenericEventHandlers=(function (_super) {
		__extends(GenericEventHandlers, _super);
		function GenericEventHandlers(context, parentObject, name, eventInfo) {
			var _this=_super.call(this, context, parentObject, name, eventInfo) || this;
			_this.m_genericEventInfo=eventInfo;
			return _this;
		}
		GenericEventHandlers.prototype.add=function (handler) {
			var _this=this;
			if (this._handlers.length==0 && this.m_genericEventInfo.registerFunc) {
				this.m_genericEventInfo.registerFunc();
			}
			if (!GenericEventRegistration.getGenericEventRegistration().isReady) {
				this._context._pendingRequest._addPreSyncPromise(GenericEventRegistration.getGenericEventRegistration().ready());
			}
			ActionFactory.createTraceMarkerForCallback(this._context, function () {
				_this._handlers.push(handler);
				if (_this._handlers.length==1) {
					GenericEventRegistration.getGenericEventRegistration().register(_this.m_genericEventInfo.eventType, _this.m_genericEventInfo.getTargetIdFunc(), _this._callback);
				}
			});
			return new EventHandlerResult(this._context, this, handler);
		};
		GenericEventHandlers.prototype.remove=function (handler) {
			var _this=this;
			if (this._handlers.length==1 && this.m_genericEventInfo.unregisterFunc) {
				this.m_genericEventInfo.unregisterFunc();
			}
			ActionFactory.createTraceMarkerForCallback(this._context, function () {
				var handlers=_this._handlers;
				for (var index=handlers.length - 1; index >=0; index--) {
					if (handlers[index]===handler) {
						handlers.splice(index, 1);
						break;
					}
				}
				if (handlers.length==0) {
					GenericEventRegistration.getGenericEventRegistration().unregister(_this.m_genericEventInfo.eventType, _this.m_genericEventInfo.getTargetIdFunc(), _this._callback);
				}
			});
		};
		GenericEventHandlers.prototype.removeAll=function () { };
		return GenericEventHandlers;
	}(EventHandlers));
	OfficeExtension_1.GenericEventHandlers=GenericEventHandlers;
	var InstantiateActionResultHandler=(function () {
		function InstantiateActionResultHandler(clientObject) {
			this.m_clientObject=clientObject;
		}
		InstantiateActionResultHandler.prototype._handleResult=function (value) {
			this.m_clientObject._handleIdResult(value);
		};
		return InstantiateActionResultHandler;
	}());
	var ObjectPathFactory=(function () {
		function ObjectPathFactory() {
		}
		ObjectPathFactory.createGlobalObjectObjectPath=function (context) {
			var objectPathInfo={
				Id: context._nextId(),
				ObjectPathType: 1,
				Name: ''
			};
			return new ObjectPath(objectPathInfo, null, false, false, 1, 4);
		};
		ObjectPathFactory.createNewObjectObjectPath=function (context, typeName, isCollection, flags) {
			var objectPathInfo={
				Id: context._nextId(),
				ObjectPathType: 2,
				Name: typeName
			};
			var ret=new ObjectPath(objectPathInfo, null, isCollection, false, 1, Utility._fixupApiFlags(flags));
			return ret;
		};
		ObjectPathFactory.createPropertyObjectPath=function (context, parent, propertyName, isCollection, isInvalidAfterRequest, flags) {
			var objectPathInfo={
				Id: context._nextId(),
				ObjectPathType: 4,
				Name: propertyName,
				ParentObjectPathId: parent._objectPath.objectPathInfo.Id
			};
			var ret=new ObjectPath(objectPathInfo, parent._objectPath, isCollection, isInvalidAfterRequest, 1, Utility._fixupApiFlags(flags));
			return ret;
		};
		ObjectPathFactory.createIndexerObjectPath=function (context, parent, args) {
			var objectPathInfo={
				Id: context._nextId(),
				ObjectPathType: 5,
				Name: '',
				ParentObjectPathId: parent._objectPath.objectPathInfo.Id,
				ArgumentInfo: {}
			};
			objectPathInfo.ArgumentInfo.Arguments=args;
			return new ObjectPath(objectPathInfo, parent._objectPath, false, false, 1, 4);
		};
		ObjectPathFactory.createIndexerObjectPathUsingParentPath=function (context, parentObjectPath, args) {
			var objectPathInfo={
				Id: context._nextId(),
				ObjectPathType: 5,
				Name: '',
				ParentObjectPathId: parentObjectPath.objectPathInfo.Id,
				ArgumentInfo: {}
			};
			objectPathInfo.ArgumentInfo.Arguments=args;
			return new ObjectPath(objectPathInfo, parentObjectPath, false, false, 1, 4);
		};
		ObjectPathFactory.createMethodObjectPath=function (context, parent, methodName, operationType, args, isCollection, isInvalidAfterRequest, getByIdMethodName, flags) {
			var objectPathInfo={
				Id: context._nextId(),
				ObjectPathType: 3,
				Name: methodName,
				ParentObjectPathId: parent._objectPath.objectPathInfo.Id,
				ArgumentInfo: {}
			};
			var argumentObjectPaths=Utility.setMethodArguments(context, objectPathInfo.ArgumentInfo, args);
			var ret=new ObjectPath(objectPathInfo, parent._objectPath, isCollection, isInvalidAfterRequest, operationType, Utility._fixupApiFlags(flags));
			ret.argumentObjectPaths=argumentObjectPaths;
			ret.getByIdMethodName=getByIdMethodName;
			return ret;
		};
		ObjectPathFactory.createReferenceIdObjectPath=function (context, referenceId) {
			var objectPathInfo={
				Id: context._nextId(),
				ObjectPathType: 6,
				Name: referenceId,
				ArgumentInfo: {}
			};
			var ret=new ObjectPath(objectPathInfo, null, false, false, 1, 4);
			return ret;
		};
		ObjectPathFactory.createChildItemObjectPathUsingIndexerOrGetItemAt=function (hasIndexerMethod, context, parent, childItem, index) {
			var id=Utility.tryGetObjectIdFromLoadOrRetrieveResult(childItem);
			if (hasIndexerMethod && !Utility.isNullOrUndefined(id)) {
				return ObjectPathFactory.createChildItemObjectPathUsingIndexer(context, parent, childItem);
			}
			else {
				return ObjectPathFactory.createChildItemObjectPathUsingGetItemAt(context, parent, childItem, index);
			}
		};
		ObjectPathFactory.createChildItemObjectPathUsingIndexer=function (context, parent, childItem) {
			var id=Utility.tryGetObjectIdFromLoadOrRetrieveResult(childItem);
			var objectPathInfo=(objectPathInfo={
				Id: context._nextId(),
				ObjectPathType: 5,
				Name: '',
				ParentObjectPathId: parent._objectPath.objectPathInfo.Id,
				ArgumentInfo: {}
			});
			objectPathInfo.ArgumentInfo.Arguments=[id];
			return new ObjectPath(objectPathInfo, parent._objectPath, false, false, 1, 4);
		};
		ObjectPathFactory.createChildItemObjectPathUsingGetItemAt=function (context, parent, childItem, index) {
			var indexFromServer=childItem[Constants.index];
			if (indexFromServer) {
				index=indexFromServer;
			}
			var objectPathInfo={
				Id: context._nextId(),
				ObjectPathType: 3,
				Name: Constants.getItemAt,
				ParentObjectPathId: parent._objectPath.objectPathInfo.Id,
				ArgumentInfo: {}
			};
			objectPathInfo.ArgumentInfo.Arguments=[index];
			return new ObjectPath(objectPathInfo, parent._objectPath, false, false, 1, 4);
		};
		return ObjectPathFactory;
	}());
	OfficeExtension_1.ObjectPathFactory=ObjectPathFactory;
	var OfficeJsRequestExecutor=(function () {
		function OfficeJsRequestExecutor(context) {
			this.m_context=context;
		}
		OfficeJsRequestExecutor.prototype.executeAsync=function (customData, requestFlags, requestMessage) {
			var _this=this;
			var messageSafearray=RichApiMessageUtility.buildMessageArrayForIRequestExecutor(customData, requestFlags, requestMessage, OfficeJsRequestExecutor.SourceLibHeaderValue);
			return new OfficeExtension_1.Promise(function (resolve, reject) {
				OSF.DDA.RichApi.executeRichApiRequestAsync(messageSafearray, function (result) {
					CoreUtility.log('Response:');
					CoreUtility.log(JSON.stringify(result));
					var response;
					if (result.status=='succeeded') {
						response=RichApiMessageUtility.buildResponseOnSuccess(RichApiMessageUtility.getResponseBody(result), RichApiMessageUtility.getResponseHeaders(result));
					}
					else {
						response=RichApiMessageUtility.buildResponseOnError(result.error.code, result.error.message);
						_this.m_context._processOfficeJsErrorResponse(result.error.code, response);
					}
					resolve(response);
				});
			});
		};
		OfficeJsRequestExecutor.SourceLibHeaderValue='officejs';
		return OfficeJsRequestExecutor;
	}());
	var TrackedObjects=(function () {
		function TrackedObjects(context) {
			this._autoCleanupList={};
			this.m_context=context;
		}
		TrackedObjects.prototype.add=function (param) {
			var _this=this;
			if (Array.isArray(param)) {
				param.forEach(function (item) { return _this._addCommon(item, true); });
			}
			else {
				this._addCommon(param, true);
			}
		};
		TrackedObjects.prototype._autoAdd=function (object) {
			this._addCommon(object, false);
			this._autoCleanupList[object._objectPath.objectPathInfo.Id]=object;
		};
		TrackedObjects.prototype._autoTrackIfNecessaryWhenHandleObjectResultValue=function (object, resultValue) {
			var shouldAutoTrack=this.m_context._autoCleanup &&
				!object[Constants.isTracked] &&
				object !==this.m_context._rootObject &&
				resultValue &&
				!Utility.isNullOrEmptyString(resultValue[Constants.referenceId]);
			if (shouldAutoTrack) {
				this._autoCleanupList[object._objectPath.objectPathInfo.Id]=object;
				object[Constants.isTracked]=true;
			}
		};
		TrackedObjects.prototype._addCommon=function (object, isExplicitlyAdded) {
			if (object[Constants.isTracked]) {
				if (isExplicitlyAdded && this.m_context._autoCleanup) {
					delete this._autoCleanupList[object._objectPath.objectPathInfo.Id];
				}
				return;
			}
			var referenceId=object[Constants.referenceId];
			var donotKeepReference=object._objectPath.objectPathInfo[Constants.objectPathInfoDoNotKeepReferenceFieldName];
			if (donotKeepReference) {
				throw Utility.createRuntimeError(CoreErrorCodes.generalException, CoreUtility._getResourceString(ResourceStrings.objectIsUntracked), null);
			}
			if (Utility.isNullOrEmptyString(referenceId) && object._KeepReference) {
				object._KeepReference();
				ActionFactory.createInstantiateAction(this.m_context, object);
				if (isExplicitlyAdded && this.m_context._autoCleanup) {
					delete this._autoCleanupList[object._objectPath.objectPathInfo.Id];
				}
				object[Constants.isTracked]=true;
			}
		};
		TrackedObjects.prototype.remove=function (param) {
			var _this=this;
			if (Array.isArray(param)) {
				param.forEach(function (item) { return _this._removeCommon(item); });
			}
			else {
				this._removeCommon(param);
			}
		};
		TrackedObjects.prototype._removeCommon=function (object) {
			object._objectPath.objectPathInfo[Constants.objectPathInfoDoNotKeepReferenceFieldName]=true;
			object.context._pendingRequest._removeKeepReferenceAction(object._objectPath.objectPathInfo.Id);
			var referenceId=object[Constants.referenceId];
			if (!Utility.isNullOrEmptyString(referenceId)) {
				var rootObject=this.m_context._rootObject;
				if (rootObject._RemoveReference) {
					rootObject._RemoveReference(referenceId);
				}
			}
			delete object[Constants.isTracked];
		};
		TrackedObjects.prototype._retrieveAndClearAutoCleanupList=function () {
			var list=this._autoCleanupList;
			this._autoCleanupList={};
			return list;
		};
		return TrackedObjects;
	}());
	OfficeExtension_1.TrackedObjects=TrackedObjects;
	var RequestPrettyPrinter=(function () {
		function RequestPrettyPrinter(globalObjName, referencedObjectPaths, actions, showDispose, removePII) {
			if (!globalObjName) {
				globalObjName='root';
			}
			this.m_globalObjName=globalObjName;
			this.m_referencedObjectPaths=referencedObjectPaths;
			this.m_actions=actions;
			this.m_statements=[];
			this.m_variableNameForObjectPathMap={};
			this.m_variableNameToObjectPathMap={};
			this.m_declaredObjectPathMap={};
			this.m_showDispose=showDispose;
			this.m_removePII=removePII;
		}
		RequestPrettyPrinter.prototype.process=function () {
			if (this.m_showDispose) {
				ClientRequest._calculateLastUsedObjectPathIds(this.m_actions);
			}
			for (var i=0; i < this.m_actions.length; i++) {
				this.processOneAction(this.m_actions[i]);
			}
			return this.m_statements;
		};
		RequestPrettyPrinter.prototype.processForDebugStatementInfo=function (actionIndex) {
			if (this.m_showDispose) {
				ClientRequest._calculateLastUsedObjectPathIds(this.m_actions);
			}
			var surroundingCount=5;
			this.m_statements=[];
			var oneStatement='';
			var statementIndex=-1;
			for (var i=0; i < this.m_actions.length; i++) {
				this.processOneAction(this.m_actions[i]);
				if (actionIndex==i) {
					statementIndex=this.m_statements.length - 1;
				}
				if (statementIndex >=0 && this.m_statements.length > statementIndex+surroundingCount+1) {
					break;
				}
			}
			if (statementIndex < 0) {
				return null;
			}
			var startIndex=statementIndex - surroundingCount;
			if (startIndex < 0) {
				startIndex=0;
			}
			var endIndex=statementIndex+1+surroundingCount;
			if (endIndex > this.m_statements.length) {
				endIndex=this.m_statements.length;
			}
			var surroundingStatements=[];
			if (startIndex !=0) {
				surroundingStatements.push('...');
			}
			for (var i_1=startIndex; i_1 < statementIndex; i_1++) {
				surroundingStatements.push(this.m_statements[i_1]);
			}
			surroundingStatements.push('// >>>>>');
			surroundingStatements.push(this.m_statements[statementIndex]);
			surroundingStatements.push('// <<<<<');
			for (var i_2=statementIndex+1; i_2 < endIndex; i_2++) {
				surroundingStatements.push(this.m_statements[i_2]);
			}
			if (endIndex < this.m_statements.length) {
				surroundingStatements.push('...');
			}
			return {
				statement: this.m_statements[statementIndex],
				surroundingStatements: surroundingStatements
			};
		};
		RequestPrettyPrinter.prototype.processOneAction=function (action) {
			var actionInfo=action.actionInfo;
			switch (actionInfo.ActionType) {
				case 1:
					this.processInstantiateAction(action);
					break;
				case 3:
					this.processMethodAction(action);
					break;
				case 2:
					this.processQueryAction(action);
					break;
				case 7:
					this.processQueryAsJsonAction(action);
					break;
				case 6:
					this.processRecursiveQueryAction(action);
					break;
				case 4:
					this.processSetPropertyAction(action);
					break;
				case 5:
					this.processTraceAction(action);
					break;
				case 8:
					this.processEnsureUnchangedAction(action);
					break;
				case 9:
					this.processUpdateAction(action);
					break;
			}
		};
		RequestPrettyPrinter.prototype.processInstantiateAction=function (action) {
			var objId=action.actionInfo.ObjectPathId;
			var objPath=this.m_referencedObjectPaths[objId];
			var varName=this.getObjVarName(objId);
			if (!this.m_declaredObjectPathMap[objId]) {
				var statement='var '+varName+'='+this.buildObjectPathExpressionWithParent(objPath)+';';
				statement=this.appendDisposeCommentIfRelevant(statement, action);
				this.m_statements.push(statement);
				this.m_declaredObjectPathMap[objId]=varName;
			}
			else {
				var statement='// Instantiate {'+varName+'}';
				statement=this.appendDisposeCommentIfRelevant(statement, action);
				this.m_statements.push(statement);
			}
		};
		RequestPrettyPrinter.prototype.processMethodAction=function (action) {
			var methodName=action.actionInfo.Name;
			if (methodName==='_KeepReference') {
				if (!OfficeExtension_1._internalConfig.showInternalApiInDebugInfo) {
					return;
				}
				methodName='track';
			}
			var statement=this.getObjVarName(action.actionInfo.ObjectPathId)+				'.'+				Utility._toCamelLowerCase(methodName)+				'('+				this.buildArgumentsExpression(action.actionInfo.ArgumentInfo)+				');';
			statement=this.appendDisposeCommentIfRelevant(statement, action);
			this.m_statements.push(statement);
		};
		RequestPrettyPrinter.prototype.processQueryAction=function (action) {
			var queryExp=this.buildQueryExpression(action);
			var statement=this.getObjVarName(action.actionInfo.ObjectPathId)+'.load('+queryExp+');';
			statement=this.appendDisposeCommentIfRelevant(statement, action);
			this.m_statements.push(statement);
		};
		RequestPrettyPrinter.prototype.processQueryAsJsonAction=function (action) {
			var queryExp=this.buildQueryExpression(action);
			var statement=this.getObjVarName(action.actionInfo.ObjectPathId)+'.retrieve('+queryExp+');';
			statement=this.appendDisposeCommentIfRelevant(statement, action);
			this.m_statements.push(statement);
		};
		RequestPrettyPrinter.prototype.processRecursiveQueryAction=function (action) {
			var queryExp='';
			if (action.actionInfo.RecursiveQueryInfo) {
				queryExp=JSON.stringify(action.actionInfo.RecursiveQueryInfo);
			}
			var statement=this.getObjVarName(action.actionInfo.ObjectPathId)+'.loadRecursive('+queryExp+');';
			statement=this.appendDisposeCommentIfRelevant(statement, action);
			this.m_statements.push(statement);
		};
		RequestPrettyPrinter.prototype.processSetPropertyAction=function (action) {
			var statement=this.getObjVarName(action.actionInfo.ObjectPathId)+				'.'+				Utility._toCamelLowerCase(action.actionInfo.Name)+				'='+				this.buildArgumentsExpression(action.actionInfo.ArgumentInfo)+				';';
			statement=this.appendDisposeCommentIfRelevant(statement, action);
			this.m_statements.push(statement);
		};
		RequestPrettyPrinter.prototype.processTraceAction=function (action) {
			var statement='context.trace();';
			statement=this.appendDisposeCommentIfRelevant(statement, action);
			this.m_statements.push(statement);
		};
		RequestPrettyPrinter.prototype.processEnsureUnchangedAction=function (action) {
			var statement=this.getObjVarName(action.actionInfo.ObjectPathId)+				'.ensureUnchanged('+				JSON.stringify(action.actionInfo.ObjectState)+				');';
			statement=this.appendDisposeCommentIfRelevant(statement, action);
			this.m_statements.push(statement);
		};
		RequestPrettyPrinter.prototype.processUpdateAction=function (action) {
			var statement=this.getObjVarName(action.actionInfo.ObjectPathId)+				'.update('+				JSON.stringify(action.actionInfo.ObjectState)+				');';
			statement=this.appendDisposeCommentIfRelevant(statement, action);
			this.m_statements.push(statement);
		};
		RequestPrettyPrinter.prototype.appendDisposeCommentIfRelevant=function (statement, action) {
			var _this=this;
			if (this.m_showDispose) {
				var lastUsedObjectPathIds=action.actionInfo.L;
				if (lastUsedObjectPathIds && lastUsedObjectPathIds.length > 0) {
					var objectNamesToDispose=lastUsedObjectPathIds.map(function (item) { return _this.getObjVarName(item); }).join(', ');
					return statement+' // And then dispose {'+objectNamesToDispose+'}';
				}
			}
			return statement;
		};
		RequestPrettyPrinter.prototype.buildQueryExpression=function (action) {
			if (action.actionInfo.QueryInfo) {
				var option={};
				option.select=action.actionInfo.QueryInfo.Select;
				option.expand=action.actionInfo.QueryInfo.Expand;
				option.skip=action.actionInfo.QueryInfo.Skip;
				option.top=action.actionInfo.QueryInfo.Top;
				if (typeof option.top==='undefined' &&
					typeof option.skip==='undefined' &&
					typeof option.expand==='undefined') {
					if (typeof option.select==='undefined') {
						return '';
					}
					else {
						return JSON.stringify(option.select);
					}
				}
				else {
					return JSON.stringify(option);
				}
			}
			return '';
		};
		RequestPrettyPrinter.prototype.buildObjectPathExpressionWithParent=function (objPath) {
			var hasParent=objPath.objectPathInfo.ObjectPathType==5 ||
				objPath.objectPathInfo.ObjectPathType==3 ||
				objPath.objectPathInfo.ObjectPathType==4;
			if (hasParent && objPath.objectPathInfo.ParentObjectPathId) {
				return (this.getObjVarName(objPath.objectPathInfo.ParentObjectPathId)+'.'+this.buildObjectPathExpression(objPath));
			}
			return this.buildObjectPathExpression(objPath);
		};
		RequestPrettyPrinter.prototype.buildObjectPathExpression=function (objPath) {
			var expr=this.buildObjectPathInfoExpression(objPath.objectPathInfo);
			var originalObjectPathInfo=objPath.originalObjectPathInfo;
			if (originalObjectPathInfo) {
				expr=expr+' /* originally '+this.buildObjectPathInfoExpression(originalObjectPathInfo)+' */';
			}
			return expr;
		};
		RequestPrettyPrinter.prototype.buildObjectPathInfoExpression=function (objectPathInfo) {
			switch (objectPathInfo.ObjectPathType) {
				case 1:
					return 'context.'+this.m_globalObjName;
				case 5:
					return 'getItem('+this.buildArgumentsExpression(objectPathInfo.ArgumentInfo)+')';
				case 3:
					return (Utility._toCamelLowerCase(objectPathInfo.Name)+						'('+						this.buildArgumentsExpression(objectPathInfo.ArgumentInfo)+						')');
				case 2:
					return objectPathInfo.Name+'.newObject()';
				case 7:
					return 'null';
				case 4:
					return Utility._toCamelLowerCase(objectPathInfo.Name);
				case 6:
					return ('context.'+this.m_globalObjName+'._getObjectByReferenceId('+JSON.stringify(objectPathInfo.Name)+')');
			}
		};
		RequestPrettyPrinter.prototype.buildArgumentsExpression=function (args) {
			var ret='';
			if (!args.Arguments || args.Arguments.length===0) {
				return ret;
			}
			if (this.m_removePII) {
				if (typeof args.Arguments[0]==='undefined') {
					return ret;
				}
				return '...';
			}
			for (var i=0; i < args.Arguments.length; i++) {
				if (i > 0) {
					ret=ret+', ';
				}
				ret=					ret+						this.buildArgumentLiteral(args.Arguments[i], args.ReferencedObjectPathIds ? args.ReferencedObjectPathIds[i] : null);
			}
			if (ret==='undefined') {
				ret='';
			}
			return ret;
		};
		RequestPrettyPrinter.prototype.buildArgumentLiteral=function (value, objectPathId) {
			if (typeof value=='number' && value===objectPathId) {
				return this.getObjVarName(objectPathId);
			}
			else {
				return JSON.stringify(value);
			}
		};
		RequestPrettyPrinter.prototype.getObjVarNameBase=function (objectPathId) {
			var ret='v';
			var objPath=this.m_referencedObjectPaths[objectPathId];
			if (objPath) {
				switch (objPath.objectPathInfo.ObjectPathType) {
					case 1:
						ret=this.m_globalObjName;
						break;
					case 4:
						ret=Utility._toCamelLowerCase(objPath.objectPathInfo.Name);
						break;
					case 3:
						var methodName=objPath.objectPathInfo.Name;
						if (methodName.length > 3 && methodName.substr(0, 3)==='Get') {
							methodName=methodName.substr(3);
						}
						ret=Utility._toCamelLowerCase(methodName);
						break;
					case 5:
						var parentName=this.getObjVarNameBase(objPath.objectPathInfo.ParentObjectPathId);
						if (parentName.charAt(parentName.length - 1)==='s') {
							ret=parentName.substr(0, parentName.length - 1);
						}
						else {
							ret=parentName+'Item';
						}
						break;
				}
			}
			return ret;
		};
		RequestPrettyPrinter.prototype.getObjVarName=function (objectPathId) {
			if (this.m_variableNameForObjectPathMap[objectPathId]) {
				return this.m_variableNameForObjectPathMap[objectPathId];
			}
			var ret=this.getObjVarNameBase(objectPathId);
			if (!this.m_variableNameToObjectPathMap[ret]) {
				this.m_variableNameForObjectPathMap[objectPathId]=ret;
				this.m_variableNameToObjectPathMap[ret]=objectPathId;
				return ret;
			}
			var i=1;
			while (this.m_variableNameToObjectPathMap[ret+i.toString()]) {
				i++;
			}
			ret=ret+i.toString();
			this.m_variableNameForObjectPathMap[objectPathId]=ret;
			this.m_variableNameToObjectPathMap[ret]=objectPathId;
			return ret;
		};
		return RequestPrettyPrinter;
	}());
	var ResourceStrings=(function (_super) {
		__extends(ResourceStrings, _super);
		function ResourceStrings() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		ResourceStrings.cannotRegisterEvent='CannotRegisterEvent';
		ResourceStrings.connectionFailureWithStatus='ConnectionFailureWithStatus';
		ResourceStrings.connectionFailureWithDetails='ConnectionFailureWithDetails';
		ResourceStrings.propertyNotLoaded='PropertyNotLoaded';
		ResourceStrings.runMustReturnPromise='RunMustReturnPromise';
		ResourceStrings.moreInfoInnerError='MoreInfoInnerError';
		ResourceStrings.cannotApplyPropertyThroughSetMethod='CannotApplyPropertyThroughSetMethod';
		ResourceStrings.invalidOperationInCellEditMode='InvalidOperationInCellEditMode';
		ResourceStrings.objectIsUntracked='ObjectIsUntracked';
		ResourceStrings.customFunctionDefintionMissing='CustomFunctionDefintionMissing';
		ResourceStrings.customFunctionImplementationMissing='CustomFunctionImplementationMissing';
		ResourceStrings.customFunctionNameContainsBadChars='CustomFunctionNameContainsBadChars';
		ResourceStrings.customFunctionNameCannotSplit='CustomFunctionNameCannotSplit';
		ResourceStrings.customFunctionUnexpectedNumberOfEntriesInResultBatch='CustomFunctionUnexpectedNumberOfEntriesInResultBatch';
		ResourceStrings.customFunctionCancellationHandlerMissing='CustomFunctionCancellationHandlerMissing';
		ResourceStrings.customFunctionInvalidFunction='CustomFunctionInvalidFunction';
		ResourceStrings.customFunctionInvalidFunctionMapping='CustomFunctionInvalidFunctionMapping';
		ResourceStrings.customFunctionWindowMissing='CustomFunctionWindowMissing';
		ResourceStrings.customFunctionDefintionMissingOnWindow='CustomFunctionDefintionMissingOnWindow';
		ResourceStrings.pendingBatchInProgress='PendingBatchInProgress';
		ResourceStrings.notInsideBatch='NotInsideBatch';
		ResourceStrings.cannotUpdateReadOnlyProperty='CannotUpdateReadOnlyProperty';
		return ResourceStrings;
	}(CommonResourceStrings));
	OfficeExtension_1.ResourceStrings=ResourceStrings;
	CoreUtility.addResourceStringValues({
		CannotRegisterEvent: 'The event handler cannot be registered.',
		PropertyNotLoaded: "The property '{0}' is not available. Before reading the property's value, call the load method on the containing object and call \"context.sync()\" on the associated request context.",
		RunMustReturnPromise: 'The batch function passed to the ".run" method didn\'t return a promise. The function must return a promise, so that any automatically-tracked objects can be released at the completion of the batch operation. Typically, you return a promise by returning the response from "context.sync()".',
		InvalidOrTimedOutSessionMessage: 'Your Office Online session has expired or is invalid. To continue, refresh the page.',
		InvalidOperationInCellEditMode: 'Excel is in cell-editing mode. Please exit the edit mode by pressing ENTER or TAB or selecting another cell, and then try again.',
		CustomFunctionDefintionMissing: "A property with the name '{0}' that represents the function's definition must exist on Excel.Script.CustomFunctions.",
		CustomFunctionDefintionMissingOnWindow: "A property with the name '{0}' that represents the function's definition must exist on the window object.",
		CustomFunctionImplementationMissing: "The property with the name '{0}' on Excel.Script.CustomFunctions that represents the function's definition must contain a 'call' property that implements the function.",
		CustomFunctionNameContainsBadChars: 'The function name may only contain letters, digits, underscores, and periods.',
		CustomFunctionNameCannotSplit: 'The function name must contain a non-empty namespace and a non-empty short name.',
		CustomFunctionUnexpectedNumberOfEntriesInResultBatch: "The batching function returned a number of results that doesn't match the number of parameter value sets that were passed into it.",
		CustomFunctionCancellationHandlerMissing: 'The cancellation handler onCanceled is missing in the function. The handler must be present as the function is defined as cancelable.',
		CustomFunctionInvalidFunction: "The property with the name '{0}' that represents the function's definition is not a valid function.",
		CustomFunctionInvalidFunctionMapping: "The property with the name '{0}' on CustomFunctionMappings that represents the function's definition is not a valid function.",
		CustomFunctionWindowMissing: 'The window object was not found.',
		PendingBatchInProgress: 'There is a pending batch in progress. The batch method may not be called inside another batch, or simultaneously with another batch.',
		NotInsideBatch: 'Operations may not be invoked outside of a batch method.',
		CannotUpdateReadOnlyProperty: "The property '{0}' is read-only and it cannot be updated.",
		ObjectIsUntracked: 'The object is untracked.'
	});
	var Utility=(function (_super) {
		__extends(Utility, _super);
		function Utility() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Utility.fixObjectPathIfNecessary=function (clientObject, value) {
			if (clientObject && clientObject._objectPath && value) {
				clientObject._objectPath.updateUsingObjectData(value, clientObject);
			}
		};
		Utility.load=function (clientObj, option) {
			clientObj.context.load(clientObj, option);
			return clientObj;
		};
		Utility.loadAndSync=function (clientObj, option) {
			clientObj.context.load(clientObj, option);
			return clientObj.context.sync().then(function () { return clientObj; });
		};
		Utility.retrieve=function (clientObj, option) {
			var shouldPolyfill=OfficeExtension_1._internalConfig.alwaysPolyfillClientObjectRetrieveMethod;
			if (!shouldPolyfill) {
				shouldPolyfill=!Utility.isSetSupported('RichApiRuntime', '1.1');
			}
			var result=new RetrieveResultImpl(clientObj, shouldPolyfill);
			clientObj._retrieve(option, result);
			return result;
		};
		Utility.retrieveAndSync=function (clientObj, option) {
			var result=Utility.retrieve(clientObj, option);
			return clientObj.context.sync().then(function () { return result; });
		};
		Utility.toJson=function (clientObj, scalarProperties, navigationProperties, collectionItemsIfAny) {
			var result={};
			for (var prop in scalarProperties) {
				var value=scalarProperties[prop];
				if (typeof value !=='undefined') {
					result[prop]=value;
				}
			}
			for (var prop in navigationProperties) {
				var value=navigationProperties[prop];
				if (typeof value !=='undefined') {
					if (value[Utility.fieldName_isCollection] && typeof value[Utility.fieldName_m__items] !=='undefined') {
						result[prop]=value.toJSON()['items'];
					}
					else {
						result[prop]=value.toJSON();
					}
				}
			}
			if (collectionItemsIfAny) {
				result['items']=collectionItemsIfAny.map(function (item) { return item.toJSON(); });
			}
			return result;
		};
		Utility.throwError=function (resourceId, arg, errorLocation) {
			throw new _Internal.RuntimeError({
				code: resourceId,
				message: CoreUtility._getResourceString(resourceId, arg),
				debugInfo: errorLocation ? { errorLocation: errorLocation } : undefined
			});
		};
		Utility.createRuntimeError=function (code, message, location) {
			return new _Internal.RuntimeError({
				code: code,
				message: message,
				debugInfo: { errorLocation: location }
			});
		};
		Utility.throwIfNotLoaded=function (propertyName, fieldValue, entityName, isNull) {
			if (!isNull &&
				CoreUtility.isUndefined(fieldValue) &&
				propertyName.charCodeAt(0) !=Utility.s_underscoreCharCode) {
				throw Utility.createPropertyNotLoadedException(entityName, propertyName);
			}
		};
		Utility.createPropertyNotLoadedException=function (entityName, propertyName) {
			return new _Internal.RuntimeError({
				code: ErrorCodes.propertyNotLoaded,
				message: CoreUtility._getResourceString(ResourceStrings.propertyNotLoaded, propertyName),
				debugInfo: entityName ? { errorLocation: entityName+'.'+propertyName } : undefined
			});
		};
		Utility.createCannotUpdateReadOnlyPropertyException=function (entityName, propertyName) {
			return new _Internal.RuntimeError({
				code: ErrorCodes.cannotUpdateReadOnlyProperty,
				message: CoreUtility._getResourceString(ResourceStrings.cannotUpdateReadOnlyProperty, propertyName),
				debugInfo: entityName ? { errorLocation: entityName+'.'+propertyName } : undefined
			});
		};
		Utility.promisify=function (action) {
			return new OfficeExtension_1.Promise(function (resolve, reject) {
				var callback=function (result) {
					if (result.status=='failed') {
						reject(result.error);
					}
					else {
						resolve(result.value);
					}
				};
				action(callback);
			});
		};
		Utility._addActionResultHandler=function (clientObj, action, resultHandler) {
			clientObj.context._pendingRequest.addActionResultHandler(action, resultHandler);
		};
		Utility._handleNavigationPropertyResults=function (clientObj, objectValue, propertyNames) {
			for (var i=0; i < propertyNames.length - 1; i+=2) {
				if (!CoreUtility.isUndefined(objectValue[propertyNames[i+1]])) {
					clientObj[propertyNames[i]]._handleResult(objectValue[propertyNames[i+1]]);
				}
			}
		};
		Utility._fixupApiFlags=function (flags) {
			if (typeof flags==='boolean') {
				if (flags) {
					flags=1;
				}
				else {
					flags=0;
				}
			}
			return flags;
		};
		Utility.definePropertyThrowUnloadedException=function (obj, typeName, propertyName) {
			Object.defineProperty(obj, propertyName, {
				configurable: true,
				enumerable: true,
				get: function () {
					throw Utility.createPropertyNotLoadedException(typeName, propertyName);
				},
				set: function () {
					throw Utility.createCannotUpdateReadOnlyPropertyException(typeName, propertyName);
				}
			});
		};
		Utility.defineReadOnlyPropertyWithValue=function (obj, propertyName, value) {
			Object.defineProperty(obj, propertyName, {
				configurable: true,
				enumerable: true,
				get: function () {
					return value;
				},
				set: function () {
					throw Utility.createCannotUpdateReadOnlyPropertyException(null, propertyName);
				}
			});
		};
		Utility.processRetrieveResult=function (proxy, value, result, childItemCreateFunc) {
			if (CoreUtility.isNullOrUndefined(value)) {
				return;
			}
			if (childItemCreateFunc) {
				var data=value[Constants.itemsLowerCase];
				if (Array.isArray(data)) {
					var itemsResult=[];
					for (var i=0; i < data.length; i++) {
						var itemProxy=childItemCreateFunc(data[i], i);
						var itemResult={};
						itemResult[Constants.proxy]=itemProxy;
						itemProxy._handleRetrieveResult(data[i], itemResult);
						itemsResult.push(itemResult);
					}
					Utility.defineReadOnlyPropertyWithValue(result, Constants.itemsLowerCase, itemsResult);
				}
			}
			else {
				var scalarPropertyNames=proxy[Constants.scalarPropertyNames];
				var navigationPropertyNames=proxy[Constants.navigationPropertyNames];
				var typeName=proxy[Constants.className];
				if (scalarPropertyNames) {
					for (var i=0; i < scalarPropertyNames.length; i++) {
						var propName=scalarPropertyNames[i];
						var propValue=value[propName];
						if (CoreUtility.isUndefined(propValue)) {
							Utility.definePropertyThrowUnloadedException(result, typeName, propName);
						}
						else {
							Utility.defineReadOnlyPropertyWithValue(result, propName, propValue);
						}
					}
				}
				if (navigationPropertyNames) {
					for (var i=0; i < navigationPropertyNames.length; i++) {
						var propName=navigationPropertyNames[i];
						var propValue=value[propName];
						if (CoreUtility.isUndefined(propValue)) {
							Utility.definePropertyThrowUnloadedException(result, typeName, propName);
						}
						else {
							var propProxy=proxy[propName];
							var propResult={};
							propProxy._handleRetrieveResult(propValue, propResult);
							propResult[Constants.proxy]=propProxy;
							if (Array.isArray(propResult[Constants.itemsLowerCase])) {
								propResult=propResult[Constants.itemsLowerCase];
							}
							Utility.defineReadOnlyPropertyWithValue(result, propName, propResult);
						}
					}
				}
			}
		};
		Utility.fieldName_m__items='m__items';
		Utility.fieldName_isCollection='_isCollection';
		Utility._synchronousCleanup=false;
		Utility.s_underscoreCharCode='_'.charCodeAt(0);
		return Utility;
	}(CommonUtility));
	OfficeExtension_1.Utility=Utility;
	var BatchApiHelper=(function () {
		function BatchApiHelper() {
		}
		BatchApiHelper.invokeMethod=function (obj, methodName, operationType, args, flags, resultProcessType) {
			var action=ActionFactory.createMethodAction(obj.context, obj, methodName, operationType, args, flags);
			var result=new ClientResult(resultProcessType);
			Utility._addActionResultHandler(obj, action, result);
			return result;
		};
		BatchApiHelper.invokeEnsureUnchanged=function (obj, objectState) {
			ActionFactory.createEnsureUnchangedAction(obj.context, obj, objectState);
		};
		BatchApiHelper.invokeSetProperty=function (obj, propName, propValue, flags) {
			ActionFactory.createSetPropertyAction(obj.context, obj, propName, propValue, flags);
		};
		BatchApiHelper.createRootServiceObject=function (type, context) {
			var objectPath=ObjectPathFactory.createGlobalObjectObjectPath(context);
			return new type(context, objectPath);
		};
		BatchApiHelper.createObjectFromReferenceId=function (type, context, referenceId) {
			var objectPath=ObjectPathFactory.createReferenceIdObjectPath(context, referenceId);
			return new type(context, objectPath);
		};
		BatchApiHelper.createTopLevelServiceObject=function (type, context, typeName, isCollection, flags) {
			var objectPath=ObjectPathFactory.createNewObjectObjectPath(context, typeName, isCollection, flags);
			return new type(context, objectPath);
		};
		BatchApiHelper.createPropertyObject=function (type, parent, propertyName, isCollection, flags) {
			var objectPath=ObjectPathFactory.createPropertyObjectPath(parent.context, parent, propertyName, isCollection, false, flags);
			return new type(parent.context, objectPath);
		};
		BatchApiHelper.createIndexerObject=function (type, parent, args) {
			var objectPath=ObjectPathFactory.createIndexerObjectPath(parent.context, parent, args);
			return new type(parent.context, objectPath);
		};
		BatchApiHelper.createMethodObject=function (type, parent, methodName, operationType, args, isCollection, isInvalidAfterRequest, getByIdMethodName, flags) {
			var objectPath=ObjectPathFactory.createMethodObjectPath(parent.context, parent, methodName, operationType, args, isCollection, isInvalidAfterRequest, getByIdMethodName, flags);
			return new type(parent.context, objectPath);
		};
		BatchApiHelper.createChildItemObject=function (type, hasIndexerMethod, parent, chileItem, index) {
			var objectPath=ObjectPathFactory.createChildItemObjectPathUsingIndexerOrGetItemAt(hasIndexerMethod, parent.context, parent, chileItem, index);
			return new type(parent.context, objectPath);
		};
		return BatchApiHelper;
	}());
	OfficeExtension_1.BatchApiHelper=BatchApiHelper;
	var versionToken=1;
	var internalConfiguration={
		invokeRequestModifier: function (request) {
			request.DdaMethod.Version=versionToken;
			return request;
		},
		invokeResponseModifier: function (args) {
			versionToken=args.Version;
			if (args.Error) {
				args.error={};
				args.error.Code=args.Error;
			}
			return args;
		}
	};
	var CommunicationConstants;
	(function (CommunicationConstants) {
		CommunicationConstants["SendingId"]="sId";
		CommunicationConstants["RespondingId"]="rId";
		CommunicationConstants["CommandKey"]="command";
		CommunicationConstants["SessionInfoKey"]="sessionInfo";
		CommunicationConstants["ParamsKey"]="params";
		CommunicationConstants["ApiReadyCommand"]="apiready";
		CommunicationConstants["ExecuteMethodCommand"]="executeMethod";
		CommunicationConstants["GetAppContextCommand"]="getAppContext";
		CommunicationConstants["RegisterEventCommand"]="registerEvent";
		CommunicationConstants["UnregisterEventCommand"]="unregisterEvent";
		CommunicationConstants["FireEventCommand"]="fireEvent";
	})(CommunicationConstants || (CommunicationConstants={}));
	var EmbeddedConstants=(function () {
		function EmbeddedConstants() {
		}
		EmbeddedConstants.sessionContext='sc';
		EmbeddedConstants.embeddingPageOrigin='EmbeddingPageOrigin';
		EmbeddedConstants.embeddingPageSessionInfo='EmbeddingPageSessionInfo';
		return EmbeddedConstants;
	}());
	OfficeExtension_1.EmbeddedConstants=EmbeddedConstants;
	var EmbeddedSession=(function (_super) {
		__extends(EmbeddedSession, _super);
		function EmbeddedSession(url, options) {
			var _this=_super.call(this) || this;
			_this.m_chosenWindow=null;
			_this.m_chosenOrigin=null;
			_this.m_enabled=true;
			_this.m_onMessageHandler=_this._onMessage.bind(_this);
			_this.m_callbackList={};
			_this.m_id=0;
			_this.m_timeoutId=-1;
			_this.m_appContext=null;
			_this.m_url=url;
			_this.m_options=options;
			if (!_this.m_options) {
				_this.m_options={ sessionKey: Math.random().toString() };
			}
			if (!_this.m_options.sessionKey) {
				_this.m_options.sessionKey=Math.random().toString();
			}
			if (!_this.m_options.container) {
				_this.m_options.container=document.body;
			}
			if (!_this.m_options.timeoutInMilliseconds) {
				_this.m_options.timeoutInMilliseconds=60000;
			}
			if (!_this.m_options.height) {
				_this.m_options.height='400px';
			}
			if (!_this.m_options.width) {
				_this.m_options.width='100%';
			}
			if (!(_this.m_options.webApplication &&
				_this.m_options.webApplication.accessToken &&
				_this.m_options.webApplication.accessTokenTtl)) {
				_this.m_options.webApplication=null;
			}
			return _this;
		}
		EmbeddedSession.prototype._getIFrameSrc=function () {
			var origin=window.location.protocol+'//'+window.location.host;
			var toAppend=EmbeddedConstants.embeddingPageOrigin+				'='+				encodeURIComponent(origin)+				'&'+				EmbeddedConstants.embeddingPageSessionInfo+				'='+				encodeURIComponent(this.m_options.sessionKey);
			var useHash=false;
			if (this.m_url.toLowerCase().indexOf('/_layouts/preauth.aspx') > 0 ||
				this.m_url.toLowerCase().indexOf('/_layouts/15/preauth.aspx') > 0) {
				useHash=true;
			}
			var a=document.createElement('a');
			a.href=this.m_url;
			if (this.m_options.webApplication) {
				var toAppendWAC=EmbeddedConstants.embeddingPageOrigin+					'='+					origin+					'&'+					EmbeddedConstants.embeddingPageSessionInfo+					'='+					this.m_options.sessionKey;
				if (a.search.length===0 || a.search==='?') {
					a.search='?'+EmbeddedConstants.sessionContext+'='+encodeURIComponent(toAppendWAC);
				}
				else {
					a.search=a.search+'&'+EmbeddedConstants.sessionContext+'='+encodeURIComponent(toAppendWAC);
				}
			}
			else if (useHash) {
				if (a.hash.length===0 || a.hash==='#') {
					a.hash='#'+toAppend;
				}
				else {
					a.hash=a.hash+'&'+toAppend;
				}
			}
			else {
				if (a.search.length===0 || a.search==='?') {
					a.search='?'+toAppend;
				}
				else {
					a.search=a.search+'&'+toAppend;
				}
			}
			var iframeSrc=a.href;
			return iframeSrc;
		};
		EmbeddedSession.prototype.init=function () {
			var _this=this;
			window.addEventListener('message', this.m_onMessageHandler);
			var iframeSrc=this._getIFrameSrc();
			return CoreUtility.createPromise(function (resolve, reject) {
				var iframeElement=document.createElement('iframe');
				if (_this.m_options.id) {
					iframeElement.id=_this.m_options.id;
					iframeElement.name=_this.m_options.id;
				}
				iframeElement.style.height=_this.m_options.height;
				iframeElement.style.width=_this.m_options.width;
				if (!_this.m_options.webApplication) {
					iframeElement.src=iframeSrc;
					_this.m_options.container.appendChild(iframeElement);
				}
				else {
					var webApplicationForm=document.createElement('form');
					webApplicationForm.setAttribute('action', iframeSrc);
					webApplicationForm.setAttribute('method', 'post');
					webApplicationForm.setAttribute('target', iframeElement.name);
					_this.m_options.container.appendChild(webApplicationForm);
					var token_input=document.createElement('input');
					token_input.setAttribute('type', 'hidden');
					token_input.setAttribute('name', 'access_token');
					token_input.setAttribute('value', _this.m_options.webApplication.accessToken);
					webApplicationForm.appendChild(token_input);
					var token_ttl_input=document.createElement('input');
					token_ttl_input.setAttribute('type', 'hidden');
					token_ttl_input.setAttribute('name', 'access_token_ttl');
					token_ttl_input.setAttribute('value', _this.m_options.webApplication.accessTokenTtl);
					webApplicationForm.appendChild(token_ttl_input);
					_this.m_options.container.appendChild(iframeElement);
					webApplicationForm.submit();
				}
				_this.m_timeoutId=window.setTimeout(function () {
					_this.close();
					var err=Utility.createRuntimeError(CoreErrorCodes.timeout, CoreUtility._getResourceString(CoreResourceStrings.timeout), 'EmbeddedSession.init');
					reject(err);
				}, _this.m_options.timeoutInMilliseconds);
				_this.m_promiseResolver=resolve;
			});
		};
		EmbeddedSession.prototype._invoke=function (method, callback, params) {
			if (!this.m_enabled) {
				callback(5001, null);
				return;
			}
			if (internalConfiguration.invokeRequestModifier) {
				params=internalConfiguration.invokeRequestModifier(params);
			}
			this._sendMessageWithCallback(this.m_id++, method, params, function (args) {
				if (internalConfiguration.invokeResponseModifier) {
					args=internalConfiguration.invokeResponseModifier(args);
				}
				var errorCode=args['Error'];
				delete args['Error'];
				callback(errorCode || 0, args);
			});
		};
		EmbeddedSession.prototype.close=function () {
			window.removeEventListener('message', this.m_onMessageHandler);
			window.clearTimeout(this.m_timeoutId);
			this.m_enabled=false;
		};
		Object.defineProperty(EmbeddedSession.prototype, "eventRegistration", {
			get: function () {
				if (!this.m_sessionEventManager) {
					this.m_sessionEventManager=new EventRegistration(this._registerEventImpl.bind(this), this._unregisterEventImpl.bind(this));
				}
				return this.m_sessionEventManager;
			},
			enumerable: true,
			configurable: true
		});
		EmbeddedSession.prototype._createRequestExecutorOrNull=function () {
			return new EmbeddedRequestExecutor(this);
		};
		EmbeddedSession.prototype._resolveRequestUrlAndHeaderInfo=function () {
			return CoreUtility._createPromiseFromResult(null);
		};
		EmbeddedSession.prototype._registerEventImpl=function (eventId, targetId) {
			var _this=this;
			return CoreUtility.createPromise(function (resolve, reject) {
				_this._sendMessageWithCallback(_this.m_id++, CommunicationConstants.RegisterEventCommand, { EventId: eventId, TargetId: targetId }, function () {
					resolve(null);
				});
			});
		};
		EmbeddedSession.prototype._unregisterEventImpl=function (eventId, targetId) {
			var _this=this;
			return CoreUtility.createPromise(function (resolve, reject) {
				_this._sendMessageWithCallback(_this.m_id++, CommunicationConstants.UnregisterEventCommand, { EventId: eventId, TargetId: targetId }, function () {
					resolve();
				});
			});
		};
		EmbeddedSession.prototype._onMessage=function (event) {
			var _this=this;
			if (!this.m_enabled) {
				return;
			}
			if (this.m_chosenWindow && (this.m_chosenWindow !==event.source || this.m_chosenOrigin !==event.origin)) {
				return;
			}
			var eventData=event.data;
			if (eventData && eventData[CommunicationConstants.CommandKey]===CommunicationConstants.ApiReadyCommand) {
				if (!this.m_chosenWindow &&
					this._isValidDescendant(event.source) &&
					eventData[CommunicationConstants.SessionInfoKey]===this.m_options.sessionKey) {
					this.m_chosenWindow=event.source;
					this.m_chosenOrigin=event.origin;
					this._sendMessageWithCallback(this.m_id++, CommunicationConstants.GetAppContextCommand, null, function (appContext) {
						_this._setupContext(appContext);
						window.clearTimeout(_this.m_timeoutId);
						_this.m_promiseResolver();
					});
				}
				return;
			}
			if (eventData && eventData[CommunicationConstants.CommandKey]===CommunicationConstants.FireEventCommand) {
				var msg=eventData[CommunicationConstants.ParamsKey];
				var eventId=msg['EventId'];
				var targetId=msg['TargetId'];
				var data=msg['Data'];
				if (this.m_sessionEventManager) {
					var handlers=this.m_sessionEventManager.getHandlers(eventId, targetId);
					for (var i=0; i < handlers.length; i++) {
						handlers[i](data);
					}
				}
				return;
			}
			if (eventData && eventData.hasOwnProperty(CommunicationConstants.RespondingId)) {
				var rId=eventData[CommunicationConstants.RespondingId];
				var callback=this.m_callbackList[rId];
				if (typeof callback==='function') {
					callback(eventData[CommunicationConstants.ParamsKey]);
				}
				delete this.m_callbackList[rId];
			}
		};
		EmbeddedSession.prototype._sendMessageWithCallback=function (id, command, data, callback) {
			this.m_callbackList[id]=callback;
			var message={};
			message[CommunicationConstants.SendingId]=id;
			message[CommunicationConstants.CommandKey]=command;
			message[CommunicationConstants.ParamsKey]=data;
			this.m_chosenWindow.postMessage(JSON.stringify(message), this.m_chosenOrigin);
		};
		EmbeddedSession.prototype._isValidDescendant=function (wnd) {
			var container=this.m_options.container || document.body;
			function doesFrameWindow(containerWindow) {
				if (containerWindow===wnd) {
					return true;
				}
				for (var i=0, len=containerWindow.frames.length; i < len; i++) {
					if (doesFrameWindow(containerWindow.frames[i])) {
						return true;
					}
				}
				return false;
			}
			var iframes=container.getElementsByTagName('iframe');
			for (var i=0, len=iframes.length; i < len; i++) {
				if (doesFrameWindow(iframes[i].contentWindow)) {
					return true;
				}
			}
			return false;
		};
		EmbeddedSession.prototype._setupContext=function (appContext) {
			if (!(this.m_appContext=appContext)) {
				return;
			}
		};
		return EmbeddedSession;
	}(SessionBase));
	OfficeExtension_1.EmbeddedSession=EmbeddedSession;
	var EmbeddedRequestExecutor=(function () {
		function EmbeddedRequestExecutor(session) {
			this.m_session=session;
		}
		EmbeddedRequestExecutor.prototype.executeAsync=function (customData, requestFlags, requestMessage) {
			var _this=this;
			var messageSafearray=RichApiMessageUtility.buildMessageArrayForIRequestExecutor(customData, requestFlags, requestMessage, EmbeddedRequestExecutor.SourceLibHeaderValue);
			return CoreUtility.createPromise(function (resolve, reject) {
				_this.m_session._invoke(CommunicationConstants.ExecuteMethodCommand, function (status, result) {
					CoreUtility.log('Response:');
					CoreUtility.log(JSON.stringify(result));
					var response;
					if (status==0) {
						response=RichApiMessageUtility.buildResponseOnSuccess(RichApiMessageUtility.getResponseBodyFromSafeArray(result.Data), RichApiMessageUtility.getResponseHeadersFromSafeArray(result.Data));
					}
					else {
						response=RichApiMessageUtility.buildResponseOnError(result.error.Code, result.error.Message);
					}
					resolve(response);
				}, EmbeddedRequestExecutor._transformMessageArrayIntoParams(messageSafearray));
			});
		};
		EmbeddedRequestExecutor._transformMessageArrayIntoParams=function (msgArray) {
			return {
				ArrayData: msgArray,
				DdaMethod: {
					DispatchId: EmbeddedRequestExecutor.DispidExecuteRichApiRequestMethod
				}
			};
		};
		EmbeddedRequestExecutor.DispidExecuteRichApiRequestMethod=93;
		EmbeddedRequestExecutor.SourceLibHeaderValue='Embedded';
		return EmbeddedRequestExecutor;
	}());
})(OfficeExtension || (OfficeExtension={}));
var __extends=(this && this.__extends) || (function () {
	var extendStatics=Object.setPrototypeOf ||
		({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__=b; }) ||
		function (d, b) { for (var p in b)
			if (b.hasOwnProperty(p))
				d[p]=b[p]; };
	return function (d, b) {
		extendStatics(d, b);
		function __() { this.constructor=d; }
		d.prototype=b===null ? Object.create(b) : (__.prototype=b.prototype, new __());
	};
})();
var OfficeCore;
(function (OfficeCore) {
	var _hostName="OfficeCore";
	var _defaultApiSetName="AgaveVisualApi";
	var _createPropertyObject=OfficeExtension.BatchApiHelper.createPropertyObject;
	var _createMethodObject=OfficeExtension.BatchApiHelper.createMethodObject;
	var _createIndexerObject=OfficeExtension.BatchApiHelper.createIndexerObject;
	var _createRootServiceObject=OfficeExtension.BatchApiHelper.createRootServiceObject;
	var _createTopLevelServiceObject=OfficeExtension.BatchApiHelper.createTopLevelServiceObject;
	var _createChildItemObject=OfficeExtension.BatchApiHelper.createChildItemObject;
	var _invokeMethod=OfficeExtension.BatchApiHelper.invokeMethod;
	var _invokeEnsureUnchanged=OfficeExtension.BatchApiHelper.invokeEnsureUnchanged;
	var _invokeSetProperty=OfficeExtension.BatchApiHelper.invokeSetProperty;
	var _isNullOrUndefined=OfficeExtension.Utility.isNullOrUndefined;
	var _isUndefined=OfficeExtension.Utility.isUndefined;
	var _throwIfNotLoaded=OfficeExtension.Utility.throwIfNotLoaded;
	var _throwIfApiNotSupported=OfficeExtension.Utility.throwIfApiNotSupported;
	var _load=OfficeExtension.Utility.load;
	var _retrieve=OfficeExtension.Utility.retrieve;
	var _toJson=OfficeExtension.Utility.toJson;
	var _fixObjectPathIfNecessary=OfficeExtension.Utility.fixObjectPathIfNecessary;
	var _handleNavigationPropertyResults=OfficeExtension.Utility._handleNavigationPropertyResults;
	var _adjustToDateTime=OfficeExtension.Utility.adjustToDateTime;
	var _processRetrieveResult=OfficeExtension.Utility.processRetrieveResult;
	var _typeBiShim="BiShim";
	var BiShim=(function (_super) {
		__extends(BiShim, _super);
		function BiShim() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(BiShim.prototype, "_className", {
			get: function () {
				return "BiShim";
			},
			enumerable: true,
			configurable: true
		});
		BiShim.prototype.initialize=function (capabilities) {
			_invokeMethod(this, "Initialize", 0, [capabilities], 0, 0);
		};
		BiShim.prototype.getData=function () {
			return _invokeMethod(this, "getData", 1, [], 4, 0);
		};
		BiShim.prototype.setVisualObjects=function (visualObjects) {
			_invokeMethod(this, "setVisualObjects", 0, [visualObjects], 2, 0);
		};
		BiShim.prototype.setVisualObjectsToPersist=function (visualObjectsToPersist) {
			_invokeMethod(this, "setVisualObjectsToPersist", 0, [visualObjectsToPersist], 2, 0);
		};
		BiShim.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
		};
		BiShim.prototype._handleRetrieveResult=function (value, result) {
			_super.prototype._handleRetrieveResult.call(this, value, result);
			_processRetrieveResult(this, value, result);
		};
		BiShim.newObject=function (context) {
			return _createTopLevelServiceObject(OfficeCore.BiShim, context, "Microsoft.AgaveVisual.BiShim", false, 4);
		};
		BiShim.prototype.toJSON=function () {
			return _toJson(this, {}, {});
		};
		return BiShim;
	}(OfficeExtension.ClientObject));
	OfficeCore.BiShim=BiShim;
	var AgaveVisualErrorCodes;
	(function (AgaveVisualErrorCodes) {
		AgaveVisualErrorCodes["generalException"]="GeneralException";
	})(AgaveVisualErrorCodes=OfficeCore.AgaveVisualErrorCodes || (OfficeCore.AgaveVisualErrorCodes={}));
})(OfficeCore || (OfficeCore={}));
var OfficeCore;
(function (OfficeCore) {
	var _hostName="OfficeCore";
	var _defaultApiSetName="ExperimentApi";
	var _createPropertyObject=OfficeExtension.BatchApiHelper.createPropertyObject;
	var _createMethodObject=OfficeExtension.BatchApiHelper.createMethodObject;
	var _createIndexerObject=OfficeExtension.BatchApiHelper.createIndexerObject;
	var _createRootServiceObject=OfficeExtension.BatchApiHelper.createRootServiceObject;
	var _createTopLevelServiceObject=OfficeExtension.BatchApiHelper.createTopLevelServiceObject;
	var _createChildItemObject=OfficeExtension.BatchApiHelper.createChildItemObject;
	var _invokeMethod=OfficeExtension.BatchApiHelper.invokeMethod;
	var _invokeEnsureUnchanged=OfficeExtension.BatchApiHelper.invokeEnsureUnchanged;
	var _invokeSetProperty=OfficeExtension.BatchApiHelper.invokeSetProperty;
	var _isNullOrUndefined=OfficeExtension.Utility.isNullOrUndefined;
	var _isUndefined=OfficeExtension.Utility.isUndefined;
	var _throwIfNotLoaded=OfficeExtension.Utility.throwIfNotLoaded;
	var _throwIfApiNotSupported=OfficeExtension.Utility.throwIfApiNotSupported;
	var _load=OfficeExtension.Utility.load;
	var _retrieve=OfficeExtension.Utility.retrieve;
	var _toJson=OfficeExtension.Utility.toJson;
	var _fixObjectPathIfNecessary=OfficeExtension.Utility.fixObjectPathIfNecessary;
	var _handleNavigationPropertyResults=OfficeExtension.Utility._handleNavigationPropertyResults;
	var _adjustToDateTime=OfficeExtension.Utility.adjustToDateTime;
	var _processRetrieveResult=OfficeExtension.Utility.processRetrieveResult;
	var _typeFlightingService="FlightingService";
	var FlightingService=(function (_super) {
		__extends(FlightingService, _super);
		function FlightingService() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(FlightingService.prototype, "_className", {
			get: function () {
				return "FlightingService";
			},
			enumerable: true,
			configurable: true
		});
		FlightingService.prototype.getClientSessionId=function () {
			return _invokeMethod(this, "GetClientSessionId", 1, [], 4, 0);
		};
		FlightingService.prototype.getDeferredFlights=function () {
			return _invokeMethod(this, "GetDeferredFlights", 1, [], 4, 0);
		};
		FlightingService.prototype.getFeature=function (featureName, type, defaultValue, possibleValues) {
			return _createMethodObject(OfficeCore.ABType, this, "GetFeature", 1, [featureName, type, defaultValue, possibleValues], false, false, null, 4);
		};
		FlightingService.prototype.getFeatureGate=function (featureName, scope) {
			return _createMethodObject(OfficeCore.ABType, this, "GetFeatureGate", 1, [featureName, scope], false, false, null, 4);
		};
		FlightingService.prototype.resetOverride=function (featureName) {
			_invokeMethod(this, "ResetOverride", 0, [featureName], 0, 0);
		};
		FlightingService.prototype.setOverride=function (featureName, type, value) {
			_invokeMethod(this, "SetOverride", 0, [featureName, type, value], 0, 0);
		};
		FlightingService.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
		};
		FlightingService.prototype._handleRetrieveResult=function (value, result) {
			_super.prototype._handleRetrieveResult.call(this, value, result);
			_processRetrieveResult(this, value, result);
		};
		FlightingService.newObject=function (context) {
			return _createTopLevelServiceObject(OfficeCore.FlightingService, context, "Microsoft.Experiment.FlightingService", false, 4);
		};
		FlightingService.prototype.toJSON=function () {
			return _toJson(this, {}, {});
		};
		return FlightingService;
	}(OfficeExtension.ClientObject));
	OfficeCore.FlightingService=FlightingService;
	var _typeABType="ABType";
	var ABType=(function (_super) {
		__extends(ABType, _super);
		function ABType() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(ABType.prototype, "_className", {
			get: function () {
				return "ABType";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ABType.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["value"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ABType.prototype, "value", {
			get: function () {
				_throwIfNotLoaded("value", this._V, _typeABType, this._isNull);
				return this._V;
			},
			enumerable: true,
			configurable: true
		});
		ABType.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["Value"])) {
				this._V=obj["Value"];
			}
		};
		ABType.prototype.load=function (option) {
			return _load(this, option);
		};
		ABType.prototype.retrieve=function (option) {
			return _retrieve(this, option);
		};
		ABType.prototype._handleRetrieveResult=function (value, result) {
			_super.prototype._handleRetrieveResult.call(this, value, result);
			_processRetrieveResult(this, value, result);
		};
		ABType.prototype.toJSON=function () {
			return _toJson(this, {
				"value": this._V
			}, {});
		};
		ABType.prototype.ensureUnchanged=function (data) {
			_invokeEnsureUnchanged(this, data);
			return;
		};
		return ABType;
	}(OfficeExtension.ClientObject));
	OfficeCore.ABType=ABType;
	var FeatureType;
	(function (FeatureType) {
		FeatureType["boolean"]="Boolean";
		FeatureType["integer"]="Integer";
		FeatureType["string"]="String";
	})(FeatureType=OfficeCore.FeatureType || (OfficeCore.FeatureType={}));
	var ExperimentErrorCodes;
	(function (ExperimentErrorCodes) {
		ExperimentErrorCodes["generalException"]="GeneralException";
	})(ExperimentErrorCodes=OfficeCore.ExperimentErrorCodes || (OfficeCore.ExperimentErrorCodes={}));
})(OfficeCore || (OfficeCore={}));
var OfficeCore;
(function (OfficeCore) {
	OfficeCore.OfficeOnlineDomainList=[
		"*.dod.online.office365.us",
		"*.gov.online.office365.us",
		"*.officeapps-df.live.com",
		"*.officeapps.live.com",
		"*.online.office.de",
		"*.partner.officewebapps.cn"
	];
	function isHostOriginTrusted() {
		if (typeof window.external==='undefined' ||
			typeof window.external.GetContext==='undefined') {
			var hostUrl=OSF.getClientEndPoint()._targetUrl;
			var hostname_1=getHostNameFromUrl(hostUrl);
			if (hostUrl.indexOf("https:") !=0) {
				return false;
			}
			OfficeCore.OfficeOnlineDomainList.forEach(function (domain) {
				if (domain.indexOf("*.")==0) {
					domain=domain.substring(2);
				}
				if (hostname_1.indexOf(domain)==hostname_1.length - domain.length) {
					return true;
				}
			});
			return false;
		}
		return true;
	}
	OfficeCore.isHostOriginTrusted=isHostOriginTrusted;
	function getHostNameFromUrl(url) {
		var hostName="";
		hostName=url.split("/")[2];
		hostName=hostName.split(":")[0];
		hostName=hostName.split("?")[0];
		return hostName;
	}
})(OfficeCore || (OfficeCore={}));
var OfficeCore;
(function (OfficeCore) {
	var FirstPartyApis=(function () {
		function FirstPartyApis(context) {
			this.context=context;
		}
		Object.defineProperty(FirstPartyApis.prototype, "roamingSettings", {
			get: function () {
				if (!this.m_roamingSettings) {
					this.m_roamingSettings=OfficeCore.AuthenticationService.newObject(this.context).roamingSettings;
				}
				return this.m_roamingSettings;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(FirstPartyApis.prototype, "tap", {
			get: function () {
				if (!this.m_tap) {
					this.m_tap=OfficeCore.Tap.newObject(this.context);
				}
				return this.m_tap;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(FirstPartyApis.prototype, "skill", {
			get: function () {
				if (!this.m_skill) {
					this.m_skill=OfficeCore.Skill.newObject(this.context);
				}
				return this.m_skill;
			},
			enumerable: true,
			configurable: true
		});
		return FirstPartyApis;
	}());
	OfficeCore.FirstPartyApis=FirstPartyApis;
	var RequestContext=(function (_super) {
		__extends(RequestContext, _super);
		function RequestContext(url) {
			return _super.call(this, url) || this;
		}
		Object.defineProperty(RequestContext.prototype, "firstParty", {
			get: function () {
				if (!this.m_firstPartyApis) {
					this.m_firstPartyApis=new FirstPartyApis(this);
				}
				return this.m_firstPartyApis;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(RequestContext.prototype, "flighting", {
			get: function () {
				return this.flightingService;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(RequestContext.prototype, "telemetry", {
			get: function () {
				if (!this.m_telemetry) {
					this.m_telemetry=OfficeCore.TelemetryService.newObject(this);
				}
				return this.m_telemetry;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(RequestContext.prototype, "bi", {
			get: function () {
				if (!this.m_biShim) {
					this.m_biShim=OfficeCore.BiShim.newObject(this);
				}
				return this.m_biShim;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(RequestContext.prototype, "flightingService", {
			get: function () {
				if (!this.m_flightingService) {
					this.m_flightingService=OfficeCore.FlightingService.newObject(this);
				}
				return this.m_flightingService;
			},
			enumerable: true,
			configurable: true
		});
		return RequestContext;
	}(OfficeExtension.ClientRequestContext));
	OfficeCore.RequestContext=RequestContext;
	function run(arg1, arg2) {
		return OfficeExtension.ClientRequestContext._runBatch("OfficeCore.run", arguments, function (requestInfo) { return new OfficeCore.RequestContext(requestInfo); });
	}
	OfficeCore.run=run;
})(OfficeCore || (OfficeCore={}));
var OfficeCore;
(function (OfficeCore) {
	var _hostName="Office";
	var _defaultApiSetName="OfficeSharedApi";
	var _createPropertyObject=OfficeExtension.BatchApiHelper.createPropertyObject;
	var _createMethodObject=OfficeExtension.BatchApiHelper.createMethodObject;
	var _createIndexerObject=OfficeExtension.BatchApiHelper.createIndexerObject;
	var _createRootServiceObject=OfficeExtension.BatchApiHelper.createRootServiceObject;
	var _createTopLevelServiceObject=OfficeExtension.BatchApiHelper.createTopLevelServiceObject;
	var _createChildItemObject=OfficeExtension.BatchApiHelper.createChildItemObject;
	var _invokeMethod=OfficeExtension.BatchApiHelper.invokeMethod;
	var _invokeEnsureUnchanged=OfficeExtension.BatchApiHelper.invokeEnsureUnchanged;
	var _invokeSetProperty=OfficeExtension.BatchApiHelper.invokeSetProperty;
	var _isNullOrUndefined=OfficeExtension.Utility.isNullOrUndefined;
	var _isUndefined=OfficeExtension.Utility.isUndefined;
	var _throwIfNotLoaded=OfficeExtension.Utility.throwIfNotLoaded;
	var _throwIfApiNotSupported=OfficeExtension.Utility.throwIfApiNotSupported;
	var _load=OfficeExtension.Utility.load;
	var _retrieve=OfficeExtension.Utility.retrieve;
	var _toJson=OfficeExtension.Utility.toJson;
	var _fixObjectPathIfNecessary=OfficeExtension.Utility.fixObjectPathIfNecessary;
	var _handleNavigationPropertyResults=OfficeExtension.Utility._handleNavigationPropertyResults;
	var _adjustToDateTime=OfficeExtension.Utility.adjustToDateTime;
	var _processRetrieveResult=OfficeExtension.Utility.processRetrieveResult;
	var _typeSkill="Skill";
	var Skill=(function (_super) {
		__extends(Skill, _super);
		function Skill() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(Skill.prototype, "_className", {
			get: function () {
				return "Skill";
			},
			enumerable: true,
			configurable: true
		});
		Skill.prototype.executeAction=function (paneId, actionId, actionDescriptor) {
			return _invokeMethod(this, "ExecuteAction", 1, [paneId, actionId, actionDescriptor], 4 | 1, 0);
		};
		Skill.prototype.notifyPaneEvent=function (paneId, eventDescriptor) {
			_invokeMethod(this, "NotifyPaneEvent", 1, [paneId, eventDescriptor], 4 | 1, 0);
		};
		Skill.prototype.registerHostSkillEvent=function () {
			_invokeMethod(this, "RegisterHostSkillEvent", 0, [], 1, 0);
		};
		Skill.prototype.testFireEvent=function () {
			_invokeMethod(this, "TestFireEvent", 0, [], 1, 0);
		};
		Skill.prototype.unregisterHostSkillEvent=function () {
			_invokeMethod(this, "UnregisterHostSkillEvent", 0, [], 1, 0);
		};
		Skill.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
		};
		Skill.prototype._handleRetrieveResult=function (value, result) {
			_super.prototype._handleRetrieveResult.call(this, value, result);
			_processRetrieveResult(this, value, result);
		};
		Skill.newObject=function (context) {
			return _createTopLevelServiceObject(OfficeCore.Skill, context, "Microsoft.SkillApi.Skill", false, 4);
		};
		Object.defineProperty(Skill.prototype, "onHostSkillEvent", {
			get: function () {
				var _this=this;
				if (!this.m_hostSkillEvent) {
					this.m_hostSkillEvent=new OfficeExtension.GenericEventHandlers(this.context, this, "HostSkillEvent", {
						eventType: 65538,
						registerFunc: function () { return _this.registerHostSkillEvent(); },
						unregisterFunc: function () { return _this.unregisterHostSkillEvent(); },
						getTargetIdFunc: function () { return null; },
						eventArgsTransformFunc: function (args) {
							var transformedArgs={
								type: args.type,
								data: args.data
							};
							return OfficeExtension.Utility._createPromiseFromResult(transformedArgs);
						}
					});
				}
				return this.m_hostSkillEvent;
			},
			enumerable: true,
			configurable: true
		});
		Skill.prototype.toJSON=function () {
			return _toJson(this, {}, {});
		};
		return Skill;
	}(OfficeExtension.ClientObject));
	OfficeCore.Skill=Skill;
	var SkillErrorCodes;
	(function (SkillErrorCodes) {
		SkillErrorCodes["generalException"]="GeneralException";
	})(SkillErrorCodes=OfficeCore.SkillErrorCodes || (OfficeCore.SkillErrorCodes={}));
})(OfficeCore || (OfficeCore={}));
var OfficeCore;
(function (OfficeCore) {
	var _hostName="OfficeCore";
	var _defaultApiSetName="TelemetryApi";
	var _createPropertyObject=OfficeExtension.BatchApiHelper.createPropertyObject;
	var _createMethodObject=OfficeExtension.BatchApiHelper.createMethodObject;
	var _createIndexerObject=OfficeExtension.BatchApiHelper.createIndexerObject;
	var _createRootServiceObject=OfficeExtension.BatchApiHelper.createRootServiceObject;
	var _createTopLevelServiceObject=OfficeExtension.BatchApiHelper.createTopLevelServiceObject;
	var _createChildItemObject=OfficeExtension.BatchApiHelper.createChildItemObject;
	var _invokeMethod=OfficeExtension.BatchApiHelper.invokeMethod;
	var _invokeEnsureUnchanged=OfficeExtension.BatchApiHelper.invokeEnsureUnchanged;
	var _invokeSetProperty=OfficeExtension.BatchApiHelper.invokeSetProperty;
	var _isNullOrUndefined=OfficeExtension.Utility.isNullOrUndefined;
	var _isUndefined=OfficeExtension.Utility.isUndefined;
	var _throwIfNotLoaded=OfficeExtension.Utility.throwIfNotLoaded;
	var _throwIfApiNotSupported=OfficeExtension.Utility.throwIfApiNotSupported;
	var _load=OfficeExtension.Utility.load;
	var _retrieve=OfficeExtension.Utility.retrieve;
	var _toJson=OfficeExtension.Utility.toJson;
	var _fixObjectPathIfNecessary=OfficeExtension.Utility.fixObjectPathIfNecessary;
	var _handleNavigationPropertyResults=OfficeExtension.Utility._handleNavigationPropertyResults;
	var _adjustToDateTime=OfficeExtension.Utility.adjustToDateTime;
	var _processRetrieveResult=OfficeExtension.Utility.processRetrieveResult;
	var _typeTelemetryService="TelemetryService";
	var TelemetryService=(function (_super) {
		__extends(TelemetryService, _super);
		function TelemetryService() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(TelemetryService.prototype, "_className", {
			get: function () {
				return "TelemetryService";
			},
			enumerable: true,
			configurable: true
		});
		TelemetryService.prototype.sendTelemetryEvent=function (telemetryProperties, eventName, eventContract, eventFlags, value) {
			_invokeMethod(this, "SendTelemetryEvent", 1, [telemetryProperties, eventName, eventContract, eventFlags, value], 4, 0);
		};
		TelemetryService.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
		};
		TelemetryService.prototype._handleRetrieveResult=function (value, result) {
			_super.prototype._handleRetrieveResult.call(this, value, result);
			_processRetrieveResult(this, value, result);
		};
		TelemetryService.newObject=function (context) {
			return _createTopLevelServiceObject(OfficeCore.TelemetryService, context, "Microsoft.Telemetry.TelemetryService", false, 4);
		};
		TelemetryService.prototype.toJSON=function () {
			return _toJson(this, {}, {});
		};
		return TelemetryService;
	}(OfficeExtension.ClientObject));
	OfficeCore.TelemetryService=TelemetryService;
	var DataFieldType;
	(function (DataFieldType) {
		DataFieldType["unset"]="Unset";
		DataFieldType["string"]="String";
		DataFieldType["boolean"]="Boolean";
		DataFieldType["int64"]="Int64";
		DataFieldType["double"]="Double";
	})(DataFieldType=OfficeCore.DataFieldType || (OfficeCore.DataFieldType={}));
	var TelemetryErrorCodes;
	(function (TelemetryErrorCodes) {
		TelemetryErrorCodes["generalException"]="GeneralException";
	})(TelemetryErrorCodes=OfficeCore.TelemetryErrorCodes || (OfficeCore.TelemetryErrorCodes={}));
})(OfficeCore || (OfficeCore={}));
var OfficeFirstPartyAuth;
(function (OfficeFirstPartyAuth) {
	function getAccessToken(options) {
		var context=new OfficeCore.RequestContext();
		var auth=OfficeCore.AuthenticationService.newObject(context);
		context._customData="WacPartition";
		if (OSF._OfficeAppFactory.getHostInfo().hostPlatform=="web" && OSF._OfficeAppFactory.getHostInfo().hostType=="word") {
			var result_1=auth.getAccessToken(options, null);
			return context.sync().then(function () { return result_1.value; });
		}
		else {
			return new OfficeExtension.CoreUtility.Promise(function (resolve, reject) {
				var handler=auth.onTokenReceived.add(function (arg) {
					if (!OfficeExtension.CoreUtility.isNullOrUndefined(arg)) {
						handler.remove();
						context.sync()["catch"](function () {
						});
						if (arg.code==0) {
							resolve(arg.tokenValue);
						}
						else {
							if (OfficeExtension.CoreUtility.isNullOrUndefined(arg.errorInfo)) {
								reject({ code: arg.code });
							}
							else {
								try {
									reject(JSON.parse(arg.errorInfo));
								}
								catch (e) {
									reject({ code: arg.code, message: arg.errorInfo });
								}
							}
						}
					}
					return null;
				});
				context.sync()
					.then(function () {
					var apiResult=auth.getAccessToken(options, auth._targetId);
					return context.sync()
						.then(function () {
						if (OfficeExtension.CoreUtility.isNullOrUndefined(apiResult.value)) {
							return null;
						}
						var tokenValue=apiResult.value.accessToken;
						if (!OfficeExtension.CoreUtility.isNullOrUndefined(tokenValue)) {
							resolve(apiResult.value);
						}
					});
				})["catch"](function (e) {
					reject(e);
				});
			});
		}
	}
	OfficeFirstPartyAuth.getAccessToken=getAccessToken;
	function getPrimaryIdentityInfo() {
		var context=new OfficeCore.RequestContext();
		var auth=OfficeCore.AuthenticationService.newObject(context);
		context._customData="WacPartition";
		var result=auth.getPrimaryIdentityInfo();
		return context.sync().then(function () { return result.value; });
	}
	OfficeFirstPartyAuth.getPrimaryIdentityInfo=getPrimaryIdentityInfo;
})(OfficeFirstPartyAuth || (OfficeFirstPartyAuth={}));
var OfficeCore;
(function (OfficeCore) {
	var _hostName="Office";
	var _defaultApiSetName="OfficeSharedApi";
	var _createPropertyObject=OfficeExtension.BatchApiHelper.createPropertyObject;
	var _createMethodObject=OfficeExtension.BatchApiHelper.createMethodObject;
	var _createIndexerObject=OfficeExtension.BatchApiHelper.createIndexerObject;
	var _createRootServiceObject=OfficeExtension.BatchApiHelper.createRootServiceObject;
	var _createTopLevelServiceObject=OfficeExtension.BatchApiHelper.createTopLevelServiceObject;
	var _createChildItemObject=OfficeExtension.BatchApiHelper.createChildItemObject;
	var _invokeMethod=OfficeExtension.BatchApiHelper.invokeMethod;
	var _invokeEnsureUnchanged=OfficeExtension.BatchApiHelper.invokeEnsureUnchanged;
	var _invokeSetProperty=OfficeExtension.BatchApiHelper.invokeSetProperty;
	var _isNullOrUndefined=OfficeExtension.Utility.isNullOrUndefined;
	var _isUndefined=OfficeExtension.Utility.isUndefined;
	var _throwIfNotLoaded=OfficeExtension.Utility.throwIfNotLoaded;
	var _throwIfApiNotSupported=OfficeExtension.Utility.throwIfApiNotSupported;
	var _load=OfficeExtension.Utility.load;
	var _retrieve=OfficeExtension.Utility.retrieve;
	var _toJson=OfficeExtension.Utility.toJson;
	var _fixObjectPathIfNecessary=OfficeExtension.Utility.fixObjectPathIfNecessary;
	var _handleNavigationPropertyResults=OfficeExtension.Utility._handleNavigationPropertyResults;
	var _adjustToDateTime=OfficeExtension.Utility.adjustToDateTime;
	var _processRetrieveResult=OfficeExtension.Utility.processRetrieveResult;
	var IdentityType;
	(function (IdentityType) {
		IdentityType["organizationAccount"]="OrganizationAccount";
		IdentityType["microsoftAccount"]="MicrosoftAccount";
	})(IdentityType=OfficeCore.IdentityType || (OfficeCore.IdentityType={}));
	var _typeAuthenticationService="AuthenticationService";
	var AuthenticationService=(function (_super) {
		__extends(AuthenticationService, _super);
		function AuthenticationService() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(AuthenticationService.prototype, "_className", {
			get: function () {
				return "AuthenticationService";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(AuthenticationService.prototype, "_navigationPropertyNames", {
			get: function () {
				return ["roamingSettings"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(AuthenticationService.prototype, "_targetId", {
			get: function () {
				if (this.m_targetId==undefined) {
					if (typeof (OSF) !=='undefined' && OSF.OUtil) {
						this.m_targetId=OSF.OUtil.Guid.generateNewGuid();
					}
					else {
						this.m_targetId=""+this.context._nextId();
					}
				}
				return this.m_targetId;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(AuthenticationService.prototype, "roamingSettings", {
			get: function () {
				if (!this._R) {
					this._R=_createPropertyObject(OfficeCore.RoamingSettingCollection, this, "RoamingSettings", false, 4);
				}
				return this._R;
			},
			enumerable: true,
			configurable: true
		});
		AuthenticationService.prototype.getAccessToken=function (tokenParameters, targetId) {
			return _invokeMethod(this, "GetAccessToken", 1, [tokenParameters, targetId], 4 | 1, 0);
		};
		AuthenticationService.prototype.getPrimaryIdentityInfo=function () {
			_throwIfApiNotSupported("AuthenticationService.getPrimaryIdentityInfo", "FirstPartyAuthentication", "1.2", _hostName);
			return _invokeMethod(this, "GetPrimaryIdentityInfo", 1, [], 4 | 1, 0);
		};
		AuthenticationService.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			_handleNavigationPropertyResults(this, obj, ["roamingSettings", "RoamingSettings"]);
		};
		AuthenticationService.prototype.load=function (option) {
			return _load(this, option);
		};
		AuthenticationService.prototype.retrieve=function (option) {
			return _retrieve(this, option);
		};
		AuthenticationService.prototype._handleRetrieveResult=function (value, result) {
			_super.prototype._handleRetrieveResult.call(this, value, result);
			_processRetrieveResult(this, value, result);
		};
		AuthenticationService.newObject=function (context) {
			return _createTopLevelServiceObject(OfficeCore.AuthenticationService, context, "Microsoft.Authentication.AuthenticationService", false, 4);
		};
		Object.defineProperty(AuthenticationService.prototype, "onTokenReceived", {
			get: function () {
				var _this=this;
				_throwIfApiNotSupported("AuthenticationService.onTokenReceived", "FirstPartyAuthentication", "1.2", _hostName);
				if (!this.m_tokenReceived) {
					this.m_tokenReceived=new OfficeExtension.GenericEventHandlers(this.context, this, "TokenReceived", {
						eventType: 3001,
						registerFunc: function () { },
						unregisterFunc: function () { },
						getTargetIdFunc: function () { return _this._targetId; },
						eventArgsTransformFunc: function (value) {
							var newArgs={
								tokenValue: value.tokenValue,
								code: value.code,
								errorInfo: value.errorInfo
							};
							return OfficeExtension.Utility._createPromiseFromResult(newArgs);
						}
					});
				}
				return this.m_tokenReceived;
			},
			enumerable: true,
			configurable: true
		});
		AuthenticationService.prototype.toJSON=function () {
			return _toJson(this, {}, {});
		};
		return AuthenticationService;
	}(OfficeExtension.ClientObject));
	OfficeCore.AuthenticationService=AuthenticationService;
	var _typeRoamingSetting="RoamingSetting";
	var RoamingSetting=(function (_super) {
		__extends(RoamingSetting, _super);
		function RoamingSetting() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(RoamingSetting.prototype, "_className", {
			get: function () {
				return "RoamingSetting";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(RoamingSetting.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["id", "value"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(RoamingSetting.prototype, "_scalarPropertyUpdateable", {
			get: function () {
				return [false, true];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(RoamingSetting.prototype, "id", {
			get: function () {
				_throwIfNotLoaded("id", this._I, _typeRoamingSetting, this._isNull);
				return this._I;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(RoamingSetting.prototype, "value", {
			get: function () {
				_throwIfNotLoaded("value", this._V, _typeRoamingSetting, this._isNull);
				return this._V;
			},
			set: function (value) {
				this._V=value;
				_invokeSetProperty(this, "Value", value, 0);
			},
			enumerable: true,
			configurable: true
		});
		RoamingSetting.prototype.set=function (properties, options) {
			this._recursivelySet(properties, options, ["value"], [], []);
		};
		RoamingSetting.prototype.update=function (properties) {
			this._recursivelyUpdate(properties);
		};
		RoamingSetting.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["Id"])) {
				this._I=obj["Id"];
			}
			if (!_isUndefined(obj["Value"])) {
				this._V=obj["Value"];
			}
		};
		RoamingSetting.prototype.load=function (option) {
			return _load(this, option);
		};
		RoamingSetting.prototype.retrieve=function (option) {
			return _retrieve(this, option);
		};
		RoamingSetting.prototype._handleIdResult=function (value) {
			_super.prototype._handleIdResult.call(this, value);
			if (_isNullOrUndefined(value)) {
				return;
			}
			if (!_isUndefined(value["Id"])) {
				this._I=value["Id"];
			}
		};
		RoamingSetting.prototype._handleRetrieveResult=function (value, result) {
			_super.prototype._handleRetrieveResult.call(this, value, result);
			_processRetrieveResult(this, value, result);
		};
		RoamingSetting.prototype.toJSON=function () {
			return _toJson(this, {
				"id": this._I,
				"value": this._V
			}, {});
		};
		RoamingSetting.prototype.ensureUnchanged=function (data) {
			_invokeEnsureUnchanged(this, data);
			return;
		};
		return RoamingSetting;
	}(OfficeExtension.ClientObject));
	OfficeCore.RoamingSetting=RoamingSetting;
	var _typeRoamingSettingCollection="RoamingSettingCollection";
	var RoamingSettingCollection=(function (_super) {
		__extends(RoamingSettingCollection, _super);
		function RoamingSettingCollection() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(RoamingSettingCollection.prototype, "_className", {
			get: function () {
				return "RoamingSettingCollection";
			},
			enumerable: true,
			configurable: true
		});
		RoamingSettingCollection.prototype.getItem=function (id) {
			return _createMethodObject(OfficeCore.RoamingSetting, this, "GetItem", 1, [id], false, false, null, 4);
		};
		RoamingSettingCollection.prototype.getItemOrNullObject=function (id) {
			return _createMethodObject(OfficeCore.RoamingSetting, this, "GetItemOrNullObject", 1, [id], false, false, null, 4);
		};
		RoamingSettingCollection.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
		};
		RoamingSettingCollection.prototype._handleRetrieveResult=function (value, result) {
			_super.prototype._handleRetrieveResult.call(this, value, result);
			_processRetrieveResult(this, value, result);
		};
		RoamingSettingCollection.prototype.toJSON=function () {
			return _toJson(this, {}, {});
		};
		return RoamingSettingCollection;
	}(OfficeExtension.ClientObject));
	OfficeCore.RoamingSettingCollection=RoamingSettingCollection;
	var _typeComment="Comment";
	var Comment=(function (_super) {
		__extends(Comment, _super);
		function Comment() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(Comment.prototype, "_className", {
			get: function () {
				return "Comment";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Comment.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["id", "text", "created", "level", "resolved", "author", "mentions"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Comment.prototype, "_scalarPropertyUpdateable", {
			get: function () {
				return [false, true, false, false, true, false, false];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Comment.prototype, "_navigationPropertyNames", {
			get: function () {
				return ["parent", "parentOrNullObject", "replies"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Comment.prototype, "parent", {
			get: function () {
				if (!this._P) {
					this._P=_createPropertyObject(OfficeCore.Comment, this, "Parent", false, 4);
				}
				return this._P;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Comment.prototype, "parentOrNullObject", {
			get: function () {
				if (!this._Pa) {
					this._Pa=_createPropertyObject(OfficeCore.Comment, this, "ParentOrNullObject", false, 4);
				}
				return this._Pa;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Comment.prototype, "replies", {
			get: function () {
				if (!this._R) {
					this._R=_createPropertyObject(OfficeCore.CommentCollection, this, "Replies", true, 4);
				}
				return this._R;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Comment.prototype, "author", {
			get: function () {
				_throwIfNotLoaded("author", this._A, _typeComment, this._isNull);
				return this._A;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Comment.prototype, "created", {
			get: function () {
				_throwIfNotLoaded("created", this._C, _typeComment, this._isNull);
				return this._C;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Comment.prototype, "id", {
			get: function () {
				_throwIfNotLoaded("id", this._I, _typeComment, this._isNull);
				return this._I;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Comment.prototype, "level", {
			get: function () {
				_throwIfNotLoaded("level", this._L, _typeComment, this._isNull);
				return this._L;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Comment.prototype, "mentions", {
			get: function () {
				_throwIfNotLoaded("mentions", this._M, _typeComment, this._isNull);
				return this._M;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Comment.prototype, "resolved", {
			get: function () {
				_throwIfNotLoaded("resolved", this._Re, _typeComment, this._isNull);
				return this._Re;
			},
			set: function (value) {
				this._Re=value;
				_invokeSetProperty(this, "Resolved", value, 0);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Comment.prototype, "text", {
			get: function () {
				_throwIfNotLoaded("text", this._T, _typeComment, this._isNull);
				return this._T;
			},
			set: function (value) {
				this._T=value;
				_invokeSetProperty(this, "Text", value, 0);
			},
			enumerable: true,
			configurable: true
		});
		Comment.prototype.set=function (properties, options) {
			this._recursivelySet(properties, options, ["text", "resolved"], [], [
				"parent",
				"parentOrNullObject",
				"replies"
			]);
		};
		Comment.prototype.update=function (properties) {
			this._recursivelyUpdate(properties);
		};
		Comment.prototype["delete"]=function () {
			_invokeMethod(this, "Delete", 0, [], 0, 0);
		};
		Comment.prototype.getParentOrSelf=function () {
			return _createMethodObject(OfficeCore.Comment, this, "GetParentOrSelf", 1, [], false, false, null, 4);
		};
		Comment.prototype.getRichText=function (format) {
			return _invokeMethod(this, "GetRichText", 1, [format], 4, 0);
		};
		Comment.prototype.reply=function (text, format) {
			return _createMethodObject(OfficeCore.Comment, this, "Reply", 0, [text, format], false, false, null, 0);
		};
		Comment.prototype.setRichText=function (text, format) {
			return _invokeMethod(this, "SetRichText", 0, [text, format], 0, 0);
		};
		Comment.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["Author"])) {
				this._A=obj["Author"];
			}
			if (!_isUndefined(obj["Created"])) {
				this._C=_adjustToDateTime(obj["Created"]);
			}
			if (!_isUndefined(obj["Id"])) {
				this._I=obj["Id"];
			}
			if (!_isUndefined(obj["Level"])) {
				this._L=obj["Level"];
			}
			if (!_isUndefined(obj["Mentions"])) {
				this._M=obj["Mentions"];
			}
			if (!_isUndefined(obj["Resolved"])) {
				this._Re=obj["Resolved"];
			}
			if (!_isUndefined(obj["Text"])) {
				this._T=obj["Text"];
			}
			_handleNavigationPropertyResults(this, obj, ["parent", "Parent", "parentOrNullObject", "ParentOrNullObject", "replies", "Replies"]);
		};
		Comment.prototype.load=function (option) {
			return _load(this, option);
		};
		Comment.prototype.retrieve=function (option) {
			return _retrieve(this, option);
		};
		Comment.prototype._handleIdResult=function (value) {
			_super.prototype._handleIdResult.call(this, value);
			if (_isNullOrUndefined(value)) {
				return;
			}
			if (!_isUndefined(value["Id"])) {
				this._I=value["Id"];
			}
		};
		Comment.prototype._handleRetrieveResult=function (value, result) {
			_super.prototype._handleRetrieveResult.call(this, value, result);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			if (!_isUndefined(obj["Created"])) {
				obj["created"]=_adjustToDateTime(obj["created"]);
			}
			_processRetrieveResult(this, value, result);
		};
		Comment.prototype.toJSON=function () {
			return _toJson(this, {
				"author": this._A,
				"created": this._C,
				"id": this._I,
				"level": this._L,
				"mentions": this._M,
				"resolved": this._Re,
				"text": this._T
			}, {
				"replies": this._R
			});
		};
		Comment.prototype.ensureUnchanged=function (data) {
			_invokeEnsureUnchanged(this, data);
			return;
		};
		return Comment;
	}(OfficeExtension.ClientObject));
	OfficeCore.Comment=Comment;
	var _typeCommentCollection="CommentCollection";
	var CommentCollection=(function (_super) {
		__extends(CommentCollection, _super);
		function CommentCollection() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(CommentCollection.prototype, "_className", {
			get: function () {
				return "CommentCollection";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(CommentCollection.prototype, "_isCollection", {
			get: function () {
				return true;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(CommentCollection.prototype, "items", {
			get: function () {
				_throwIfNotLoaded("items", this.m__items, _typeCommentCollection, this._isNull);
				return this.m__items;
			},
			enumerable: true,
			configurable: true
		});
		CommentCollection.prototype.getCount=function () {
			return _invokeMethod(this, "GetCount", 1, [], 4, 0);
		};
		CommentCollection.prototype.getItem=function (id) {
			return _createIndexerObject(OfficeCore.Comment, this, [id]);
		};
		CommentCollection.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
				this.m__items=[];
				var _data=obj[OfficeExtension.Constants.items];
				for (var i=0; i < _data.length; i++) {
					var _item=_createChildItemObject(OfficeCore.Comment, true, this, _data[i], i);
					_item._handleResult(_data[i]);
					this.m__items.push(_item);
				}
			}
		};
		CommentCollection.prototype.load=function (option) {
			return _load(this, option);
		};
		CommentCollection.prototype.retrieve=function (option) {
			return _retrieve(this, option);
		};
		CommentCollection.prototype._handleRetrieveResult=function (value, result) {
			var _this=this;
			_super.prototype._handleRetrieveResult.call(this, value, result);
			_processRetrieveResult(this, value, result, function (childItemData, index) { return _createChildItemObject(OfficeCore.Comment, true, _this, childItemData, index); });
		};
		CommentCollection.prototype.toJSON=function () {
			return _toJson(this, {}, {}, this.m__items);
		};
		return CommentCollection;
	}(OfficeExtension.ClientObject));
	OfficeCore.CommentCollection=CommentCollection;
	var CommentTextFormat;
	(function (CommentTextFormat) {
		CommentTextFormat["plain"]="Plain";
		CommentTextFormat["markdown"]="Markdown";
		CommentTextFormat["delta"]="Delta";
	})(CommentTextFormat=OfficeCore.CommentTextFormat || (OfficeCore.CommentTextFormat={}));
	var _typeTap="Tap";
	var Tap=(function (_super) {
		__extends(Tap, _super);
		function Tap() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(Tap.prototype, "_className", {
			get: function () {
				return "Tap";
			},
			enumerable: true,
			configurable: true
		});
		Tap.prototype.getEnterpriseUserInfo=function () {
			return _invokeMethod(this, "GetEnterpriseUserInfo", 1, [], 4 | 1, 0);
		};
		Tap.prototype.getMruFriendlyPath=function (documentUrl) {
			return _invokeMethod(this, "GetMruFriendlyPath", 1, [documentUrl], 4 | 1, 0);
		};
		Tap.prototype.launchFileUrlInOfficeApp=function (documentUrl, useUniversalAsBackup) {
			return _invokeMethod(this, "LaunchFileUrlInOfficeApp", 1, [documentUrl, useUniversalAsBackup], 4 | 1, 0);
		};
		Tap.prototype.performLocalSearch=function (query, numResultsRequested, supportedFileExtensions, documentUrlToExclude) {
			return _invokeMethod(this, "PerformLocalSearch", 1, [query, numResultsRequested, supportedFileExtensions, documentUrlToExclude], 4 | 1, 0);
		};
		Tap.prototype.readSearchCache=function (keyword, expiredHours, filterObjectType) {
			return _invokeMethod(this, "ReadSearchCache", 1, [keyword, expiredHours, filterObjectType], 4 | 1, 0);
		};
		Tap.prototype.writeSearchCache=function (fileContent, keyword, filterObjectType) {
			return _invokeMethod(this, "WriteSearchCache", 1, [fileContent, keyword, filterObjectType], 4 | 1, 0);
		};
		Tap.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
		};
		Tap.prototype._handleRetrieveResult=function (value, result) {
			_super.prototype._handleRetrieveResult.call(this, value, result);
			_processRetrieveResult(this, value, result);
		};
		Tap.newObject=function (context) {
			return _createTopLevelServiceObject(OfficeCore.Tap, context, "Microsoft.TapRichApi.Tap", false, 4);
		};
		Tap.prototype.toJSON=function () {
			return _toJson(this, {}, {});
		};
		return Tap;
	}(OfficeExtension.ClientObject));
	OfficeCore.Tap=Tap;
	var ObjectType;
	(function (ObjectType) {
		ObjectType["unknown"]="Unknown";
		ObjectType["chart"]="Chart";
		ObjectType["smartArt"]="SmartArt";
		ObjectType["table"]="Table";
		ObjectType["image"]="Image";
		ObjectType["slide"]="Slide";
		ObjectType["ole"]="OLE";
		ObjectType["text"]="Text";
	})(ObjectType=OfficeCore.ObjectType || (OfficeCore.ObjectType={}));
	var ErrorCodes;
	(function (ErrorCodes) {
		ErrorCodes["apiNotAvailable"]="ApiNotAvailable";
		ErrorCodes["clientError"]="ClientError";
		ErrorCodes["generalException"]="GeneralException";
		ErrorCodes["interactiveFlowAborted"]="InteractiveFlowAborted";
		ErrorCodes["invalidArgument"]="InvalidArgument";
		ErrorCodes["invalidGrant"]="InvalidGrant";
		ErrorCodes["invalidResourceUrl"]="InvalidResourceUrl";
		ErrorCodes["resourceNotSupported"]="ResourceNotSupported";
		ErrorCodes["serverError"]="ServerError";
		ErrorCodes["unsupportedUserIdentity"]="UnsupportedUserIdentity";
		ErrorCodes["userNotSignedIn"]="UserNotSignedIn";
	})(ErrorCodes=OfficeCore.ErrorCodes || (OfficeCore.ErrorCodes={}));
	var Interfaces;
	(function (Interfaces) {
	})(Interfaces=OfficeCore.Interfaces || (OfficeCore.Interfaces={}));
})(OfficeCore || (OfficeCore={}));
var __extends=(this && this.__extends) || (function () {
	var extendStatics=function (d, b) {
		extendStatics=Object.setPrototypeOf ||
			({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__=b; }) ||
			function (d, b) { for (var p in b)
				if (b.hasOwnProperty(p))
					d[p]=b[p]; };
		return extendStatics(d, b);
	};
	return function (d, b) {
		extendStatics(d, b);
		function __() { this.constructor=d; }
		d.prototype=b===null ? Object.create(b) : (__.prototype=b.prototype, new __());
	};
})();
var ExcelOp;
(function (ExcelOp) {
	var _hostName="Excel";
	var _defaultApiSetName="ExcelApi";
	var _throwIfApiNotSupported=OfficeExtension.CommonUtility.throwIfApiNotSupported;
	var _invokeRetrieve=OfficeExtension.OperationalApiHelper.invokeRetrieve;
	var _invokeMethod=OfficeExtension.OperationalApiHelper.invokeMethod;
	var _invokeRecursiveUpdate=OfficeExtension.OperationalApiHelper.invokeRecursiveUpdate;
	var _createRootServiceObject=OfficeExtension.OperationalApiHelper.createRootServiceObject;
	var _createTopLevelServiceObject=OfficeExtension.OperationalApiHelper.createTopLevelServiceObject;
	var _createPropertyObject=OfficeExtension.OperationalApiHelper.createPropertyObject;
	var _createIndexerObject=OfficeExtension.OperationalApiHelper.createIndexerObject;
	var _createMethodObject=OfficeExtension.OperationalApiHelper.createMethodObject;
	var _createAndInstantiateMethodObject=OfficeExtension.OperationalApiHelper.createAndInstantiateMethodObject;
	var _localDocumentContext=OfficeExtension.OperationalApiHelper.localDocumentContext;
	var Runtime=(function (_super) {
		__extends(Runtime, _super);
		function Runtime() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(Runtime.prototype, "_className", {
			get: function () {
				return "Runtime";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Runtime.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["enableEvents"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Runtime.prototype, "_scalarPropertyUpdateable", {
			get: function () {
				return [true];
			},
			enumerable: true,
			configurable: true
		});
		Runtime.prototype.update=function (properties) {
			return _invokeRecursiveUpdate(this, properties);
		};
		Runtime.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		Runtime.prototype.toJSON=function () {
			return {};
		};
		return Runtime;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.Runtime=Runtime;
	var Application=(function (_super) {
		__extends(Application, _super);
		function Application() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(Application.prototype, "_className", {
			get: function () {
				return "Application";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Application.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["calculationMode", "calculationEngineVersion", "calculationState"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Application.prototype, "_scalarPropertyUpdateable", {
			get: function () {
				return [true, false, false];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Application.prototype, "_navigationPropertyNames", {
			get: function () {
				return ["iterativeCalculation", "ribbon"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Application.prototype, "iterativeCalculation", {
			get: function () {
				_throwIfApiNotSupported("Application.iterativeCalculation", _defaultApiSetName, "1.9", _hostName);
				return _createPropertyObject(ExcelOp.IterativeCalculation, this, "IterativeCalculation", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Application.prototype, "ribbon", {
			get: function () {
				_throwIfApiNotSupported("Application.ribbon", _defaultApiSetName, "1.9", _hostName);
				return _createPropertyObject(ExcelOp.Ribbon, this, "Ribbon", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Application.prototype.createWorkbook=function (base64File) {
			_throwIfApiNotSupported("Application.createWorkbook", _defaultApiSetName, "1.8", _hostName);
			return _createMethodObject(ExcelOp.WorkbookCreated, this, "CreateWorkbook", 1, [base64File], false, true, "_GetWorkbookCreatedById", 0);
		};
		Application.prototype._GetWorkbookCreatedById=function (id) {
			_throwIfApiNotSupported("Application._GetWorkbookCreatedById", _defaultApiSetName, "1.8", _hostName);
			return _createMethodObject(ExcelOp.WorkbookCreated, this, "_GetWorkbookCreatedById", 1, [id], false, false, null, 4);
		};
		Application.prototype.update=function (properties) {
			return _invokeRecursiveUpdate(this, properties);
		};
		Application.prototype.calculate=function (calculationType) {
			return _invokeMethod(this, "Calculate", 0, [calculationType], 0);
		};
		Application.prototype.suspendApiCalculationUntilNextSync=function () {
			_throwIfApiNotSupported("Application.suspendApiCalculationUntilNextSync", _defaultApiSetName, "1.6", _hostName);
			return _invokeMethod(this, "SuspendApiCalculationUntilNextSync", 0, [], 0);
		};
		Application.prototype.suspendScreenUpdatingUntilNextSync=function () {
			_throwIfApiNotSupported("Application.suspendScreenUpdatingUntilNextSync", _defaultApiSetName, "1.9", _hostName);
			return _invokeMethod(this, "SuspendScreenUpdatingUntilNextSync", 0, [], 0);
		};
		Application.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		Application.prototype.toJSON=function () {
			return {};
		};
		return Application;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.Application=Application;
	var IterativeCalculation=(function (_super) {
		__extends(IterativeCalculation, _super);
		function IterativeCalculation() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(IterativeCalculation.prototype, "_className", {
			get: function () {
				return "IterativeCalculation";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(IterativeCalculation.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["enabled", "maxIteration", "maxChange"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(IterativeCalculation.prototype, "_scalarPropertyUpdateable", {
			get: function () {
				return [true, true, true];
			},
			enumerable: true,
			configurable: true
		});
		IterativeCalculation.prototype.update=function (properties) {
			return _invokeRecursiveUpdate(this, properties);
		};
		IterativeCalculation.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		IterativeCalculation.prototype.toJSON=function () {
			return {};
		};
		return IterativeCalculation;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.IterativeCalculation=IterativeCalculation;
	var Workbook=(function (_super) {
		__extends(Workbook, _super);
		function Workbook() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(Workbook.prototype, "_className", {
			get: function () {
				return "Workbook";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Workbook.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["name", "readOnly", "isDirty", "use1904DateSystem", "chartDataPointTrack", "usePrecisionAsDisplayed", "calculationEngineVersion", "autoSave", "previouslySaved"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Workbook.prototype, "_scalarPropertyUpdateable", {
			get: function () {
				return [false, false, true, true, true, true, false, false, false];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Workbook.prototype, "_navigationPropertyNames", {
			get: function () {
				return ["worksheets", "names", "tables", "application", "bindings", "functions", "_V1Api", "pivotTables", "settings", "customXmlParts", "internalTest", "properties", "styles", "protection", "dataConnections", "_Runtime", "comments", "slicers"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Workbook.prototype, "application", {
			get: function () {
				return _createPropertyObject(ExcelOp.Application, this, "Application", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Workbook.prototype, "bindings", {
			get: function () {
				return _createPropertyObject(ExcelOp.BindingCollection, this, "Bindings", true, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Workbook.prototype, "comments", {
			get: function () {
				_throwIfApiNotSupported("Workbook.comments", _defaultApiSetName, "1.9", _hostName);
				return _createPropertyObject(ExcelOp.CommentCollection, this, "Comments", true, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Workbook.prototype, "customXmlParts", {
			get: function () {
				_throwIfApiNotSupported("Workbook.customXmlParts", _defaultApiSetName, "1.5", _hostName);
				return _createPropertyObject(ExcelOp.CustomXmlPartCollection, this, "CustomXmlParts", true, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Workbook.prototype, "dataConnections", {
			get: function () {
				_throwIfApiNotSupported("Workbook.dataConnections", _defaultApiSetName, "1.7", _hostName);
				return _createPropertyObject(ExcelOp.DataConnectionCollection, this, "DataConnections", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Workbook.prototype, "functions", {
			get: function () {
				_throwIfApiNotSupported("Workbook.functions", _defaultApiSetName, "1.2", _hostName);
				return _createPropertyObject(ExcelOp.Functions, this, "Functions", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Workbook.prototype, "internalTest", {
			get: function () {
				_throwIfApiNotSupported("Workbook.internalTest", _defaultApiSetName, "1.6", _hostName);
				return _createPropertyObject(ExcelOp.InternalTest, this, "InternalTest", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Workbook.prototype, "names", {
			get: function () {
				return _createPropertyObject(ExcelOp.NamedItemCollection, this, "Names", true, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Workbook.prototype, "pivotTables", {
			get: function () {
				_throwIfApiNotSupported("Workbook.pivotTables", _defaultApiSetName, "1.3", _hostName);
				return _createPropertyObject(ExcelOp.PivotTableCollection, this, "PivotTables", true, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Workbook.prototype, "properties", {
			get: function () {
				_throwIfApiNotSupported("Workbook.properties", _defaultApiSetName, "1.7", _hostName);
				return _createPropertyObject(ExcelOp.DocumentProperties, this, "Properties", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Workbook.prototype, "protection", {
			get: function () {
				_throwIfApiNotSupported("Workbook.protection", _defaultApiSetName, "1.7", _hostName);
				return _createPropertyObject(ExcelOp.WorkbookProtection, this, "Protection", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Workbook.prototype, "settings", {
			get: function () {
				_throwIfApiNotSupported("Workbook.settings", _defaultApiSetName, "1.4", _hostName);
				return _createPropertyObject(ExcelOp.SettingCollection, this, "Settings", true, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Workbook.prototype, "slicers", {
			get: function () {
				_throwIfApiNotSupported("Workbook.slicers", _defaultApiSetName, "1.9", _hostName);
				return _createPropertyObject(ExcelOp.SlicerCollection, this, "Slicers", true, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Workbook.prototype, "styles", {
			get: function () {
				_throwIfApiNotSupported("Workbook.styles", _defaultApiSetName, "1.7", _hostName);
				return _createPropertyObject(ExcelOp.StyleCollection, this, "Styles", true, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Workbook.prototype, "tables", {
			get: function () {
				return _createPropertyObject(ExcelOp.TableCollection, this, "Tables", true, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Workbook.prototype, "worksheets", {
			get: function () {
				return _createPropertyObject(ExcelOp.WorksheetCollection, this, "Worksheets", true, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Workbook.prototype, "_Runtime", {
			get: function () {
				_throwIfApiNotSupported("Workbook._Runtime", _defaultApiSetName, "1.5", _hostName);
				return _createPropertyObject(ExcelOp.Runtime, this, "_Runtime", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Workbook.prototype, "_V1Api", {
			get: function () {
				_throwIfApiNotSupported("Workbook._V1Api", _defaultApiSetName, "1.3", _hostName);
				return _createPropertyObject(ExcelOp._V1Api, this, "_V1Api", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Workbook.prototype.getActiveCell=function () {
			_throwIfApiNotSupported("Workbook.getActiveCell", _defaultApiSetName, "1.7", _hostName);
			return _createMethodObject(ExcelOp.Range, this, "GetActiveCell", 1, [], false, true, null, 4);
		};
		Workbook.prototype.getActiveChart=function () {
			_throwIfApiNotSupported("Workbook.getActiveChart", _defaultApiSetName, "1.9", _hostName);
			return _createMethodObject(ExcelOp.Chart, this, "GetActiveChart", 1, [], false, false, null, 4);
		};
		Workbook.prototype.getActiveChartOrNullObject=function () {
			_throwIfApiNotSupported("Workbook.getActiveChartOrNullObject", _defaultApiSetName, "1.9", _hostName);
			return _createMethodObject(ExcelOp.Chart, this, "GetActiveChartOrNullObject", 1, [], false, false, null, 4);
		};
		Workbook.prototype.getActiveSlicer=function () {
			_throwIfApiNotSupported("Workbook.getActiveSlicer", _defaultApiSetName, "1.9", _hostName);
			return _createMethodObject(ExcelOp.Slicer, this, "GetActiveSlicer", 1, [], false, false, null, 4);
		};
		Workbook.prototype.getActiveSlicerOrNullObject=function () {
			_throwIfApiNotSupported("Workbook.getActiveSlicerOrNullObject", _defaultApiSetName, "1.9", _hostName);
			return _createMethodObject(ExcelOp.Slicer, this, "GetActiveSlicerOrNullObject", 1, [], false, false, null, 4);
		};
		Workbook.prototype.getSelectedRange=function () {
			return _createMethodObject(ExcelOp.Range, this, "GetSelectedRange", 1, [], false, true, null, 4);
		};
		Workbook.prototype.getSelectedRanges=function () {
			_throwIfApiNotSupported("Workbook.getSelectedRanges", _defaultApiSetName, "1.9", _hostName);
			return _createMethodObject(ExcelOp.RangeAreas, this, "GetSelectedRanges", 1, [], false, true, null, 4);
		};
		Workbook.prototype._GetRangeForEventByReferenceId=function (bstrReferenceId) {
			return _createMethodObject(ExcelOp.Range, this, "_GetRangeForEventByReferenceId", 1, [bstrReferenceId], false, false, null, 4);
		};
		Workbook.prototype._GetRangeOrNullObjectForEventByReferenceId=function (bstrReferenceId) {
			return _createMethodObject(ExcelOp.Range, this, "_GetRangeOrNullObjectForEventByReferenceId", 1, [bstrReferenceId], false, false, null, 4);
		};
		Workbook.prototype._GetRangesForEventByReferenceId=function (bstrReferenceId) {
			_throwIfApiNotSupported("Workbook._GetRangesForEventByReferenceId", _defaultApiSetName, "1.9", _hostName);
			return _createMethodObject(ExcelOp.RangeAreas, this, "_GetRangesForEventByReferenceId", 1, [bstrReferenceId], false, false, null, 4);
		};
		Workbook.prototype._GetRangesOrNullObjectForEventByReferenceId=function (bstrReferenceId) {
			_throwIfApiNotSupported("Workbook._GetRangesOrNullObjectForEventByReferenceId", _defaultApiSetName, "1.9", _hostName);
			return _createMethodObject(ExcelOp.RangeAreas, this, "_GetRangesOrNullObjectForEventByReferenceId", 1, [bstrReferenceId], false, false, null, 4);
		};
		Workbook.prototype.update=function (properties) {
			return _invokeRecursiveUpdate(this, properties);
		};
		Workbook.prototype.close=function (closeBehavior) {
			_throwIfApiNotSupported("Workbook.close", _defaultApiSetName, "1.9", _hostName);
			return _invokeMethod(this, "Close", 0, [closeBehavior], 0);
		};
		Workbook.prototype.getIsActiveCollabSession=function () {
			_throwIfApiNotSupported("Workbook.getIsActiveCollabSession", _defaultApiSetName, "1.9", _hostName);
			return _invokeMethod(this, "GetIsActiveCollabSession", 0, [], 0, 0);
		};
		Workbook.prototype.registerCustomFunctions=function (addinNamespace, metadataContent, addinId, locale, addinInvariantNamespace, addinTitle, isXllCompatible) {
			_throwIfApiNotSupported("Workbook.registerCustomFunctions", "CustomFunctions", "1.1", _hostName);
			return _invokeMethod(this, "RegisterCustomFunctions", 0, [addinNamespace, metadataContent, addinId, locale, addinInvariantNamespace, addinTitle, isXllCompatible], 0);
		};
		Workbook.prototype.save=function (saveBehavior) {
			_throwIfApiNotSupported("Workbook.save", _defaultApiSetName, "1.9", _hostName);
			return _invokeMethod(this, "Save", 0, [saveBehavior], 0);
		};
		Workbook.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		Workbook.prototype.toJSON=function () {
			return {};
		};
		return Workbook;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.Workbook=Workbook;
	ExcelOp.workbook=_createRootServiceObject(Workbook, _localDocumentContext);
	var WorkbookProtection=(function (_super) {
		__extends(WorkbookProtection, _super);
		function WorkbookProtection() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(WorkbookProtection.prototype, "_className", {
			get: function () {
				return "WorkbookProtection";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(WorkbookProtection.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["protected"];
			},
			enumerable: true,
			configurable: true
		});
		WorkbookProtection.prototype.protect=function (password) {
			return _invokeMethod(this, "Protect", 0, [password], 0);
		};
		WorkbookProtection.prototype.unprotect=function (password) {
			return _invokeMethod(this, "Unprotect", 0, [password], 0);
		};
		WorkbookProtection.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		WorkbookProtection.prototype.toJSON=function () {
			return {};
		};
		return WorkbookProtection;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.WorkbookProtection=WorkbookProtection;
	var WorkbookCreated=(function (_super) {
		__extends(WorkbookCreated, _super);
		function WorkbookCreated() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(WorkbookCreated.prototype, "_className", {
			get: function () {
				return "WorkbookCreated";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(WorkbookCreated.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["id"];
			},
			enumerable: true,
			configurable: true
		});
		WorkbookCreated.prototype.open=function () {
			return _invokeMethod(this, "Open", 1, [], 4);
		};
		WorkbookCreated.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		WorkbookCreated.prototype.toJSON=function () {
			return {};
		};
		return WorkbookCreated;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.WorkbookCreated=WorkbookCreated;
	var Worksheet=(function (_super) {
		__extends(Worksheet, _super);
		function Worksheet() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(Worksheet.prototype, "_className", {
			get: function () {
				return "Worksheet";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Worksheet.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["name", "id", "position", "visibility", "tabColor", "standardWidth", "standardHeight", "showGridlines", "showHeadings", "enableCalculation"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Worksheet.prototype, "_scalarPropertyUpdateable", {
			get: function () {
				return [true, false, true, true, true, true, false, true, true, true];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Worksheet.prototype, "_navigationPropertyNames", {
			get: function () {
				return ["charts", "tables", "protection", "pivotTables", "names", "freezePanes", "pageLayout", "visuals", "shapes", "horizontalPageBreaks", "verticalPageBreaks", "autoFilter", "slicers", "comments"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Worksheet.prototype, "autoFilter", {
			get: function () {
				_throwIfApiNotSupported("Worksheet.autoFilter", _defaultApiSetName, "1.9", _hostName);
				return _createPropertyObject(ExcelOp.AutoFilter, this, "AutoFilter", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Worksheet.prototype, "charts", {
			get: function () {
				return _createPropertyObject(ExcelOp.ChartCollection, this, "Charts", true, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Worksheet.prototype, "comments", {
			get: function () {
				_throwIfApiNotSupported("Worksheet.comments", _defaultApiSetName, "1.9", _hostName);
				return _createPropertyObject(ExcelOp.CommentCollection, this, "Comments", true, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Worksheet.prototype, "freezePanes", {
			get: function () {
				_throwIfApiNotSupported("Worksheet.freezePanes", _defaultApiSetName, "1.7", _hostName);
				return _createPropertyObject(ExcelOp.WorksheetFreezePanes, this, "FreezePanes", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Worksheet.prototype, "horizontalPageBreaks", {
			get: function () {
				_throwIfApiNotSupported("Worksheet.horizontalPageBreaks", _defaultApiSetName, "1.9", _hostName);
				return _createPropertyObject(ExcelOp.PageBreakCollection, this, "HorizontalPageBreaks", true, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Worksheet.prototype, "names", {
			get: function () {
				_throwIfApiNotSupported("Worksheet.names", _defaultApiSetName, "1.4", _hostName);
				return _createPropertyObject(ExcelOp.NamedItemCollection, this, "Names", true, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Worksheet.prototype, "pageLayout", {
			get: function () {
				_throwIfApiNotSupported("Worksheet.pageLayout", _defaultApiSetName, "1.9", _hostName);
				return _createPropertyObject(ExcelOp.PageLayout, this, "PageLayout", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Worksheet.prototype, "pivotTables", {
			get: function () {
				_throwIfApiNotSupported("Worksheet.pivotTables", _defaultApiSetName, "1.3", _hostName);
				return _createPropertyObject(ExcelOp.PivotTableCollection, this, "PivotTables", true, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Worksheet.prototype, "protection", {
			get: function () {
				_throwIfApiNotSupported("Worksheet.protection", _defaultApiSetName, "1.2", _hostName);
				return _createPropertyObject(ExcelOp.WorksheetProtection, this, "Protection", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Worksheet.prototype, "shapes", {
			get: function () {
				_throwIfApiNotSupported("Worksheet.shapes", _defaultApiSetName, "1.9", _hostName);
				return _createPropertyObject(ExcelOp.ShapeCollection, this, "Shapes", true, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Worksheet.prototype, "slicers", {
			get: function () {
				_throwIfApiNotSupported("Worksheet.slicers", _defaultApiSetName, "1.9", _hostName);
				return _createPropertyObject(ExcelOp.SlicerCollection, this, "Slicers", true, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Worksheet.prototype, "tables", {
			get: function () {
				return _createPropertyObject(ExcelOp.TableCollection, this, "Tables", true, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Worksheet.prototype, "verticalPageBreaks", {
			get: function () {
				_throwIfApiNotSupported("Worksheet.verticalPageBreaks", _defaultApiSetName, "1.9", _hostName);
				return _createPropertyObject(ExcelOp.PageBreakCollection, this, "VerticalPageBreaks", true, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Worksheet.prototype, "visuals", {
			get: function () {
				_throwIfApiNotSupported("Worksheet.visuals", _defaultApiSetName, "99.9", _hostName);
				return _createPropertyObject(ExcelOp.VisualCollection, this, "Visuals", true, 4);
			},
			enumerable: true,
			configurable: true
		});
		Worksheet.prototype.copy=function (positionType, relativeTo) {
			_throwIfApiNotSupported("Worksheet.copy", _defaultApiSetName, "1.7", _hostName);
			return _createAndInstantiateMethodObject(ExcelOp.Worksheet, this, "Copy", 0, [positionType, relativeTo], false, false, "_GetAnotherWorksheetById", 0);
		};
		Worksheet.prototype.findAll=function (text, criteria) {
			_throwIfApiNotSupported("Worksheet.findAll", _defaultApiSetName, "1.9", _hostName);
			return _createMethodObject(ExcelOp.RangeAreas, this, "FindAll", 1, [text, criteria], false, true, null, 4);
		};
		Worksheet.prototype.findAllOrNullObject=function (text, criteria) {
			_throwIfApiNotSupported("Worksheet.findAllOrNullObject", _defaultApiSetName, "1.9", _hostName);
			return _createMethodObject(ExcelOp.RangeAreas, this, "FindAllOrNullObject", 1, [text, criteria], false, true, null, 4);
		};
		Worksheet.prototype.getCell=function (row, column) {
			return _createMethodObject(ExcelOp.Range, this, "GetCell", 1, [row, column], false, true, null, 4);
		};
		Worksheet.prototype.getNext=function (visibleOnly) {
			_throwIfApiNotSupported("Worksheet.getNext", _defaultApiSetName, "1.5", _hostName);
			return _createMethodObject(ExcelOp.Worksheet, this, "GetNext", 1, [visibleOnly], false, true, "_GetSheetById", 4);
		};
		Worksheet.prototype.getNextOrNullObject=function (visibleOnly) {
			_throwIfApiNotSupported("Worksheet.getNextOrNullObject", _defaultApiSetName, "1.5", _hostName);
			return _createMethodObject(ExcelOp.Worksheet, this, "GetNextOrNullObject", 1, [visibleOnly], false, true, "_GetSheetById", 4);
		};
		Worksheet.prototype.getPrevious=function (visibleOnly) {
			_throwIfApiNotSupported("Worksheet.getPrevious", _defaultApiSetName, "1.5", _hostName);
			return _createMethodObject(ExcelOp.Worksheet, this, "GetPrevious", 1, [visibleOnly], false, true, "_GetSheetById", 4);
		};
		Worksheet.prototype.getPreviousOrNullObject=function (visibleOnly) {
			_throwIfApiNotSupported("Worksheet.getPreviousOrNullObject", _defaultApiSetName, "1.5", _hostName);
			return _createMethodObject(ExcelOp.Worksheet, this, "GetPreviousOrNullObject", 1, [visibleOnly], false, true, "_GetSheetById", 4);
		};
		Worksheet.prototype.getRange=function (address) {
			return _createMethodObject(ExcelOp.Range, this, "GetRange", 1, [address], false, true, null, 4);
		};
		Worksheet.prototype.getRangeByIndexes=function (startRow, startColumn, rowCount, columnCount) {
			_throwIfApiNotSupported("Worksheet.getRangeByIndexes", _defaultApiSetName, "1.7", _hostName);
			return _createMethodObject(ExcelOp.Range, this, "GetRangeByIndexes", 1, [startRow, startColumn, rowCount, columnCount], false, true, null, 4);
		};
		Worksheet.prototype.getRanges=function (address) {
			_throwIfApiNotSupported("Worksheet.getRanges", _defaultApiSetName, "1.9", _hostName);
			return _createMethodObject(ExcelOp.RangeAreas, this, "GetRanges", 1, [address], false, true, null, 4);
		};
		Worksheet.prototype.getUsedRange=function (valuesOnly) {
			return _createMethodObject(ExcelOp.Range, this, "GetUsedRange", 1, [valuesOnly], false, true, null, 4);
		};
		Worksheet.prototype.getUsedRangeOrNullObject=function (valuesOnly) {
			_throwIfApiNotSupported("Worksheet.getUsedRangeOrNullObject", _defaultApiSetName, "1.4", _hostName);
			return _createMethodObject(ExcelOp.Range, this, "GetUsedRangeOrNullObject", 1, [valuesOnly], false, true, null, 4);
		};
		Worksheet.prototype._GetAnotherWorksheetById=function (id) {
			_throwIfApiNotSupported("Worksheet._GetAnotherWorksheetById", _defaultApiSetName, "1.7", _hostName);
			return _createAndInstantiateMethodObject(ExcelOp.Worksheet, this, "_GetAnotherWorksheetById", 0, [id], false, false, null, 0);
		};
		Worksheet.prototype._GetSheetById=function (id) {
			_throwIfApiNotSupported("Worksheet._GetSheetById", _defaultApiSetName, "1.7", _hostName);
			return _createMethodObject(ExcelOp.Worksheet, this, "_GetSheetById", 1, [id], false, false, null, 4);
		};
		Worksheet.prototype.update=function (properties) {
			return _invokeRecursiveUpdate(this, properties);
		};
		Worksheet.prototype.activate=function () {
			return _invokeMethod(this, "Activate", 1, [], 0);
		};
		Worksheet.prototype.calculate=function (markAllDirty) {
			_throwIfApiNotSupported("Worksheet.calculate", _defaultApiSetName, "1.6", _hostName);
			return _invokeMethod(this, "Calculate", 0, [markAllDirty], 0);
		};
		Worksheet.prototype["delete"]=function () {
			return _invokeMethod(this, "Delete", 0, [], 0);
		};
		Worksheet.prototype.replaceAll=function (text, replacement, criteria) {
			_throwIfApiNotSupported("Worksheet.replaceAll", _defaultApiSetName, "1.9", _hostName);
			return _invokeMethod(this, "ReplaceAll", 0, [text, replacement, criteria], 0, 0);
		};
		Worksheet.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		Worksheet.prototype.toJSON=function () {
			return {};
		};
		return Worksheet;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.Worksheet=Worksheet;
	var WorksheetCollection=(function (_super) {
		__extends(WorksheetCollection, _super);
		function WorksheetCollection() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(WorksheetCollection.prototype, "_className", {
			get: function () {
				return "WorksheetCollection";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(WorksheetCollection.prototype, "_isCollection", {
			get: function () {
				return true;
			},
			enumerable: true,
			configurable: true
		});
		WorksheetCollection.prototype.add=function (name) {
			return _createAndInstantiateMethodObject(ExcelOp.Worksheet, this, "Add", 0, [name], false, true, null, 0);
		};
		WorksheetCollection.prototype.getActiveWorksheet=function () {
			return _createMethodObject(ExcelOp.Worksheet, this, "GetActiveWorksheet", 1, [], false, false, null, 4);
		};
		WorksheetCollection.prototype.getFirst=function (visibleOnly) {
			_throwIfApiNotSupported("WorksheetCollection.getFirst", _defaultApiSetName, "1.5", _hostName);
			return _createMethodObject(ExcelOp.Worksheet, this, "GetFirst", 1, [visibleOnly], false, true, null, 4);
		};
		WorksheetCollection.prototype.getItem=function (key) {
			return _createIndexerObject(ExcelOp.Worksheet, this, [key]);
		};
		WorksheetCollection.prototype.getItemOrNullObject=function (key) {
			_throwIfApiNotSupported("WorksheetCollection.getItemOrNullObject", _defaultApiSetName, "1.4", _hostName);
			return _createMethodObject(ExcelOp.Worksheet, this, "GetItemOrNullObject", 1, [key], false, false, null, 4);
		};
		WorksheetCollection.prototype.getLast=function (visibleOnly) {
			_throwIfApiNotSupported("WorksheetCollection.getLast", _defaultApiSetName, "1.5", _hostName);
			return _createMethodObject(ExcelOp.Worksheet, this, "GetLast", 1, [visibleOnly], false, true, null, 4);
		};
		WorksheetCollection.prototype.addFromBase64=function (base64File, sheetNamesToInsert, positionType, relativeTo) {
			_throwIfApiNotSupported("WorksheetCollection.addFromBase64", _defaultApiSetName, "1.9", _hostName);
			return _invokeMethod(this, "AddFromBase64", 0, [base64File, sheetNamesToInsert, positionType, relativeTo], 0, 0);
		};
		WorksheetCollection.prototype.getCount=function (visibleOnly) {
			_throwIfApiNotSupported("WorksheetCollection.getCount", _defaultApiSetName, "1.4", _hostName);
			return _invokeMethod(this, "GetCount", 1, [visibleOnly], 4, 0);
		};
		WorksheetCollection.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		WorksheetCollection.prototype.toJSON=function () {
			return {};
		};
		return WorksheetCollection;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.WorksheetCollection=WorksheetCollection;
	var WorksheetProtection=(function (_super) {
		__extends(WorksheetProtection, _super);
		function WorksheetProtection() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(WorksheetProtection.prototype, "_className", {
			get: function () {
				return "WorksheetProtection";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(WorksheetProtection.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["protected", "options"];
			},
			enumerable: true,
			configurable: true
		});
		WorksheetProtection.prototype.protect=function (options, password) {
			return _invokeMethod(this, "Protect", 0, [options, password], 0);
		};
		WorksheetProtection.prototype.unprotect=function (password) {
			return _invokeMethod(this, "Unprotect", 0, [password], 0);
		};
		WorksheetProtection.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		WorksheetProtection.prototype.toJSON=function () {
			return {};
		};
		return WorksheetProtection;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.WorksheetProtection=WorksheetProtection;
	var WorksheetFreezePanes=(function (_super) {
		__extends(WorksheetFreezePanes, _super);
		function WorksheetFreezePanes() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(WorksheetFreezePanes.prototype, "_className", {
			get: function () {
				return "WorksheetFreezePanes";
			},
			enumerable: true,
			configurable: true
		});
		WorksheetFreezePanes.prototype.getLocation=function () {
			return _createMethodObject(ExcelOp.Range, this, "GetLocation", 1, [], false, true, null, 4);
		};
		WorksheetFreezePanes.prototype.getLocationOrNullObject=function () {
			return _createMethodObject(ExcelOp.Range, this, "GetLocationOrNullObject", 1, [], false, true, null, 4);
		};
		WorksheetFreezePanes.prototype.freezeAt=function (frozenRange) {
			return _invokeMethod(this, "FreezeAt", 0, [frozenRange], 0);
		};
		WorksheetFreezePanes.prototype.freezeColumns=function (count) {
			return _invokeMethod(this, "FreezeColumns", 0, [count], 0);
		};
		WorksheetFreezePanes.prototype.freezeRows=function (count) {
			return _invokeMethod(this, "FreezeRows", 0, [count], 0);
		};
		WorksheetFreezePanes.prototype.unfreeze=function () {
			return _invokeMethod(this, "Unfreeze", 0, [], 0);
		};
		WorksheetFreezePanes.prototype.toJSON=function () {
			return {};
		};
		return WorksheetFreezePanes;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.WorksheetFreezePanes=WorksheetFreezePanes;
	var Range=(function (_super) {
		__extends(Range, _super);
		function Range() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(Range.prototype, "_className", {
			get: function () {
				return "Range";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Range.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["numberFormat", "numberFormatLocal", "values", "text", "formulas", "formulasLocal", "rowIndex", "columnIndex", "rowCount", "columnCount", "address", "addressLocal", "cellCount", "_ReferenceId", "valueTypes", "formulasR1C1", "hidden", "rowHidden", "columnHidden", "isEntireColumn", "isEntireRow", "hyperlink", "style", "linkedDataTypeState", "hasSpill"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Range.prototype, "_scalarPropertyUpdateable", {
			get: function () {
				return [true, true, true, false, true, true, false, false, false, false, false, false, false, false, false, true, false, true, true, false, false, true, true, false, false];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Range.prototype, "_navigationPropertyNames", {
			get: function () {
				return ["format", "worksheet", "sort", "conditionalFormats", "dataValidation"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Range.prototype, "conditionalFormats", {
			get: function () {
				_throwIfApiNotSupported("Range.conditionalFormats", _defaultApiSetName, "1.6", _hostName);
				return _createPropertyObject(ExcelOp.ConditionalFormatCollection, this, "ConditionalFormats", true, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Range.prototype, "dataValidation", {
			get: function () {
				_throwIfApiNotSupported("Range.dataValidation", _defaultApiSetName, "1.8", _hostName);
				return _createPropertyObject(ExcelOp.DataValidation, this, "DataValidation", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Range.prototype, "format", {
			get: function () {
				return _createPropertyObject(ExcelOp.RangeFormat, this, "Format", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Range.prototype, "sort", {
			get: function () {
				_throwIfApiNotSupported("Range.sort", _defaultApiSetName, "1.2", _hostName);
				return _createPropertyObject(ExcelOp.RangeSort, this, "Sort", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Range.prototype, "worksheet", {
			get: function () {
				return _createPropertyObject(ExcelOp.Worksheet, this, "Worksheet", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Range.prototype.find=function (text, criteria) {
			_throwIfApiNotSupported("Range.find", _defaultApiSetName, "1.9", _hostName);
			return _createMethodObject(ExcelOp.Range, this, "Find", 1, [text, criteria], false, true, null, 4);
		};
		Range.prototype.findOrNullObject=function (text, criteria) {
			_throwIfApiNotSupported("Range.findOrNullObject", _defaultApiSetName, "1.9", _hostName);
			return _createMethodObject(ExcelOp.Range, this, "FindOrNullObject", 1, [text, criteria], false, true, null, 4);
		};
		Range.prototype.getAbsoluteResizedRange=function (numRows, numColumns) {
			_throwIfApiNotSupported("Range.getAbsoluteResizedRange", _defaultApiSetName, "1.7", _hostName);
			return _createMethodObject(ExcelOp.Range, this, "GetAbsoluteResizedRange", 1, [numRows, numColumns], false, true, null, 4);
		};
		Range.prototype.getBoundingRect=function (anotherRange) {
			return _createMethodObject(ExcelOp.Range, this, "GetBoundingRect", 1, [anotherRange], false, true, null, 4);
		};
		Range.prototype.getCell=function (row, column) {
			return _createMethodObject(ExcelOp.Range, this, "GetCell", 1, [row, column], false, true, null, 4);
		};
		Range.prototype.getColumn=function (column) {
			return _createMethodObject(ExcelOp.Range, this, "GetColumn", 1, [column], false, true, null, 4);
		};
		Range.prototype.getColumnsAfter=function (count) {
			_throwIfApiNotSupported("Range.getColumnsAfter", _defaultApiSetName, "1.3", _hostName);
			return _createMethodObject(ExcelOp.Range, this, "GetColumnsAfter", 1, [count], false, true, null, 4);
		};
		Range.prototype.getColumnsBefore=function (count) {
			_throwIfApiNotSupported("Range.getColumnsBefore", _defaultApiSetName, "1.3", _hostName);
			return _createMethodObject(ExcelOp.Range, this, "GetColumnsBefore", 1, [count], false, true, null, 4);
		};
		Range.prototype.getEntireColumn=function () {
			return _createMethodObject(ExcelOp.Range, this, "GetEntireColumn", 1, [], false, true, null, 4);
		};
		Range.prototype.getEntireRow=function () {
			return _createMethodObject(ExcelOp.Range, this, "GetEntireRow", 1, [], false, true, null, 4);
		};
		Range.prototype.getIntersection=function (anotherRange) {
			return _createMethodObject(ExcelOp.Range, this, "GetIntersection", 1, [anotherRange], false, true, null, 4);
		};
		Range.prototype.getIntersectionOrNullObject=function (anotherRange) {
			_throwIfApiNotSupported("Range.getIntersectionOrNullObject", _defaultApiSetName, "1.4", _hostName);
			return _createMethodObject(ExcelOp.Range, this, "GetIntersectionOrNullObject", 1, [anotherRange], false, true, null, 4);
		};
		Range.prototype.getLastCell=function () {
			return _createMethodObject(ExcelOp.Range, this, "GetLastCell", 1, [], false, true, null, 4);
		};
		Range.prototype.getLastColumn=function () {
			return _createMethodObject(ExcelOp.Range, this, "GetLastColumn", 1, [], false, true, null, 4);
		};
		Range.prototype.getLastRow=function () {
			return _createMethodObject(ExcelOp.Range, this, "GetLastRow", 1, [], false, true, null, 4);
		};
		Range.prototype.getOffsetRange=function (rowOffset, columnOffset) {
			return _createMethodObject(ExcelOp.Range, this, "GetOffsetRange", 1, [rowOffset, columnOffset], false, true, null, 4);
		};
		Range.prototype.getResizedRange=function (deltaRows, deltaColumns) {
			_throwIfApiNotSupported("Range.getResizedRange", _defaultApiSetName, "1.3", _hostName);
			return _createMethodObject(ExcelOp.Range, this, "GetResizedRange", 1, [deltaRows, deltaColumns], false, true, null, 4);
		};
		Range.prototype.getRow=function (row) {
			return _createMethodObject(ExcelOp.Range, this, "GetRow", 1, [row], false, true, null, 4);
		};
		Range.prototype.getRowsAbove=function (count) {
			_throwIfApiNotSupported("Range.getRowsAbove", _defaultApiSetName, "1.3", _hostName);
			return _createMethodObject(ExcelOp.Range, this, "GetRowsAbove", 1, [count], false, true, null, 4);
		};
		Range.prototype.getRowsBelow=function (count) {
			_throwIfApiNotSupported("Range.getRowsBelow", _defaultApiSetName, "1.3", _hostName);
			return _createMethodObject(ExcelOp.Range, this, "GetRowsBelow", 1, [count], false, true, null, 4);
		};
		Range.prototype.getSpecialCells=function (cellType, cellValueType) {
			_throwIfApiNotSupported("Range.getSpecialCells", _defaultApiSetName, "1.9", _hostName);
			return _createMethodObject(ExcelOp.RangeAreas, this, "GetSpecialCells", 1, [cellType, cellValueType], false, true, null, 4);
		};
		Range.prototype.getSpecialCellsOrNullObject=function (cellType, cellValueType) {
			_throwIfApiNotSupported("Range.getSpecialCellsOrNullObject", _defaultApiSetName, "1.9", _hostName);
			return _createMethodObject(ExcelOp.RangeAreas, this, "GetSpecialCellsOrNullObject", 1, [cellType, cellValueType], false, true, null, 4);
		};
		Range.prototype.getSpillParent=function () {
			_throwIfApiNotSupported("Range.getSpillParent", _defaultApiSetName, "2", _hostName);
			return _createMethodObject(ExcelOp.Range, this, "GetSpillParent", 1, [], false, true, null, 4);
		};
		Range.prototype.getSpillingToRange=function () {
			_throwIfApiNotSupported("Range.getSpillingToRange", _defaultApiSetName, "2", _hostName);
			return _createMethodObject(ExcelOp.Range, this, "GetSpillingToRange", 1, [], false, true, null, 4);
		};
		Range.prototype.getSurroundingRegion=function () {
			_throwIfApiNotSupported("Range.getSurroundingRegion", _defaultApiSetName, "1.7", _hostName);
			return _createMethodObject(ExcelOp.Range, this, "GetSurroundingRegion", 1, [], false, true, null, 4);
		};
		Range.prototype.getTables=function (fullyContained) {
			_throwIfApiNotSupported("Range.getTables", _defaultApiSetName, "1.9", _hostName);
			return _createMethodObject(ExcelOp.TableScopedCollection, this, "GetTables", 1, [fullyContained], true, false, null, 4);
		};
		Range.prototype.getUsedRange=function (valuesOnly) {
			return _createMethodObject(ExcelOp.Range, this, "GetUsedRange", 1, [valuesOnly], false, true, null, 4);
		};
		Range.prototype.getUsedRangeOrNullObject=function (valuesOnly) {
			_throwIfApiNotSupported("Range.getUsedRangeOrNullObject", _defaultApiSetName, "1.4", _hostName);
			return _createMethodObject(ExcelOp.Range, this, "GetUsedRangeOrNullObject", 1, [valuesOnly], false, true, null, 4);
		};
		Range.prototype.getVisibleView=function () {
			_throwIfApiNotSupported("Range.getVisibleView", _defaultApiSetName, "1.3", _hostName);
			return _createMethodObject(ExcelOp.RangeView, this, "GetVisibleView", 1, [], false, false, null, 4);
		};
		Range.prototype.insert=function (shift) {
			return _createAndInstantiateMethodObject(ExcelOp.Range, this, "Insert", 0, [shift], false, true, null, 0);
		};
		Range.prototype.removeDuplicates=function (columns, includesHeader) {
			_throwIfApiNotSupported("Range.removeDuplicates", _defaultApiSetName, "1.9", _hostName);
			return _createAndInstantiateMethodObject(ExcelOp.RemoveDuplicatesResult, this, "RemoveDuplicates", 0, [columns, includesHeader], false, true, null, 0);
		};
		Range.prototype.update=function (properties) {
			return _invokeRecursiveUpdate(this, properties);
		};
		Range.prototype.autoFill=function (destinationRange, autoFillType) {
			_throwIfApiNotSupported("Range.autoFill", _defaultApiSetName, "1.9", _hostName);
			return _invokeMethod(this, "AutoFill", 0, [destinationRange, autoFillType], 0);
		};
		Range.prototype.calculate=function () {
			_throwIfApiNotSupported("Range.calculate", _defaultApiSetName, "1.6", _hostName);
			return _invokeMethod(this, "Calculate", 0, [], 0);
		};
		Range.prototype.clear=function (applyTo) {
			return _invokeMethod(this, "Clear", 0, [applyTo], 0);
		};
		Range.prototype.convertDataTypeToText=function () {
			_throwIfApiNotSupported("Range.convertDataTypeToText", _defaultApiSetName, "1.9", _hostName);
			return _invokeMethod(this, "ConvertDataTypeToText", 0, [], 0);
		};
		Range.prototype.convertToLinkedDataType=function (serviceID, languageCulture) {
			_throwIfApiNotSupported("Range.convertToLinkedDataType", _defaultApiSetName, "1.9", _hostName);
			return _invokeMethod(this, "ConvertToLinkedDataType", 0, [serviceID, languageCulture], 0);
		};
		Range.prototype.copyFrom=function (sourceRange, copyType, skipBlanks, transpose) {
			_throwIfApiNotSupported("Range.copyFrom", _defaultApiSetName, "1.9", _hostName);
			return _invokeMethod(this, "CopyFrom", 0, [sourceRange, copyType, skipBlanks, transpose], 0);
		};
		Range.prototype["delete"]=function (shift) {
			return _invokeMethod(this, "Delete", 0, [shift], 0);
		};
		Range.prototype.flashFill=function () {
			_throwIfApiNotSupported("Range.flashFill", _defaultApiSetName, "1.9", _hostName);
			return _invokeMethod(this, "FlashFill", 0, [], 0);
		};
		Range.prototype.getCellProperties=function (cellPropertiesLoadOptions) {
			_throwIfApiNotSupported("Range.getCellProperties", _defaultApiSetName, "1.9", _hostName);
			return _invokeMethod(this, "GetCellProperties", 0, [cellPropertiesLoadOptions], 0, 0);
		};
		Range.prototype.getColumnProperties=function (columnPropertiesLoadOptions) {
			_throwIfApiNotSupported("Range.getColumnProperties", _defaultApiSetName, "1.9", _hostName);
			return _invokeMethod(this, "GetColumnProperties", 0, [columnPropertiesLoadOptions], 0, 0);
		};
		Range.prototype.getImage=function () {
			_throwIfApiNotSupported("Range.getImage", _defaultApiSetName, "1.7", _hostName);
			return _invokeMethod(this, "GetImage", 1, [], 4, 0);
		};
		Range.prototype.getRowProperties=function (rowPropertiesLoadOptions) {
			_throwIfApiNotSupported("Range.getRowProperties", _defaultApiSetName, "1.9", _hostName);
			return _invokeMethod(this, "GetRowProperties", 0, [rowPropertiesLoadOptions], 0, 0);
		};
		Range.prototype.merge=function (across) {
			_throwIfApiNotSupported("Range.merge", _defaultApiSetName, "1.2", _hostName);
			return _invokeMethod(this, "Merge", 0, [across], 0);
		};
		Range.prototype.replaceAll=function (text, replacement, criteria) {
			_throwIfApiNotSupported("Range.replaceAll", _defaultApiSetName, "1.9", _hostName);
			return _invokeMethod(this, "ReplaceAll", 0, [text, replacement, criteria], 0, 0);
		};
		Range.prototype.select=function () {
			return _invokeMethod(this, "Select", 1, [], 0);
		};
		Range.prototype.setCellProperties=function (cellPropertiesData) {
			_throwIfApiNotSupported("Range.setCellProperties", _defaultApiSetName, "1.9", _hostName);
			return _invokeMethod(this, "SetCellProperties", 0, [cellPropertiesData], 0);
		};
		Range.prototype.setColumnProperties=function (columnPropertiesData) {
			_throwIfApiNotSupported("Range.setColumnProperties", _defaultApiSetName, "1.9", _hostName);
			return _invokeMethod(this, "SetColumnProperties", 0, [columnPropertiesData], 0);
		};
		Range.prototype.setDirty=function () {
			_throwIfApiNotSupported("Range.setDirty", _defaultApiSetName, "1.9", _hostName);
			return _invokeMethod(this, "SetDirty", 0, [], 0);
		};
		Range.prototype.setRowProperties=function (rowPropertiesData) {
			_throwIfApiNotSupported("Range.setRowProperties", _defaultApiSetName, "1.9", _hostName);
			return _invokeMethod(this, "SetRowProperties", 0, [rowPropertiesData], 0);
		};
		Range.prototype.showCard=function () {
			_throwIfApiNotSupported("Range.showCard", _defaultApiSetName, "1.7", _hostName);
			return _invokeMethod(this, "ShowCard", 0, [], 0);
		};
		Range.prototype.showTeachingCallout=function (title, message) {
			_throwIfApiNotSupported("Range.showTeachingCallout", _defaultApiSetName, "1.9", _hostName);
			return _invokeMethod(this, "ShowTeachingCallout", 0, [title, message], 0);
		};
		Range.prototype.unmerge=function () {
			_throwIfApiNotSupported("Range.unmerge", _defaultApiSetName, "1.2", _hostName);
			return _invokeMethod(this, "Unmerge", 0, [], 0);
		};
		Range.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		Range.prototype.toJSON=function () {
			return {};
		};
		return Range;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.Range=Range;
	var RangeAreas=(function (_super) {
		__extends(RangeAreas, _super);
		function RangeAreas() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(RangeAreas.prototype, "_className", {
			get: function () {
				return "RangeAreas";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(RangeAreas.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["_ReferenceId", "address", "addressLocal", "areaCount", "cellCount", "isEntireColumn", "isEntireRow", "style"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(RangeAreas.prototype, "_scalarPropertyUpdateable", {
			get: function () {
				return [false, false, false, false, false, false, false, true];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(RangeAreas.prototype, "_navigationPropertyNames", {
			get: function () {
				return ["areas", "conditionalFormats", "format", "dataValidation", "worksheet"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(RangeAreas.prototype, "areas", {
			get: function () {
				return _createPropertyObject(ExcelOp.RangeCollection, this, "Areas", true, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(RangeAreas.prototype, "conditionalFormats", {
			get: function () {
				return _createPropertyObject(ExcelOp.ConditionalFormatCollection, this, "ConditionalFormats", true, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(RangeAreas.prototype, "dataValidation", {
			get: function () {
				return _createPropertyObject(ExcelOp.DataValidation, this, "DataValidation", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(RangeAreas.prototype, "format", {
			get: function () {
				return _createPropertyObject(ExcelOp.RangeFormat, this, "Format", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(RangeAreas.prototype, "worksheet", {
			get: function () {
				return _createPropertyObject(ExcelOp.Worksheet, this, "Worksheet", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		RangeAreas.prototype.getEntireColumn=function () {
			return _createMethodObject(ExcelOp.RangeAreas, this, "GetEntireColumn", 1, [], false, true, null, 4);
		};
		RangeAreas.prototype.getEntireRow=function () {
			return _createMethodObject(ExcelOp.RangeAreas, this, "GetEntireRow", 1, [], false, true, null, 4);
		};
		RangeAreas.prototype.getIntersection=function (anotherRange) {
			return _createMethodObject(ExcelOp.RangeAreas, this, "GetIntersection", 1, [anotherRange], false, true, null, 4);
		};
		RangeAreas.prototype.getIntersectionOrNullObject=function (anotherRange) {
			return _createMethodObject(ExcelOp.RangeAreas, this, "GetIntersectionOrNullObject", 1, [anotherRange], false, true, null, 4);
		};
		RangeAreas.prototype.getOffsetRangeAreas=function (rowOffset, columnOffset) {
			return _createMethodObject(ExcelOp.RangeAreas, this, "GetOffsetRangeAreas", 1, [rowOffset, columnOffset], false, true, null, 4);
		};
		RangeAreas.prototype.getSpecialCells=function (cellType, cellValueType) {
			return _createMethodObject(ExcelOp.RangeAreas, this, "GetSpecialCells", 1, [cellType, cellValueType], false, true, null, 4);
		};
		RangeAreas.prototype.getSpecialCellsOrNullObject=function (cellType, cellValueType) {
			return _createMethodObject(ExcelOp.RangeAreas, this, "GetSpecialCellsOrNullObject", 1, [cellType, cellValueType], false, true, null, 4);
		};
		RangeAreas.prototype.getTables=function (fullyContained) {
			return _createMethodObject(ExcelOp.TableScopedCollection, this, "GetTables", 1, [fullyContained], true, false, null, 4);
		};
		RangeAreas.prototype.getUsedRangeAreas=function (valuesOnly) {
			return _createMethodObject(ExcelOp.RangeAreas, this, "GetUsedRangeAreas", 1, [valuesOnly], false, true, null, 4);
		};
		RangeAreas.prototype.getUsedRangeAreasOrNullObject=function (valuesOnly) {
			return _createMethodObject(ExcelOp.RangeAreas, this, "GetUsedRangeAreasOrNullObject", 1, [valuesOnly], false, true, null, 4);
		};
		RangeAreas.prototype.update=function (properties) {
			return _invokeRecursiveUpdate(this, properties);
		};
		RangeAreas.prototype.calculate=function () {
			return _invokeMethod(this, "Calculate", 0, [], 0);
		};
		RangeAreas.prototype.clear=function (applyTo) {
			return _invokeMethod(this, "Clear", 0, [applyTo], 0);
		};
		RangeAreas.prototype.convertDataTypeToText=function () {
			return _invokeMethod(this, "ConvertDataTypeToText", 0, [], 0);
		};
		RangeAreas.prototype.convertToLinkedDataType=function (serviceID, languageCulture) {
			return _invokeMethod(this, "ConvertToLinkedDataType", 0, [serviceID, languageCulture], 0);
		};
		RangeAreas.prototype.copyFrom=function (sourceRange, copyType, skipBlanks, transpose) {
			return _invokeMethod(this, "CopyFrom", 0, [sourceRange, copyType, skipBlanks, transpose], 0);
		};
		RangeAreas.prototype.setDirty=function () {
			return _invokeMethod(this, "SetDirty", 0, [], 0);
		};
		RangeAreas.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		RangeAreas.prototype.toJSON=function () {
			return {};
		};
		return RangeAreas;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.RangeAreas=RangeAreas;
	var RangeView=(function (_super) {
		__extends(RangeView, _super);
		function RangeView() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(RangeView.prototype, "_className", {
			get: function () {
				return "RangeView";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(RangeView.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["numberFormat", "values", "text", "formulas", "formulasLocal", "formulasR1C1", "valueTypes", "rowCount", "columnCount", "cellAddresses", "index"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(RangeView.prototype, "_scalarPropertyUpdateable", {
			get: function () {
				return [true, true, false, true, true, true, false, false, false, false, false];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(RangeView.prototype, "_navigationPropertyNames", {
			get: function () {
				return ["rows"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(RangeView.prototype, "rows", {
			get: function () {
				return _createPropertyObject(ExcelOp.RangeViewCollection, this, "Rows", true, 4);
			},
			enumerable: true,
			configurable: true
		});
		RangeView.prototype.getRange=function () {
			return _createMethodObject(ExcelOp.Range, this, "GetRange", 1, [], false, true, null, 4);
		};
		RangeView.prototype.update=function (properties) {
			return _invokeRecursiveUpdate(this, properties);
		};
		RangeView.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		RangeView.prototype.toJSON=function () {
			return {};
		};
		return RangeView;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.RangeView=RangeView;
	var RangeViewCollection=(function (_super) {
		__extends(RangeViewCollection, _super);
		function RangeViewCollection() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(RangeViewCollection.prototype, "_className", {
			get: function () {
				return "RangeViewCollection";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(RangeViewCollection.prototype, "_isCollection", {
			get: function () {
				return true;
			},
			enumerable: true,
			configurable: true
		});
		RangeViewCollection.prototype.getItemAt=function (index) {
			return _createMethodObject(ExcelOp.RangeView, this, "GetItemAt", 1, [index], false, false, null, 4);
		};
		RangeViewCollection.prototype.getCount=function () {
			_throwIfApiNotSupported("RangeViewCollection.getCount", _defaultApiSetName, "1.4", _hostName);
			return _invokeMethod(this, "GetCount", 1, [], 4, 0);
		};
		RangeViewCollection.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		RangeViewCollection.prototype.toJSON=function () {
			return {};
		};
		return RangeViewCollection;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.RangeViewCollection=RangeViewCollection;
	var SettingCollection=(function (_super) {
		__extends(SettingCollection, _super);
		function SettingCollection() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(SettingCollection.prototype, "_className", {
			get: function () {
				return "SettingCollection";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(SettingCollection.prototype, "_isCollection", {
			get: function () {
				return true;
			},
			enumerable: true,
			configurable: true
		});
		SettingCollection.prototype.add=function (key, value) {
			return _createAndInstantiateMethodObject(ExcelOp.Setting, this, "Add", 0, [key, value], false, true, null, 0);
		};
		SettingCollection.prototype.getItem=function (key) {
			return _createIndexerObject(ExcelOp.Setting, this, [key]);
		};
		SettingCollection.prototype.getItemOrNullObject=function (key) {
			return _createMethodObject(ExcelOp.Setting, this, "GetItemOrNullObject", 1, [key], false, false, null, 4);
		};
		SettingCollection.prototype.getCount=function () {
			return _invokeMethod(this, "GetCount", 1, [], 4, 0);
		};
		SettingCollection.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		SettingCollection.prototype.toJSON=function () {
			return {};
		};
		return SettingCollection;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.SettingCollection=SettingCollection;
	var Setting=(function (_super) {
		__extends(Setting, _super);
		function Setting() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(Setting.prototype, "_className", {
			get: function () {
				return "Setting";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Setting.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["key", "value"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Setting.prototype, "_scalarPropertyUpdateable", {
			get: function () {
				return [false, true];
			},
			enumerable: true,
			configurable: true
		});
		Setting.prototype.update=function (properties) {
			return _invokeRecursiveUpdate(this, properties);
		};
		Setting.prototype["delete"]=function () {
			return _invokeMethod(this, "Delete", 0, [], 0);
		};
		Setting.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		Setting.prototype.toJSON=function () {
			return {};
		};
		return Setting;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.Setting=Setting;
	var NamedItemCollection=(function (_super) {
		__extends(NamedItemCollection, _super);
		function NamedItemCollection() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(NamedItemCollection.prototype, "_className", {
			get: function () {
				return "NamedItemCollection";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(NamedItemCollection.prototype, "_isCollection", {
			get: function () {
				return true;
			},
			enumerable: true,
			configurable: true
		});
		NamedItemCollection.prototype.add=function (name, reference, comment) {
			_throwIfApiNotSupported("NamedItemCollection.add", _defaultApiSetName, "1.4", _hostName);
			return _createAndInstantiateMethodObject(ExcelOp.NamedItem, this, "Add", 0, [name, reference, comment], false, true, null, 0);
		};
		NamedItemCollection.prototype.addFormulaLocal=function (name, formula, comment) {
			_throwIfApiNotSupported("NamedItemCollection.addFormulaLocal", _defaultApiSetName, "1.4", _hostName);
			return _createAndInstantiateMethodObject(ExcelOp.NamedItem, this, "AddFormulaLocal", 0, [name, formula, comment], false, false, null, 0);
		};
		NamedItemCollection.prototype.getItem=function (name) {
			return _createIndexerObject(ExcelOp.NamedItem, this, [name]);
		};
		NamedItemCollection.prototype.getItemOrNullObject=function (name) {
			_throwIfApiNotSupported("NamedItemCollection.getItemOrNullObject", _defaultApiSetName, "1.4", _hostName);
			return _createMethodObject(ExcelOp.NamedItem, this, "GetItemOrNullObject", 1, [name], false, false, null, 4);
		};
		NamedItemCollection.prototype.getCount=function () {
			_throwIfApiNotSupported("NamedItemCollection.getCount", _defaultApiSetName, "1.4", _hostName);
			return _invokeMethod(this, "GetCount", 1, [], 4, 0);
		};
		NamedItemCollection.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		NamedItemCollection.prototype.toJSON=function () {
			return {};
		};
		return NamedItemCollection;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.NamedItemCollection=NamedItemCollection;
	var NamedItem=(function (_super) {
		__extends(NamedItem, _super);
		function NamedItem() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(NamedItem.prototype, "_className", {
			get: function () {
				return "NamedItem";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(NamedItem.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["name", "type", "value", "visible", "_Id", "comment", "scope", "formula"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(NamedItem.prototype, "_scalarPropertyUpdateable", {
			get: function () {
				return [false, false, false, true, false, true, false, true];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(NamedItem.prototype, "_navigationPropertyNames", {
			get: function () {
				return ["worksheet", "worksheetOrNullObject", "arrayValues"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(NamedItem.prototype, "arrayValues", {
			get: function () {
				_throwIfApiNotSupported("NamedItem.arrayValues", _defaultApiSetName, "1.7", _hostName);
				return _createPropertyObject(ExcelOp.NamedItemArrayValues, this, "ArrayValues", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(NamedItem.prototype, "worksheet", {
			get: function () {
				_throwIfApiNotSupported("NamedItem.worksheet", _defaultApiSetName, "1.4", _hostName);
				return _createPropertyObject(ExcelOp.Worksheet, this, "Worksheet", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(NamedItem.prototype, "worksheetOrNullObject", {
			get: function () {
				_throwIfApiNotSupported("NamedItem.worksheetOrNullObject", _defaultApiSetName, "1.4", _hostName);
				return _createPropertyObject(ExcelOp.Worksheet, this, "WorksheetOrNullObject", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		NamedItem.prototype.getRange=function () {
			return _createMethodObject(ExcelOp.Range, this, "GetRange", 1, [], false, true, null, 4);
		};
		NamedItem.prototype.getRangeOrNullObject=function () {
			_throwIfApiNotSupported("NamedItem.getRangeOrNullObject", _defaultApiSetName, "1.4", _hostName);
			return _createMethodObject(ExcelOp.Range, this, "GetRangeOrNullObject", 1, [], false, true, null, 4);
		};
		NamedItem.prototype.update=function (properties) {
			return _invokeRecursiveUpdate(this, properties);
		};
		NamedItem.prototype["delete"]=function () {
			_throwIfApiNotSupported("NamedItem.delete", _defaultApiSetName, "1.4", _hostName);
			return _invokeMethod(this, "Delete", 0, [], 0);
		};
		NamedItem.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		NamedItem.prototype.toJSON=function () {
			return {};
		};
		return NamedItem;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.NamedItem=NamedItem;
	var NamedItemArrayValues=(function (_super) {
		__extends(NamedItemArrayValues, _super);
		function NamedItemArrayValues() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(NamedItemArrayValues.prototype, "_className", {
			get: function () {
				return "NamedItemArrayValues";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(NamedItemArrayValues.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["values", "types"];
			},
			enumerable: true,
			configurable: true
		});
		NamedItemArrayValues.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		NamedItemArrayValues.prototype.toJSON=function () {
			return {};
		};
		return NamedItemArrayValues;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.NamedItemArrayValues=NamedItemArrayValues;
	var Binding=(function (_super) {
		__extends(Binding, _super);
		function Binding() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(Binding.prototype, "_className", {
			get: function () {
				return "Binding";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Binding.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["id", "type"];
			},
			enumerable: true,
			configurable: true
		});
		Binding.prototype.getRange=function () {
			return _createMethodObject(ExcelOp.Range, this, "GetRange", 1, [], false, false, null, 4);
		};
		Binding.prototype.getTable=function () {
			return _createMethodObject(ExcelOp.Table, this, "GetTable", 1, [], false, false, null, 4);
		};
		Binding.prototype["delete"]=function () {
			_throwIfApiNotSupported("Binding.delete", _defaultApiSetName, "1.3", _hostName);
			return _invokeMethod(this, "Delete", 0, [], 0);
		};
		Binding.prototype.getText=function () {
			return _invokeMethod(this, "GetText", 1, [], 4, 0);
		};
		Binding.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		Binding.prototype.toJSON=function () {
			return {};
		};
		return Binding;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.Binding=Binding;
	var BindingCollection=(function (_super) {
		__extends(BindingCollection, _super);
		function BindingCollection() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(BindingCollection.prototype, "_className", {
			get: function () {
				return "BindingCollection";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(BindingCollection.prototype, "_isCollection", {
			get: function () {
				return true;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(BindingCollection.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["count"];
			},
			enumerable: true,
			configurable: true
		});
		BindingCollection.prototype.add=function (range, bindingType, id) {
			_throwIfApiNotSupported("BindingCollection.add", _defaultApiSetName, "1.3", _hostName);
			return _createAndInstantiateMethodObject(ExcelOp.Binding, this, "Add", 0, [range, bindingType, id], false, true, null, 0);
		};
		BindingCollection.prototype.addFromNamedItem=function (name, bindingType, id) {
			_throwIfApiNotSupported("BindingCollection.addFromNamedItem", _defaultApiSetName, "1.3", _hostName);
			return _createAndInstantiateMethodObject(ExcelOp.Binding, this, "AddFromNamedItem", 0, [name, bindingType, id], false, false, null, 0);
		};
		BindingCollection.prototype.addFromSelection=function (bindingType, id) {
			_throwIfApiNotSupported("BindingCollection.addFromSelection", _defaultApiSetName, "1.3", _hostName);
			return _createAndInstantiateMethodObject(ExcelOp.Binding, this, "AddFromSelection", 0, [bindingType, id], false, false, null, 0);
		};
		BindingCollection.prototype.getItem=function (id) {
			return _createIndexerObject(ExcelOp.Binding, this, [id]);
		};
		BindingCollection.prototype.getItemAt=function (index) {
			return _createMethodObject(ExcelOp.Binding, this, "GetItemAt", 1, [index], false, false, null, 4);
		};
		BindingCollection.prototype.getItemOrNullObject=function (id) {
			_throwIfApiNotSupported("BindingCollection.getItemOrNullObject", _defaultApiSetName, "1.4", _hostName);
			return _createMethodObject(ExcelOp.Binding, this, "GetItemOrNullObject", 1, [id], false, false, null, 4);
		};
		BindingCollection.prototype.getCount=function () {
			_throwIfApiNotSupported("BindingCollection.getCount", _defaultApiSetName, "1.4", _hostName);
			return _invokeMethod(this, "GetCount", 1, [], 4, 0);
		};
		BindingCollection.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		BindingCollection.prototype.toJSON=function () {
			return {};
		};
		return BindingCollection;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.BindingCollection=BindingCollection;
	var TableCollection=(function (_super) {
		__extends(TableCollection, _super);
		function TableCollection() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(TableCollection.prototype, "_className", {
			get: function () {
				return "TableCollection";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableCollection.prototype, "_isCollection", {
			get: function () {
				return true;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableCollection.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["count"];
			},
			enumerable: true,
			configurable: true
		});
		TableCollection.prototype.add=function (address, hasHeaders) {
			return _createAndInstantiateMethodObject(ExcelOp.Table, this, "Add", 0, [address, hasHeaders], false, true, null, 0);
		};
		TableCollection.prototype.getItem=function (key) {
			return _createIndexerObject(ExcelOp.Table, this, [key]);
		};
		TableCollection.prototype.getItemAt=function (index) {
			return _createMethodObject(ExcelOp.Table, this, "GetItemAt", 1, [index], false, false, null, 4);
		};
		TableCollection.prototype.getItemOrNullObject=function (key) {
			_throwIfApiNotSupported("TableCollection.getItemOrNullObject", _defaultApiSetName, "1.4", _hostName);
			return _createMethodObject(ExcelOp.Table, this, "GetItemOrNullObject", 1, [key], false, false, null, 4);
		};
		TableCollection.prototype.getCount=function () {
			_throwIfApiNotSupported("TableCollection.getCount", _defaultApiSetName, "1.4", _hostName);
			return _invokeMethod(this, "GetCount", 1, [], 4, 0);
		};
		TableCollection.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		TableCollection.prototype.toJSON=function () {
			return {};
		};
		return TableCollection;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.TableCollection=TableCollection;
	var TableScopedCollection=(function (_super) {
		__extends(TableScopedCollection, _super);
		function TableScopedCollection() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(TableScopedCollection.prototype, "_className", {
			get: function () {
				return "TableScopedCollection";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableScopedCollection.prototype, "_isCollection", {
			get: function () {
				return true;
			},
			enumerable: true,
			configurable: true
		});
		TableScopedCollection.prototype.getFirst=function () {
			return _createMethodObject(ExcelOp.Table, this, "GetFirst", 1, [], false, true, null, 4);
		};
		TableScopedCollection.prototype.getItem=function (key) {
			return _createIndexerObject(ExcelOp.Table, this, [key]);
		};
		TableScopedCollection.prototype.getCount=function () {
			return _invokeMethod(this, "GetCount", 1, [], 4, 0);
		};
		TableScopedCollection.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		TableScopedCollection.prototype.toJSON=function () {
			return {};
		};
		return TableScopedCollection;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.TableScopedCollection=TableScopedCollection;
	var Table=(function (_super) {
		__extends(Table, _super);
		function Table() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(Table.prototype, "_className", {
			get: function () {
				return "Table";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Table.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["id", "name", "showHeaders", "showTotals", "style", "highlightFirstColumn", "highlightLastColumn", "showBandedRows", "showBandedColumns", "showFilterButton", "legacyId"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Table.prototype, "_scalarPropertyUpdateable", {
			get: function () {
				return [false, true, true, true, true, true, true, true, true, true, false];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Table.prototype, "_navigationPropertyNames", {
			get: function () {
				return ["columns", "rows", "sort", "worksheet", "autoFilter"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Table.prototype, "autoFilter", {
			get: function () {
				_throwIfApiNotSupported("Table.autoFilter", _defaultApiSetName, "1.9", _hostName);
				return _createPropertyObject(ExcelOp.AutoFilter, this, "AutoFilter", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Table.prototype, "columns", {
			get: function () {
				return _createPropertyObject(ExcelOp.TableColumnCollection, this, "Columns", true, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Table.prototype, "rows", {
			get: function () {
				return _createPropertyObject(ExcelOp.TableRowCollection, this, "Rows", true, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Table.prototype, "sort", {
			get: function () {
				_throwIfApiNotSupported("Table.sort", _defaultApiSetName, "1.2", _hostName);
				return _createPropertyObject(ExcelOp.TableSort, this, "Sort", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Table.prototype, "worksheet", {
			get: function () {
				_throwIfApiNotSupported("Table.worksheet", _defaultApiSetName, "1.2", _hostName);
				return _createPropertyObject(ExcelOp.Worksheet, this, "Worksheet", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Table.prototype.convertToRange=function () {
			_throwIfApiNotSupported("Table.convertToRange", _defaultApiSetName, "1.2", _hostName);
			return _createAndInstantiateMethodObject(ExcelOp.Range, this, "ConvertToRange", 0, [], false, true, null, 0);
		};
		Table.prototype.getDataBodyRange=function () {
			return _createMethodObject(ExcelOp.Range, this, "GetDataBodyRange", 1, [], false, true, null, 4);
		};
		Table.prototype.getHeaderRowRange=function () {
			return _createMethodObject(ExcelOp.Range, this, "GetHeaderRowRange", 1, [], false, true, null, 4);
		};
		Table.prototype.getRange=function () {
			return _createMethodObject(ExcelOp.Range, this, "GetRange", 1, [], false, true, null, 4);
		};
		Table.prototype.getTotalRowRange=function () {
			return _createMethodObject(ExcelOp.Range, this, "GetTotalRowRange", 1, [], false, true, null, 4);
		};
		Table.prototype.update=function (properties) {
			return _invokeRecursiveUpdate(this, properties);
		};
		Table.prototype.clearFilters=function () {
			_throwIfApiNotSupported("Table.clearFilters", _defaultApiSetName, "1.2", _hostName);
			return _invokeMethod(this, "ClearFilters", 0, [], 0);
		};
		Table.prototype.clearStyle=function () {
			_throwIfApiNotSupported("Table.clearStyle", _defaultApiSetName, "1.9", _hostName);
			return _invokeMethod(this, "ClearStyle", 0, [], 0);
		};
		Table.prototype["delete"]=function () {
			return _invokeMethod(this, "Delete", 0, [], 0);
		};
		Table.prototype.reapplyFilters=function () {
			_throwIfApiNotSupported("Table.reapplyFilters", _defaultApiSetName, "1.2", _hostName);
			return _invokeMethod(this, "ReapplyFilters", 0, [], 0);
		};
		Table.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		Table.prototype.toJSON=function () {
			return {};
		};
		return Table;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.Table=Table;
	var TableColumnCollection=(function (_super) {
		__extends(TableColumnCollection, _super);
		function TableColumnCollection() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(TableColumnCollection.prototype, "_className", {
			get: function () {
				return "TableColumnCollection";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableColumnCollection.prototype, "_isCollection", {
			get: function () {
				return true;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableColumnCollection.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["count"];
			},
			enumerable: true,
			configurable: true
		});
		TableColumnCollection.prototype.add=function (index, values, name) {
			return _createAndInstantiateMethodObject(ExcelOp.TableColumn, this, "Add", 0, [index, values, name], false, true, null, 0);
		};
		TableColumnCollection.prototype.getItem=function (key) {
			return _createIndexerObject(ExcelOp.TableColumn, this, [key]);
		};
		TableColumnCollection.prototype.getItemAt=function (index) {
			return _createMethodObject(ExcelOp.TableColumn, this, "GetItemAt", 1, [index], false, false, null, 4);
		};
		TableColumnCollection.prototype.getItemOrNullObject=function (key) {
			_throwIfApiNotSupported("TableColumnCollection.getItemOrNullObject", _defaultApiSetName, "1.4", _hostName);
			return _createMethodObject(ExcelOp.TableColumn, this, "GetItemOrNullObject", 1, [key], false, false, null, 4);
		};
		TableColumnCollection.prototype.getCount=function () {
			_throwIfApiNotSupported("TableColumnCollection.getCount", _defaultApiSetName, "1.4", _hostName);
			return _invokeMethod(this, "GetCount", 1, [], 4, 0);
		};
		TableColumnCollection.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		TableColumnCollection.prototype.toJSON=function () {
			return {};
		};
		return TableColumnCollection;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.TableColumnCollection=TableColumnCollection;
	var TableColumn=(function (_super) {
		__extends(TableColumn, _super);
		function TableColumn() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(TableColumn.prototype, "_className", {
			get: function () {
				return "TableColumn";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableColumn.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["id", "index", "values", "name"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableColumn.prototype, "_scalarPropertyUpdateable", {
			get: function () {
				return [false, false, true, true];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableColumn.prototype, "_navigationPropertyNames", {
			get: function () {
				return ["filter"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableColumn.prototype, "filter", {
			get: function () {
				_throwIfApiNotSupported("TableColumn.filter", _defaultApiSetName, "1.2", _hostName);
				return _createPropertyObject(ExcelOp.Filter, this, "Filter", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		TableColumn.prototype.getDataBodyRange=function () {
			return _createMethodObject(ExcelOp.Range, this, "GetDataBodyRange", 1, [], false, true, null, 4);
		};
		TableColumn.prototype.getHeaderRowRange=function () {
			return _createMethodObject(ExcelOp.Range, this, "GetHeaderRowRange", 1, [], false, true, null, 4);
		};
		TableColumn.prototype.getRange=function () {
			return _createMethodObject(ExcelOp.Range, this, "GetRange", 1, [], false, true, null, 4);
		};
		TableColumn.prototype.getTotalRowRange=function () {
			return _createMethodObject(ExcelOp.Range, this, "GetTotalRowRange", 1, [], false, true, null, 4);
		};
		TableColumn.prototype.update=function (properties) {
			return _invokeRecursiveUpdate(this, properties);
		};
		TableColumn.prototype["delete"]=function () {
			return _invokeMethod(this, "Delete", 0, [], 0);
		};
		TableColumn.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		TableColumn.prototype.toJSON=function () {
			return {};
		};
		return TableColumn;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.TableColumn=TableColumn;
	var TableRowCollection=(function (_super) {
		__extends(TableRowCollection, _super);
		function TableRowCollection() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(TableRowCollection.prototype, "_className", {
			get: function () {
				return "TableRowCollection";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableRowCollection.prototype, "_isCollection", {
			get: function () {
				return true;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableRowCollection.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["count"];
			},
			enumerable: true,
			configurable: true
		});
		TableRowCollection.prototype.add=function (index, values) {
			return _createAndInstantiateMethodObject(ExcelOp.TableRow, this, "Add", 0, [index, values], false, true, null, 0);
		};
		TableRowCollection.prototype.getItemAt=function (index) {
			return _createMethodObject(ExcelOp.TableRow, this, "GetItemAt", 1, [index], false, false, null, 4);
		};
		TableRowCollection.prototype.getCount=function () {
			_throwIfApiNotSupported("TableRowCollection.getCount", _defaultApiSetName, "1.4", _hostName);
			return _invokeMethod(this, "GetCount", 1, [], 4, 0);
		};
		TableRowCollection.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		TableRowCollection.prototype.toJSON=function () {
			return {};
		};
		return TableRowCollection;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.TableRowCollection=TableRowCollection;
	var TableRow=(function (_super) {
		__extends(TableRow, _super);
		function TableRow() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(TableRow.prototype, "_className", {
			get: function () {
				return "TableRow";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableRow.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["index", "values"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableRow.prototype, "_scalarPropertyUpdateable", {
			get: function () {
				return [false, true];
			},
			enumerable: true,
			configurable: true
		});
		TableRow.prototype.getRange=function () {
			return _createMethodObject(ExcelOp.Range, this, "GetRange", 1, [], false, true, null, 4);
		};
		TableRow.prototype.update=function (properties) {
			return _invokeRecursiveUpdate(this, properties);
		};
		TableRow.prototype["delete"]=function () {
			return _invokeMethod(this, "Delete", 0, [], 0);
		};
		TableRow.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		TableRow.prototype.toJSON=function () {
			return {};
		};
		return TableRow;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.TableRow=TableRow;
	var DataValidation=(function (_super) {
		__extends(DataValidation, _super);
		function DataValidation() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(DataValidation.prototype, "_className", {
			get: function () {
				return "DataValidation";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(DataValidation.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["type", "rule", "prompt", "errorAlert", "ignoreBlanks", "valid"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(DataValidation.prototype, "_scalarPropertyUpdateable", {
			get: function () {
				return [false, true, true, true, true, false];
			},
			enumerable: true,
			configurable: true
		});
		DataValidation.prototype.getInvalidCells=function () {
			_throwIfApiNotSupported("DataValidation.getInvalidCells", _defaultApiSetName, "1.9", _hostName);
			return _createMethodObject(ExcelOp.RangeAreas, this, "GetInvalidCells", 1, [], false, true, null, 4);
		};
		DataValidation.prototype.getInvalidCellsOrNullObject=function () {
			_throwIfApiNotSupported("DataValidation.getInvalidCellsOrNullObject", _defaultApiSetName, "1.9", _hostName);
			return _createMethodObject(ExcelOp.RangeAreas, this, "GetInvalidCellsOrNullObject", 1, [], false, true, null, 4);
		};
		DataValidation.prototype.update=function (properties) {
			return _invokeRecursiveUpdate(this, properties);
		};
		DataValidation.prototype.clear=function () {
			return _invokeMethod(this, "Clear", 0, [], 0);
		};
		DataValidation.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		DataValidation.prototype.toJSON=function () {
			return {};
		};
		return DataValidation;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.DataValidation=DataValidation;
	var RemoveDuplicatesResult=(function (_super) {
		__extends(RemoveDuplicatesResult, _super);
		function RemoveDuplicatesResult() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(RemoveDuplicatesResult.prototype, "_className", {
			get: function () {
				return "RemoveDuplicatesResult";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(RemoveDuplicatesResult.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["removed", "uniqueRemaining"];
			},
			enumerable: true,
			configurable: true
		});
		RemoveDuplicatesResult.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		RemoveDuplicatesResult.prototype.toJSON=function () {
			return {};
		};
		return RemoveDuplicatesResult;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.RemoveDuplicatesResult=RemoveDuplicatesResult;
	var RangeFormat=(function (_super) {
		__extends(RangeFormat, _super);
		function RangeFormat() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(RangeFormat.prototype, "_className", {
			get: function () {
				return "RangeFormat";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(RangeFormat.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["wrapText", "horizontalAlignment", "verticalAlignment", "columnWidth", "rowHeight", "textOrientation", "useStandardHeight", "useStandardWidth", "readingOrder", "shrinkToFit", "indentLevel", "autoIndent"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(RangeFormat.prototype, "_scalarPropertyUpdateable", {
			get: function () {
				return [true, true, true, true, true, true, true, true, true, true, true, true];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(RangeFormat.prototype, "_navigationPropertyNames", {
			get: function () {
				return ["fill", "font", "borders", "protection"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(RangeFormat.prototype, "borders", {
			get: function () {
				return _createPropertyObject(ExcelOp.RangeBorderCollection, this, "Borders", true, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(RangeFormat.prototype, "fill", {
			get: function () {
				return _createPropertyObject(ExcelOp.RangeFill, this, "Fill", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(RangeFormat.prototype, "font", {
			get: function () {
				return _createPropertyObject(ExcelOp.RangeFont, this, "Font", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(RangeFormat.prototype, "protection", {
			get: function () {
				_throwIfApiNotSupported("RangeFormat.protection", _defaultApiSetName, "1.2", _hostName);
				return _createPropertyObject(ExcelOp.FormatProtection, this, "Protection", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		RangeFormat.prototype.update=function (properties) {
			return _invokeRecursiveUpdate(this, properties);
		};
		RangeFormat.prototype.autofitColumns=function () {
			_throwIfApiNotSupported("RangeFormat.autofitColumns", _defaultApiSetName, "1.2", _hostName);
			return _invokeMethod(this, "AutofitColumns", 0, [], 0);
		};
		RangeFormat.prototype.autofitRows=function () {
			_throwIfApiNotSupported("RangeFormat.autofitRows", _defaultApiSetName, "1.2", _hostName);
			return _invokeMethod(this, "AutofitRows", 0, [], 0);
		};
		RangeFormat.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		RangeFormat.prototype.toJSON=function () {
			return {};
		};
		return RangeFormat;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.RangeFormat=RangeFormat;
	var FormatProtection=(function (_super) {
		__extends(FormatProtection, _super);
		function FormatProtection() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(FormatProtection.prototype, "_className", {
			get: function () {
				return "FormatProtection";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(FormatProtection.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["locked", "formulaHidden"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(FormatProtection.prototype, "_scalarPropertyUpdateable", {
			get: function () {
				return [true, true];
			},
			enumerable: true,
			configurable: true
		});
		FormatProtection.prototype.update=function (properties) {
			return _invokeRecursiveUpdate(this, properties);
		};
		FormatProtection.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		FormatProtection.prototype.toJSON=function () {
			return {};
		};
		return FormatProtection;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.FormatProtection=FormatProtection;
	var RangeFill=(function (_super) {
		__extends(RangeFill, _super);
		function RangeFill() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(RangeFill.prototype, "_className", {
			get: function () {
				return "RangeFill";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(RangeFill.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["color", "tintAndShade", "patternTintAndShade", "pattern", "patternColor"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(RangeFill.prototype, "_scalarPropertyUpdateable", {
			get: function () {
				return [true, true, true, true, true];
			},
			enumerable: true,
			configurable: true
		});
		RangeFill.prototype.update=function (properties) {
			return _invokeRecursiveUpdate(this, properties);
		};
		RangeFill.prototype.clear=function () {
			return _invokeMethod(this, "Clear", 0, [], 0);
		};
		RangeFill.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		RangeFill.prototype.toJSON=function () {
			return {};
		};
		return RangeFill;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.RangeFill=RangeFill;
	var RangeBorder=(function (_super) {
		__extends(RangeBorder, _super);
		function RangeBorder() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(RangeBorder.prototype, "_className", {
			get: function () {
				return "RangeBorder";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(RangeBorder.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["sideIndex", "style", "weight", "color", "tintAndShade"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(RangeBorder.prototype, "_scalarPropertyUpdateable", {
			get: function () {
				return [false, true, true, true, true];
			},
			enumerable: true,
			configurable: true
		});
		RangeBorder.prototype.update=function (properties) {
			return _invokeRecursiveUpdate(this, properties);
		};
		RangeBorder.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		RangeBorder.prototype.toJSON=function () {
			return {};
		};
		return RangeBorder;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.RangeBorder=RangeBorder;
	var RangeBorderCollection=(function (_super) {
		__extends(RangeBorderCollection, _super);
		function RangeBorderCollection() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(RangeBorderCollection.prototype, "_className", {
			get: function () {
				return "RangeBorderCollection";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(RangeBorderCollection.prototype, "_isCollection", {
			get: function () {
				return true;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(RangeBorderCollection.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["count", "tintAndShade"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(RangeBorderCollection.prototype, "_scalarPropertyUpdateable", {
			get: function () {
				return [false, true];
			},
			enumerable: true,
			configurable: true
		});
		RangeBorderCollection.prototype.getItem=function (index) {
			return _createIndexerObject(ExcelOp.RangeBorder, this, [index]);
		};
		RangeBorderCollection.prototype.getItemAt=function (index) {
			return _createMethodObject(ExcelOp.RangeBorder, this, "GetItemAt", 1, [index], false, false, null, 4);
		};
		RangeBorderCollection.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		RangeBorderCollection.prototype.toJSON=function () {
			return {};
		};
		return RangeBorderCollection;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.RangeBorderCollection=RangeBorderCollection;
	var RangeFont=(function (_super) {
		__extends(RangeFont, _super);
		function RangeFont() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(RangeFont.prototype, "_className", {
			get: function () {
				return "RangeFont";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(RangeFont.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["name", "size", "color", "italic", "bold", "underline", "strikethrough", "subscript", "superscript", "tintAndShade"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(RangeFont.prototype, "_scalarPropertyUpdateable", {
			get: function () {
				return [true, true, true, true, true, true, true, true, true, true];
			},
			enumerable: true,
			configurable: true
		});
		RangeFont.prototype.update=function (properties) {
			return _invokeRecursiveUpdate(this, properties);
		};
		RangeFont.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		RangeFont.prototype.toJSON=function () {
			return {};
		};
		return RangeFont;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.RangeFont=RangeFont;
	var ChartCollection=(function (_super) {
		__extends(ChartCollection, _super);
		function ChartCollection() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(ChartCollection.prototype, "_className", {
			get: function () {
				return "ChartCollection";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartCollection.prototype, "_isCollection", {
			get: function () {
				return true;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartCollection.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["count"];
			},
			enumerable: true,
			configurable: true
		});
		ChartCollection.prototype.add=function (type, sourceData, seriesBy) {
			return _createAndInstantiateMethodObject(ExcelOp.Chart, this, "Add", 0, [type, sourceData, seriesBy], false, true, null, 0);
		};
		ChartCollection.prototype.getItem=function (name) {
			return _createMethodObject(ExcelOp.Chart, this, "GetItem", 1, [name], false, false, null, 4);
		};
		ChartCollection.prototype.getItemAt=function (index) {
			return _createMethodObject(ExcelOp.Chart, this, "GetItemAt", 1, [index], false, false, null, 4);
		};
		ChartCollection.prototype.getItemOrNullObject=function (name) {
			_throwIfApiNotSupported("ChartCollection.getItemOrNullObject", _defaultApiSetName, "1.4", _hostName);
			return _createMethodObject(ExcelOp.Chart, this, "GetItemOrNullObject", 1, [name], false, false, null, 4);
		};
		ChartCollection.prototype._GetItem=function (key) {
			return _createIndexerObject(ExcelOp.Chart, this, [key]);
		};
		ChartCollection.prototype.getCount=function () {
			_throwIfApiNotSupported("ChartCollection.getCount", _defaultApiSetName, "1.4", _hostName);
			return _invokeMethod(this, "GetCount", 1, [], 4, 0);
		};
		ChartCollection.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		ChartCollection.prototype.toJSON=function () {
			return {};
		};
		return ChartCollection;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.ChartCollection=ChartCollection;
	var Chart=(function (_super) {
		__extends(Chart, _super);
		function Chart() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(Chart.prototype, "_className", {
			get: function () {
				return "Chart";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Chart.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["name", "top", "left", "width", "height", "id", "showAllFieldButtons", "chartType", "showDataLabelsOverMaximum", "categoryLabelLevel", "style", "displayBlanksAs", "plotBy", "plotVisibleOnly", "seriesNameLevel"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Chart.prototype, "_scalarPropertyUpdateable", {
			get: function () {
				return [true, true, true, true, true, false, true, true, true, true, true, true, true, true, true];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Chart.prototype, "_navigationPropertyNames", {
			get: function () {
				return ["title", "dataLabels", "legend", "series", "axes", "format", "worksheet", "plotArea", "pivotOptions"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Chart.prototype, "axes", {
			get: function () {
				return _createPropertyObject(ExcelOp.ChartAxes, this, "Axes", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Chart.prototype, "dataLabels", {
			get: function () {
				return _createPropertyObject(ExcelOp.ChartDataLabels, this, "DataLabels", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Chart.prototype, "format", {
			get: function () {
				return _createPropertyObject(ExcelOp.ChartAreaFormat, this, "Format", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Chart.prototype, "legend", {
			get: function () {
				return _createPropertyObject(ExcelOp.ChartLegend, this, "Legend", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Chart.prototype, "pivotOptions", {
			get: function () {
				_throwIfApiNotSupported("Chart.pivotOptions", _defaultApiSetName, "1.9", _hostName);
				return _createPropertyObject(ExcelOp.ChartPivotOptions, this, "PivotOptions", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Chart.prototype, "plotArea", {
			get: function () {
				_throwIfApiNotSupported("Chart.plotArea", _defaultApiSetName, "1.8", _hostName);
				return _createPropertyObject(ExcelOp.ChartPlotArea, this, "PlotArea", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Chart.prototype, "series", {
			get: function () {
				return _createPropertyObject(ExcelOp.ChartSeriesCollection, this, "Series", true, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Chart.prototype, "title", {
			get: function () {
				return _createPropertyObject(ExcelOp.ChartTitle, this, "Title", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Chart.prototype, "worksheet", {
			get: function () {
				_throwIfApiNotSupported("Chart.worksheet", _defaultApiSetName, "1.2", _hostName);
				return _createPropertyObject(ExcelOp.Worksheet, this, "Worksheet", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Chart.prototype.update=function (properties) {
			return _invokeRecursiveUpdate(this, properties);
		};
		Chart.prototype.activate=function () {
			_throwIfApiNotSupported("Chart.activate", _defaultApiSetName, "1.9", _hostName);
			return _invokeMethod(this, "Activate", 1, [], 0);
		};
		Chart.prototype["delete"]=function () {
			return _invokeMethod(this, "Delete", 0, [], 0);
		};
		Chart.prototype.getImage=function (width, height, fittingMode) {
			_throwIfApiNotSupported("Chart.getImage", _defaultApiSetName, "1.2", _hostName);
			return _invokeMethod(this, "GetImage", 1, [width, height, fittingMode], 4, 0);
		};
		Chart.prototype.setData=function (sourceData, seriesBy) {
			return _invokeMethod(this, "SetData", 0, [sourceData, seriesBy], 0);
		};
		Chart.prototype.setPosition=function (startCell, endCell) {
			return _invokeMethod(this, "SetPosition", 0, [startCell, endCell], 0);
		};
		Chart.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		Chart.prototype.toJSON=function () {
			return {};
		};
		return Chart;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.Chart=Chart;
	var ChartPivotOptions=(function (_super) {
		__extends(ChartPivotOptions, _super);
		function ChartPivotOptions() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(ChartPivotOptions.prototype, "_className", {
			get: function () {
				return "ChartPivotOptions";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartPivotOptions.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["showAxisFieldButtons", "showLegendFieldButtons", "showReportFilterFieldButtons", "showValueFieldButtons"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartPivotOptions.prototype, "_scalarPropertyUpdateable", {
			get: function () {
				return [true, true, true, true];
			},
			enumerable: true,
			configurable: true
		});
		ChartPivotOptions.prototype.update=function (properties) {
			return _invokeRecursiveUpdate(this, properties);
		};
		ChartPivotOptions.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		ChartPivotOptions.prototype.toJSON=function () {
			return {};
		};
		return ChartPivotOptions;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.ChartPivotOptions=ChartPivotOptions;
	var ChartAreaFormat=(function (_super) {
		__extends(ChartAreaFormat, _super);
		function ChartAreaFormat() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(ChartAreaFormat.prototype, "_className", {
			get: function () {
				return "ChartAreaFormat";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartAreaFormat.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["roundedCorners", "colorScheme"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartAreaFormat.prototype, "_scalarPropertyUpdateable", {
			get: function () {
				return [true, true];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartAreaFormat.prototype, "_navigationPropertyNames", {
			get: function () {
				return ["fill", "font", "border"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartAreaFormat.prototype, "border", {
			get: function () {
				_throwIfApiNotSupported("ChartAreaFormat.border", _defaultApiSetName, "1.7", _hostName);
				return _createPropertyObject(ExcelOp.ChartBorder, this, "Border", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartAreaFormat.prototype, "fill", {
			get: function () {
				return _createPropertyObject(ExcelOp.ChartFill, this, "Fill", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartAreaFormat.prototype, "font", {
			get: function () {
				return _createPropertyObject(ExcelOp.ChartFont, this, "Font", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		ChartAreaFormat.prototype.update=function (properties) {
			return _invokeRecursiveUpdate(this, properties);
		};
		ChartAreaFormat.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		ChartAreaFormat.prototype.toJSON=function () {
			return {};
		};
		return ChartAreaFormat;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.ChartAreaFormat=ChartAreaFormat;
	var ChartSeriesCollection=(function (_super) {
		__extends(ChartSeriesCollection, _super);
		function ChartSeriesCollection() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(ChartSeriesCollection.prototype, "_className", {
			get: function () {
				return "ChartSeriesCollection";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartSeriesCollection.prototype, "_isCollection", {
			get: function () {
				return true;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartSeriesCollection.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["count"];
			},
			enumerable: true,
			configurable: true
		});
		ChartSeriesCollection.prototype.add=function (name, index) {
			_throwIfApiNotSupported("ChartSeriesCollection.add", _defaultApiSetName, "1.7", _hostName);
			return _createAndInstantiateMethodObject(ExcelOp.ChartSeries, this, "Add", 0, [name, index], false, true, null, 0);
		};
		ChartSeriesCollection.prototype.getItemAt=function (index) {
			return _createMethodObject(ExcelOp.ChartSeries, this, "GetItemAt", 1, [index], false, false, null, 4);
		};
		ChartSeriesCollection.prototype.getCount=function () {
			_throwIfApiNotSupported("ChartSeriesCollection.getCount", _defaultApiSetName, "1.4", _hostName);
			return _invokeMethod(this, "GetCount", 1, [], 4, 0);
		};
		ChartSeriesCollection.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		ChartSeriesCollection.prototype.toJSON=function () {
			return {};
		};
		return ChartSeriesCollection;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.ChartSeriesCollection=ChartSeriesCollection;
	var ChartSeries=(function (_super) {
		__extends(ChartSeries, _super);
		function ChartSeries() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(ChartSeries.prototype, "_className", {
			get: function () {
				return "ChartSeries";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartSeries.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["name", "chartType", "hasDataLabels", "filtered", "markerSize", "markerStyle", "showShadow", "markerBackgroundColor", "markerForegroundColor", "smooth", "plotOrder", "gapWidth", "doughnutHoleSize", "axisGroup", "explosion", "firstSliceAngle", "invertIfNegative", "bubbleScale", "secondPlotSize", "splitType", "splitValue", "varyByCategories", "showLeaderLines", "overlap", "gradientStyle", "gradientMinimumType", "gradientMidpointType", "gradientMaximumType", "gradientMinimumValue", "gradientMidpointValue", "gradientMaximumValue", "gradientMinimumColor", "gradientMidpointColor", "gradientMaximumColor", "parentLabelStrategy", "showConnectorLines", "invertColor"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartSeries.prototype, "_scalarPropertyUpdateable", {
			get: function () {
				return [true, true, true, true, true, true, true, true, true, true, true, true, true, true, true, true, true, true, true, true, true, true, true, true, true, true, true, true, true, true, true, true, true, true, true, true, true];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartSeries.prototype, "_navigationPropertyNames", {
			get: function () {
				return ["points", "format", "trendlines", "xerrorBars", "yerrorBars", "dataLabels", "binOptions", "mapOptions", "boxwhiskerOptions"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartSeries.prototype, "binOptions", {
			get: function () {
				_throwIfApiNotSupported("ChartSeries.binOptions", _defaultApiSetName, "1.9", _hostName);
				return _createPropertyObject(ExcelOp.ChartBinOptions, this, "BinOptions", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartSeries.prototype, "boxwhiskerOptions", {
			get: function () {
				_throwIfApiNotSupported("ChartSeries.boxwhiskerOptions", _defaultApiSetName, "1.9", _hostName);
				return _createPropertyObject(ExcelOp.ChartBoxwhiskerOptions, this, "BoxwhiskerOptions", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartSeries.prototype, "dataLabels", {
			get: function () {
				_throwIfApiNotSupported("ChartSeries.dataLabels", _defaultApiSetName, "1.8", _hostName);
				return _createPropertyObject(ExcelOp.ChartDataLabels, this, "DataLabels", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartSeries.prototype, "format", {
			get: function () {
				return _createPropertyObject(ExcelOp.ChartSeriesFormat, this, "Format", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartSeries.prototype, "mapOptions", {
			get: function () {
				_throwIfApiNotSupported("ChartSeries.mapOptions", _defaultApiSetName, "1.9", _hostName);
				return _createPropertyObject(ExcelOp.ChartMapOptions, this, "MapOptions", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartSeries.prototype, "points", {
			get: function () {
				return _createPropertyObject(ExcelOp.ChartPointsCollection, this, "Points", true, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartSeries.prototype, "trendlines", {
			get: function () {
				_throwIfApiNotSupported("ChartSeries.trendlines", _defaultApiSetName, "1.7", _hostName);
				return _createPropertyObject(ExcelOp.ChartTrendlineCollection, this, "Trendlines", true, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartSeries.prototype, "xerrorBars", {
			get: function () {
				_throwIfApiNotSupported("ChartSeries.xerrorBars", _defaultApiSetName, "1.9", _hostName);
				return _createPropertyObject(ExcelOp.ChartErrorBars, this, "XErrorBars", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartSeries.prototype, "yerrorBars", {
			get: function () {
				_throwIfApiNotSupported("ChartSeries.yerrorBars", _defaultApiSetName, "1.9", _hostName);
				return _createPropertyObject(ExcelOp.ChartErrorBars, this, "YErrorBars", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		ChartSeries.prototype.update=function (properties) {
			return _invokeRecursiveUpdate(this, properties);
		};
		ChartSeries.prototype["delete"]=function () {
			_throwIfApiNotSupported("ChartSeries.delete", _defaultApiSetName, "1.7", _hostName);
			return _invokeMethod(this, "Delete", 0, [], 0);
		};
		ChartSeries.prototype.setBubbleSizes=function (sourceData) {
			_throwIfApiNotSupported("ChartSeries.setBubbleSizes", _defaultApiSetName, "1.7", _hostName);
			return _invokeMethod(this, "SetBubbleSizes", 0, [sourceData], 0);
		};
		ChartSeries.prototype.setValues=function (sourceData) {
			_throwIfApiNotSupported("ChartSeries.setValues", _defaultApiSetName, "1.7", _hostName);
			return _invokeMethod(this, "SetValues", 0, [sourceData], 0);
		};
		ChartSeries.prototype.setXAxisValues=function (sourceData) {
			_throwIfApiNotSupported("ChartSeries.setXAxisValues", _defaultApiSetName, "1.7", _hostName);
			return _invokeMethod(this, "SetXAxisValues", 0, [sourceData], 0);
		};
		ChartSeries.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		ChartSeries.prototype.toJSON=function () {
			return {};
		};
		return ChartSeries;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.ChartSeries=ChartSeries;
	var ChartSeriesFormat=(function (_super) {
		__extends(ChartSeriesFormat, _super);
		function ChartSeriesFormat() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(ChartSeriesFormat.prototype, "_className", {
			get: function () {
				return "ChartSeriesFormat";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartSeriesFormat.prototype, "_navigationPropertyNames", {
			get: function () {
				return ["fill", "line"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartSeriesFormat.prototype, "fill", {
			get: function () {
				return _createPropertyObject(ExcelOp.ChartFill, this, "Fill", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartSeriesFormat.prototype, "line", {
			get: function () {
				return _createPropertyObject(ExcelOp.ChartLineFormat, this, "Line", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		ChartSeriesFormat.prototype.toJSON=function () {
			return {};
		};
		return ChartSeriesFormat;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.ChartSeriesFormat=ChartSeriesFormat;
	var ChartPointsCollection=(function (_super) {
		__extends(ChartPointsCollection, _super);
		function ChartPointsCollection() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(ChartPointsCollection.prototype, "_className", {
			get: function () {
				return "ChartPointsCollection";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartPointsCollection.prototype, "_isCollection", {
			get: function () {
				return true;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartPointsCollection.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["count"];
			},
			enumerable: true,
			configurable: true
		});
		ChartPointsCollection.prototype.getItemAt=function (index) {
			return _createMethodObject(ExcelOp.ChartPoint, this, "GetItemAt", 1, [index], false, false, null, 4);
		};
		ChartPointsCollection.prototype.getCount=function () {
			_throwIfApiNotSupported("ChartPointsCollection.getCount", _defaultApiSetName, "1.4", _hostName);
			return _invokeMethod(this, "GetCount", 1, [], 4, 0);
		};
		ChartPointsCollection.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		ChartPointsCollection.prototype.toJSON=function () {
			return {};
		};
		return ChartPointsCollection;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.ChartPointsCollection=ChartPointsCollection;
	var ChartPoint=(function (_super) {
		__extends(ChartPoint, _super);
		function ChartPoint() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(ChartPoint.prototype, "_className", {
			get: function () {
				return "ChartPoint";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartPoint.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["value", "hasDataLabel", "markerStyle", "markerSize", "markerBackgroundColor", "markerForegroundColor"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartPoint.prototype, "_scalarPropertyUpdateable", {
			get: function () {
				return [false, true, true, true, true, true];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartPoint.prototype, "_navigationPropertyNames", {
			get: function () {
				return ["format", "dataLabel"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartPoint.prototype, "dataLabel", {
			get: function () {
				_throwIfApiNotSupported("ChartPoint.dataLabel", _defaultApiSetName, "1.7", _hostName);
				return _createPropertyObject(ExcelOp.ChartDataLabel, this, "DataLabel", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartPoint.prototype, "format", {
			get: function () {
				return _createPropertyObject(ExcelOp.ChartPointFormat, this, "Format", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		ChartPoint.prototype.update=function (properties) {
			return _invokeRecursiveUpdate(this, properties);
		};
		ChartPoint.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		ChartPoint.prototype.toJSON=function () {
			return {};
		};
		return ChartPoint;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.ChartPoint=ChartPoint;
	var ChartPointFormat=(function (_super) {
		__extends(ChartPointFormat, _super);
		function ChartPointFormat() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(ChartPointFormat.prototype, "_className", {
			get: function () {
				return "ChartPointFormat";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartPointFormat.prototype, "_navigationPropertyNames", {
			get: function () {
				return ["fill", "border"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartPointFormat.prototype, "border", {
			get: function () {
				_throwIfApiNotSupported("ChartPointFormat.border", _defaultApiSetName, "1.7", _hostName);
				return _createPropertyObject(ExcelOp.ChartBorder, this, "Border", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartPointFormat.prototype, "fill", {
			get: function () {
				return _createPropertyObject(ExcelOp.ChartFill, this, "Fill", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		ChartPointFormat.prototype.toJSON=function () {
			return {};
		};
		return ChartPointFormat;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.ChartPointFormat=ChartPointFormat;
	var ChartAxes=(function (_super) {
		__extends(ChartAxes, _super);
		function ChartAxes() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(ChartAxes.prototype, "_className", {
			get: function () {
				return "ChartAxes";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartAxes.prototype, "_navigationPropertyNames", {
			get: function () {
				return ["categoryAxis", "seriesAxis", "valueAxis"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartAxes.prototype, "categoryAxis", {
			get: function () {
				return _createPropertyObject(ExcelOp.ChartAxis, this, "CategoryAxis", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartAxes.prototype, "seriesAxis", {
			get: function () {
				return _createPropertyObject(ExcelOp.ChartAxis, this, "SeriesAxis", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartAxes.prototype, "valueAxis", {
			get: function () {
				return _createPropertyObject(ExcelOp.ChartAxis, this, "ValueAxis", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		ChartAxes.prototype.getItem=function (type, group) {
			_throwIfApiNotSupported("ChartAxes.getItem", _defaultApiSetName, "1.7", _hostName);
			return _createMethodObject(ExcelOp.ChartAxis, this, "GetItem", 1, [type, group], false, false, null, 4);
		};
		ChartAxes.prototype.toJSON=function () {
			return {};
		};
		return ChartAxes;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.ChartAxes=ChartAxes;
	var ChartAxis=(function (_super) {
		__extends(ChartAxis, _super);
		function ChartAxis() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(ChartAxis.prototype, "_className", {
			get: function () {
				return "ChartAxis";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartAxis.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["majorUnit", "maximum", "minimum", "minorUnit", "displayUnit", "showDisplayUnitLabel", "customDisplayUnit", "type", "minorTimeUnitScale", "majorTimeUnitScale", "baseTimeUnit", "categoryType", "axisGroup", "scaleType", "logBase", "left", "top", "height", "width", "reversePlotOrder", "crosses", "crossesAt", "visible", "isBetweenCategories", "majorTickMark", "minorTickMark", "tickMarkSpacing", "tickLabelPosition", "tickLabelSpacing", "alignment", "multiLevel", "numberFormat", "linkNumberFormat", "offset", "textOrientation", "position", "positionAt"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartAxis.prototype, "_scalarPropertyUpdateable", {
			get: function () {
				return [true, true, true, true, true, true, false, false, true, true, true, true, false, true, true, false, false, false, false, true, true, false, true, true, true, true, true, true, true, true, true, true, true, true, true, true, false];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartAxis.prototype, "_navigationPropertyNames", {
			get: function () {
				return ["majorGridlines", "minorGridlines", "title", "format"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartAxis.prototype, "format", {
			get: function () {
				return _createPropertyObject(ExcelOp.ChartAxisFormat, this, "Format", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartAxis.prototype, "majorGridlines", {
			get: function () {
				return _createPropertyObject(ExcelOp.ChartGridlines, this, "MajorGridlines", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartAxis.prototype, "minorGridlines", {
			get: function () {
				return _createPropertyObject(ExcelOp.ChartGridlines, this, "MinorGridlines", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartAxis.prototype, "title", {
			get: function () {
				return _createPropertyObject(ExcelOp.ChartAxisTitle, this, "Title", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		ChartAxis.prototype.update=function (properties) {
			return _invokeRecursiveUpdate(this, properties);
		};
		ChartAxis.prototype.setCategoryNames=function (sourceData) {
			_throwIfApiNotSupported("ChartAxis.setCategoryNames", _defaultApiSetName, "1.7", _hostName);
			return _invokeMethod(this, "SetCategoryNames", 0, [sourceData], 0);
		};
		ChartAxis.prototype.setCrossesAt=function (value) {
			_throwIfApiNotSupported("ChartAxis.setCrossesAt", _defaultApiSetName, "1.7", _hostName);
			return _invokeMethod(this, "SetCrossesAt", 0, [value], 0);
		};
		ChartAxis.prototype.setCustomDisplayUnit=function (value) {
			_throwIfApiNotSupported("ChartAxis.setCustomDisplayUnit", _defaultApiSetName, "1.7", _hostName);
			return _invokeMethod(this, "SetCustomDisplayUnit", 0, [value], 0);
		};
		ChartAxis.prototype.setPositionAt=function (value) {
			_throwIfApiNotSupported("ChartAxis.setPositionAt", _defaultApiSetName, "1.8", _hostName);
			return _invokeMethod(this, "SetPositionAt", 0, [value], 0);
		};
		ChartAxis.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		ChartAxis.prototype.toJSON=function () {
			return {};
		};
		return ChartAxis;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.ChartAxis=ChartAxis;
	var ChartAxisFormat=(function (_super) {
		__extends(ChartAxisFormat, _super);
		function ChartAxisFormat() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(ChartAxisFormat.prototype, "_className", {
			get: function () {
				return "ChartAxisFormat";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartAxisFormat.prototype, "_navigationPropertyNames", {
			get: function () {
				return ["font", "line", "fill"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartAxisFormat.prototype, "fill", {
			get: function () {
				_throwIfApiNotSupported("ChartAxisFormat.fill", _defaultApiSetName, "1.8", _hostName);
				return _createPropertyObject(ExcelOp.ChartFill, this, "Fill", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartAxisFormat.prototype, "font", {
			get: function () {
				return _createPropertyObject(ExcelOp.ChartFont, this, "Font", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartAxisFormat.prototype, "line", {
			get: function () {
				return _createPropertyObject(ExcelOp.ChartLineFormat, this, "Line", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		ChartAxisFormat.prototype.toJSON=function () {
			return {};
		};
		return ChartAxisFormat;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.ChartAxisFormat=ChartAxisFormat;
	var ChartAxisTitle=(function (_super) {
		__extends(ChartAxisTitle, _super);
		function ChartAxisTitle() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(ChartAxisTitle.prototype, "_className", {
			get: function () {
				return "ChartAxisTitle";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartAxisTitle.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["text", "visible"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartAxisTitle.prototype, "_scalarPropertyUpdateable", {
			get: function () {
				return [true, true];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartAxisTitle.prototype, "_navigationPropertyNames", {
			get: function () {
				return ["format"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartAxisTitle.prototype, "format", {
			get: function () {
				return _createPropertyObject(ExcelOp.ChartAxisTitleFormat, this, "Format", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		ChartAxisTitle.prototype.update=function (properties) {
			return _invokeRecursiveUpdate(this, properties);
		};
		ChartAxisTitle.prototype.setFormula=function (formula) {
			_throwIfApiNotSupported("ChartAxisTitle.setFormula", _defaultApiSetName, "1.8", _hostName);
			return _invokeMethod(this, "SetFormula", 0, [formula], 0);
		};
		ChartAxisTitle.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		ChartAxisTitle.prototype.toJSON=function () {
			return {};
		};
		return ChartAxisTitle;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.ChartAxisTitle=ChartAxisTitle;
	var ChartAxisTitleFormat=(function (_super) {
		__extends(ChartAxisTitleFormat, _super);
		function ChartAxisTitleFormat() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(ChartAxisTitleFormat.prototype, "_className", {
			get: function () {
				return "ChartAxisTitleFormat";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartAxisTitleFormat.prototype, "_navigationPropertyNames", {
			get: function () {
				return ["font", "fill", "border"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartAxisTitleFormat.prototype, "border", {
			get: function () {
				_throwIfApiNotSupported("ChartAxisTitleFormat.border", _defaultApiSetName, "1.8", _hostName);
				return _createPropertyObject(ExcelOp.ChartBorder, this, "Border", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartAxisTitleFormat.prototype, "fill", {
			get: function () {
				_throwIfApiNotSupported("ChartAxisTitleFormat.fill", _defaultApiSetName, "1.8", _hostName);
				return _createPropertyObject(ExcelOp.ChartFill, this, "Fill", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartAxisTitleFormat.prototype, "font", {
			get: function () {
				return _createPropertyObject(ExcelOp.ChartFont, this, "Font", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		ChartAxisTitleFormat.prototype.toJSON=function () {
			return {};
		};
		return ChartAxisTitleFormat;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.ChartAxisTitleFormat=ChartAxisTitleFormat;
	var ChartDataLabels=(function (_super) {
		__extends(ChartDataLabels, _super);
		function ChartDataLabels() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(ChartDataLabels.prototype, "_className", {
			get: function () {
				return "ChartDataLabels";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartDataLabels.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["position", "showValue", "showSeriesName", "showCategoryName", "showLegendKey", "showPercentage", "showBubbleSize", "separator", "numberFormat", "linkNumberFormat", "textOrientation", "autoText", "horizontalAlignment", "verticalAlignment"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartDataLabels.prototype, "_scalarPropertyUpdateable", {
			get: function () {
				return [true, true, true, true, true, true, true, true, true, true, true, true, true, true];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartDataLabels.prototype, "_navigationPropertyNames", {
			get: function () {
				return ["format"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartDataLabels.prototype, "format", {
			get: function () {
				return _createPropertyObject(ExcelOp.ChartDataLabelFormat, this, "Format", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		ChartDataLabels.prototype.update=function (properties) {
			return _invokeRecursiveUpdate(this, properties);
		};
		ChartDataLabels.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		ChartDataLabels.prototype.toJSON=function () {
			return {};
		};
		return ChartDataLabels;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.ChartDataLabels=ChartDataLabels;
	var ChartDataLabel=(function (_super) {
		__extends(ChartDataLabel, _super);
		function ChartDataLabel() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(ChartDataLabel.prototype, "_className", {
			get: function () {
				return "ChartDataLabel";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartDataLabel.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["position", "showValue", "showSeriesName", "showCategoryName", "showLegendKey", "showPercentage", "showBubbleSize", "separator", "top", "left", "width", "height", "formula", "textOrientation", "horizontalAlignment", "verticalAlignment", "text", "autoText", "numberFormat", "linkNumberFormat"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartDataLabel.prototype, "_scalarPropertyUpdateable", {
			get: function () {
				return [true, true, true, true, true, true, true, true, true, true, false, false, true, true, true, true, true, true, true, true];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartDataLabel.prototype, "_navigationPropertyNames", {
			get: function () {
				return ["format"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartDataLabel.prototype, "format", {
			get: function () {
				_throwIfApiNotSupported("ChartDataLabel.format", _defaultApiSetName, "1.8", _hostName);
				return _createPropertyObject(ExcelOp.ChartDataLabelFormat, this, "Format", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		ChartDataLabel.prototype.update=function (properties) {
			return _invokeRecursiveUpdate(this, properties);
		};
		ChartDataLabel.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		ChartDataLabel.prototype.toJSON=function () {
			return {};
		};
		return ChartDataLabel;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.ChartDataLabel=ChartDataLabel;
	var ChartDataLabelFormat=(function (_super) {
		__extends(ChartDataLabelFormat, _super);
		function ChartDataLabelFormat() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(ChartDataLabelFormat.prototype, "_className", {
			get: function () {
				return "ChartDataLabelFormat";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartDataLabelFormat.prototype, "_navigationPropertyNames", {
			get: function () {
				return ["font", "fill", "border"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartDataLabelFormat.prototype, "border", {
			get: function () {
				_throwIfApiNotSupported("ChartDataLabelFormat.border", _defaultApiSetName, "1.8", _hostName);
				return _createPropertyObject(ExcelOp.ChartBorder, this, "Border", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartDataLabelFormat.prototype, "fill", {
			get: function () {
				return _createPropertyObject(ExcelOp.ChartFill, this, "Fill", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartDataLabelFormat.prototype, "font", {
			get: function () {
				return _createPropertyObject(ExcelOp.ChartFont, this, "Font", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		ChartDataLabelFormat.prototype.toJSON=function () {
			return {};
		};
		return ChartDataLabelFormat;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.ChartDataLabelFormat=ChartDataLabelFormat;
	var ChartErrorBars=(function (_super) {
		__extends(ChartErrorBars, _super);
		function ChartErrorBars() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(ChartErrorBars.prototype, "_className", {
			get: function () {
				return "ChartErrorBars";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartErrorBars.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["endStyleCap", "include", "type", "visible"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartErrorBars.prototype, "_scalarPropertyUpdateable", {
			get: function () {
				return [true, true, true, true];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartErrorBars.prototype, "_navigationPropertyNames", {
			get: function () {
				return ["format"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartErrorBars.prototype, "format", {
			get: function () {
				return _createPropertyObject(ExcelOp.ChartErrorBarsFormat, this, "Format", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		ChartErrorBars.prototype.update=function (properties) {
			return _invokeRecursiveUpdate(this, properties);
		};
		ChartErrorBars.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		ChartErrorBars.prototype.toJSON=function () {
			return {};
		};
		return ChartErrorBars;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.ChartErrorBars=ChartErrorBars;
	var ChartErrorBarsFormat=(function (_super) {
		__extends(ChartErrorBarsFormat, _super);
		function ChartErrorBarsFormat() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(ChartErrorBarsFormat.prototype, "_className", {
			get: function () {
				return "ChartErrorBarsFormat";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartErrorBarsFormat.prototype, "_navigationPropertyNames", {
			get: function () {
				return ["line"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartErrorBarsFormat.prototype, "line", {
			get: function () {
				return _createPropertyObject(ExcelOp.ChartLineFormat, this, "Line", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		ChartErrorBarsFormat.prototype.toJSON=function () {
			return {};
		};
		return ChartErrorBarsFormat;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.ChartErrorBarsFormat=ChartErrorBarsFormat;
	var ChartGridlines=(function (_super) {
		__extends(ChartGridlines, _super);
		function ChartGridlines() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(ChartGridlines.prototype, "_className", {
			get: function () {
				return "ChartGridlines";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartGridlines.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["visible"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartGridlines.prototype, "_scalarPropertyUpdateable", {
			get: function () {
				return [true];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartGridlines.prototype, "_navigationPropertyNames", {
			get: function () {
				return ["format"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartGridlines.prototype, "format", {
			get: function () {
				return _createPropertyObject(ExcelOp.ChartGridlinesFormat, this, "Format", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		ChartGridlines.prototype.update=function (properties) {
			return _invokeRecursiveUpdate(this, properties);
		};
		ChartGridlines.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		ChartGridlines.prototype.toJSON=function () {
			return {};
		};
		return ChartGridlines;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.ChartGridlines=ChartGridlines;
	var ChartGridlinesFormat=(function (_super) {
		__extends(ChartGridlinesFormat, _super);
		function ChartGridlinesFormat() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(ChartGridlinesFormat.prototype, "_className", {
			get: function () {
				return "ChartGridlinesFormat";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartGridlinesFormat.prototype, "_navigationPropertyNames", {
			get: function () {
				return ["line"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartGridlinesFormat.prototype, "line", {
			get: function () {
				return _createPropertyObject(ExcelOp.ChartLineFormat, this, "Line", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		ChartGridlinesFormat.prototype.toJSON=function () {
			return {};
		};
		return ChartGridlinesFormat;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.ChartGridlinesFormat=ChartGridlinesFormat;
	var ChartLegend=(function (_super) {
		__extends(ChartLegend, _super);
		function ChartLegend() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(ChartLegend.prototype, "_className", {
			get: function () {
				return "ChartLegend";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartLegend.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["visible", "position", "overlay", "left", "top", "width", "height", "showShadow"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartLegend.prototype, "_scalarPropertyUpdateable", {
			get: function () {
				return [true, true, true, true, true, true, true, true];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartLegend.prototype, "_navigationPropertyNames", {
			get: function () {
				return ["format", "legendEntries"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartLegend.prototype, "format", {
			get: function () {
				return _createPropertyObject(ExcelOp.ChartLegendFormat, this, "Format", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartLegend.prototype, "legendEntries", {
			get: function () {
				_throwIfApiNotSupported("ChartLegend.legendEntries", _defaultApiSetName, "1.7", _hostName);
				return _createPropertyObject(ExcelOp.ChartLegendEntryCollection, this, "LegendEntries", true, 4);
			},
			enumerable: true,
			configurable: true
		});
		ChartLegend.prototype.update=function (properties) {
			return _invokeRecursiveUpdate(this, properties);
		};
		ChartLegend.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		ChartLegend.prototype.toJSON=function () {
			return {};
		};
		return ChartLegend;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.ChartLegend=ChartLegend;
	var ChartLegendEntry=(function (_super) {
		__extends(ChartLegendEntry, _super);
		function ChartLegendEntry() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(ChartLegendEntry.prototype, "_className", {
			get: function () {
				return "ChartLegendEntry";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartLegendEntry.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["visible", "left", "top", "width", "height", "index"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartLegendEntry.prototype, "_scalarPropertyUpdateable", {
			get: function () {
				return [true, false, false, false, false, false];
			},
			enumerable: true,
			configurable: true
		});
		ChartLegendEntry.prototype.update=function (properties) {
			return _invokeRecursiveUpdate(this, properties);
		};
		ChartLegendEntry.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		ChartLegendEntry.prototype.toJSON=function () {
			return {};
		};
		return ChartLegendEntry;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.ChartLegendEntry=ChartLegendEntry;
	var ChartLegendEntryCollection=(function (_super) {
		__extends(ChartLegendEntryCollection, _super);
		function ChartLegendEntryCollection() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(ChartLegendEntryCollection.prototype, "_className", {
			get: function () {
				return "ChartLegendEntryCollection";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartLegendEntryCollection.prototype, "_isCollection", {
			get: function () {
				return true;
			},
			enumerable: true,
			configurable: true
		});
		ChartLegendEntryCollection.prototype.getItemAt=function (index) {
			return _createMethodObject(ExcelOp.ChartLegendEntry, this, "GetItemAt", 1, [index], false, false, null, 4);
		};
		ChartLegendEntryCollection.prototype.getCount=function () {
			return _invokeMethod(this, "GetCount", 1, [], 4, 0);
		};
		ChartLegendEntryCollection.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		ChartLegendEntryCollection.prototype.toJSON=function () {
			return {};
		};
		return ChartLegendEntryCollection;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.ChartLegendEntryCollection=ChartLegendEntryCollection;
	var ChartLegendFormat=(function (_super) {
		__extends(ChartLegendFormat, _super);
		function ChartLegendFormat() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(ChartLegendFormat.prototype, "_className", {
			get: function () {
				return "ChartLegendFormat";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartLegendFormat.prototype, "_navigationPropertyNames", {
			get: function () {
				return ["font", "fill", "border"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartLegendFormat.prototype, "border", {
			get: function () {
				_throwIfApiNotSupported("ChartLegendFormat.border", _defaultApiSetName, "1.8", _hostName);
				return _createPropertyObject(ExcelOp.ChartBorder, this, "Border", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartLegendFormat.prototype, "fill", {
			get: function () {
				return _createPropertyObject(ExcelOp.ChartFill, this, "Fill", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartLegendFormat.prototype, "font", {
			get: function () {
				return _createPropertyObject(ExcelOp.ChartFont, this, "Font", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		ChartLegendFormat.prototype.toJSON=function () {
			return {};
		};
		return ChartLegendFormat;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.ChartLegendFormat=ChartLegendFormat;
	var ChartMapOptions=(function (_super) {
		__extends(ChartMapOptions, _super);
		function ChartMapOptions() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(ChartMapOptions.prototype, "_className", {
			get: function () {
				return "ChartMapOptions";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartMapOptions.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["level", "labelStrategy", "projectionType"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartMapOptions.prototype, "_scalarPropertyUpdateable", {
			get: function () {
				return [true, true, true];
			},
			enumerable: true,
			configurable: true
		});
		ChartMapOptions.prototype.update=function (properties) {
			return _invokeRecursiveUpdate(this, properties);
		};
		ChartMapOptions.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		ChartMapOptions.prototype.toJSON=function () {
			return {};
		};
		return ChartMapOptions;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.ChartMapOptions=ChartMapOptions;
	var ChartTitle=(function (_super) {
		__extends(ChartTitle, _super);
		function ChartTitle() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(ChartTitle.prototype, "_className", {
			get: function () {
				return "ChartTitle";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartTitle.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["visible", "text", "overlay", "horizontalAlignment", "top", "left", "width", "height", "verticalAlignment", "textOrientation", "position", "showShadow"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartTitle.prototype, "_scalarPropertyUpdateable", {
			get: function () {
				return [true, true, true, true, true, true, false, false, true, true, true, true];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartTitle.prototype, "_navigationPropertyNames", {
			get: function () {
				return ["format"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartTitle.prototype, "format", {
			get: function () {
				return _createPropertyObject(ExcelOp.ChartTitleFormat, this, "Format", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		ChartTitle.prototype.getSubstring=function (start, length) {
			_throwIfApiNotSupported("ChartTitle.getSubstring", _defaultApiSetName, "1.7", _hostName);
			return _createMethodObject(ExcelOp.ChartFormatString, this, "GetSubstring", 1, [start, length], false, false, null, 4);
		};
		ChartTitle.prototype.update=function (properties) {
			return _invokeRecursiveUpdate(this, properties);
		};
		ChartTitle.prototype.setFormula=function (formula) {
			_throwIfApiNotSupported("ChartTitle.setFormula", _defaultApiSetName, "1.7", _hostName);
			return _invokeMethod(this, "SetFormula", 0, [formula], 0);
		};
		ChartTitle.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		ChartTitle.prototype.toJSON=function () {
			return {};
		};
		return ChartTitle;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.ChartTitle=ChartTitle;
	var ChartFormatString=(function (_super) {
		__extends(ChartFormatString, _super);
		function ChartFormatString() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(ChartFormatString.prototype, "_className", {
			get: function () {
				return "ChartFormatString";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartFormatString.prototype, "_navigationPropertyNames", {
			get: function () {
				return ["font"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartFormatString.prototype, "font", {
			get: function () {
				return _createPropertyObject(ExcelOp.ChartFont, this, "Font", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		ChartFormatString.prototype.toJSON=function () {
			return {};
		};
		return ChartFormatString;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.ChartFormatString=ChartFormatString;
	var ChartTitleFormat=(function (_super) {
		__extends(ChartTitleFormat, _super);
		function ChartTitleFormat() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(ChartTitleFormat.prototype, "_className", {
			get: function () {
				return "ChartTitleFormat";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartTitleFormat.prototype, "_navigationPropertyNames", {
			get: function () {
				return ["font", "fill", "border"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartTitleFormat.prototype, "border", {
			get: function () {
				_throwIfApiNotSupported("ChartTitleFormat.border", _defaultApiSetName, "1.7", _hostName);
				return _createPropertyObject(ExcelOp.ChartBorder, this, "Border", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartTitleFormat.prototype, "fill", {
			get: function () {
				return _createPropertyObject(ExcelOp.ChartFill, this, "Fill", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartTitleFormat.prototype, "font", {
			get: function () {
				return _createPropertyObject(ExcelOp.ChartFont, this, "Font", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		ChartTitleFormat.prototype.toJSON=function () {
			return {};
		};
		return ChartTitleFormat;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.ChartTitleFormat=ChartTitleFormat;
	var ChartFill=(function (_super) {
		__extends(ChartFill, _super);
		function ChartFill() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(ChartFill.prototype, "_className", {
			get: function () {
				return "ChartFill";
			},
			enumerable: true,
			configurable: true
		});
		ChartFill.prototype.clear=function () {
			return _invokeMethod(this, "Clear", 0, [], 0);
		};
		ChartFill.prototype.setSolidColor=function (color) {
			return _invokeMethod(this, "SetSolidColor", 0, [color], 0);
		};
		ChartFill.prototype.toJSON=function () {
			return {};
		};
		return ChartFill;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.ChartFill=ChartFill;
	var ChartBorder=(function (_super) {
		__extends(ChartBorder, _super);
		function ChartBorder() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(ChartBorder.prototype, "_className", {
			get: function () {
				return "ChartBorder";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartBorder.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["color", "lineStyle", "weight"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartBorder.prototype, "_scalarPropertyUpdateable", {
			get: function () {
				return [true, true, true];
			},
			enumerable: true,
			configurable: true
		});
		ChartBorder.prototype.update=function (properties) {
			return _invokeRecursiveUpdate(this, properties);
		};
		ChartBorder.prototype.clear=function () {
			_throwIfApiNotSupported("ChartBorder.clear", _defaultApiSetName, "1.8", _hostName);
			return _invokeMethod(this, "Clear", 0, [], 0);
		};
		ChartBorder.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		ChartBorder.prototype.toJSON=function () {
			return {};
		};
		return ChartBorder;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.ChartBorder=ChartBorder;
	var ChartBinOptions=(function (_super) {
		__extends(ChartBinOptions, _super);
		function ChartBinOptions() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(ChartBinOptions.prototype, "_className", {
			get: function () {
				return "ChartBinOptions";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartBinOptions.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["type", "width", "count", "allowOverflow", "allowUnderflow", "overflowValue", "underflowValue"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartBinOptions.prototype, "_scalarPropertyUpdateable", {
			get: function () {
				return [true, true, true, true, true, true, true];
			},
			enumerable: true,
			configurable: true
		});
		ChartBinOptions.prototype.update=function (properties) {
			return _invokeRecursiveUpdate(this, properties);
		};
		ChartBinOptions.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		ChartBinOptions.prototype.toJSON=function () {
			return {};
		};
		return ChartBinOptions;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.ChartBinOptions=ChartBinOptions;
	var ChartBoxwhiskerOptions=(function (_super) {
		__extends(ChartBoxwhiskerOptions, _super);
		function ChartBoxwhiskerOptions() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(ChartBoxwhiskerOptions.prototype, "_className", {
			get: function () {
				return "ChartBoxwhiskerOptions";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartBoxwhiskerOptions.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["showInnerPoints", "showOutlierPoints", "showMeanMarker", "showMeanLine", "quartileCalculation"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartBoxwhiskerOptions.prototype, "_scalarPropertyUpdateable", {
			get: function () {
				return [true, true, true, true, true];
			},
			enumerable: true,
			configurable: true
		});
		ChartBoxwhiskerOptions.prototype.update=function (properties) {
			return _invokeRecursiveUpdate(this, properties);
		};
		ChartBoxwhiskerOptions.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		ChartBoxwhiskerOptions.prototype.toJSON=function () {
			return {};
		};
		return ChartBoxwhiskerOptions;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.ChartBoxwhiskerOptions=ChartBoxwhiskerOptions;
	var ChartLineFormat=(function (_super) {
		__extends(ChartLineFormat, _super);
		function ChartLineFormat() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(ChartLineFormat.prototype, "_className", {
			get: function () {
				return "ChartLineFormat";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartLineFormat.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["color", "lineStyle", "weight"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartLineFormat.prototype, "_scalarPropertyUpdateable", {
			get: function () {
				return [true, true, true];
			},
			enumerable: true,
			configurable: true
		});
		ChartLineFormat.prototype.update=function (properties) {
			return _invokeRecursiveUpdate(this, properties);
		};
		ChartLineFormat.prototype.clear=function () {
			return _invokeMethod(this, "Clear", 0, [], 0);
		};
		ChartLineFormat.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		ChartLineFormat.prototype.toJSON=function () {
			return {};
		};
		return ChartLineFormat;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.ChartLineFormat=ChartLineFormat;
	var ChartFont=(function (_super) {
		__extends(ChartFont, _super);
		function ChartFont() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(ChartFont.prototype, "_className", {
			get: function () {
				return "ChartFont";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartFont.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["bold", "color", "italic", "name", "size", "underline"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartFont.prototype, "_scalarPropertyUpdateable", {
			get: function () {
				return [true, true, true, true, true, true];
			},
			enumerable: true,
			configurable: true
		});
		ChartFont.prototype.update=function (properties) {
			return _invokeRecursiveUpdate(this, properties);
		};
		ChartFont.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		ChartFont.prototype.toJSON=function () {
			return {};
		};
		return ChartFont;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.ChartFont=ChartFont;
	var ChartTrendline=(function (_super) {
		__extends(ChartTrendline, _super);
		function ChartTrendline() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(ChartTrendline.prototype, "_className", {
			get: function () {
				return "ChartTrendline";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartTrendline.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["type", "polynomialOrder", "movingAveragePeriod", "_Id", "showEquation", "showRSquared", "forwardPeriod", "backwardPeriod", "name", "intercept"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartTrendline.prototype, "_scalarPropertyUpdateable", {
			get: function () {
				return [true, true, true, false, true, true, true, true, true, true];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartTrendline.prototype, "_navigationPropertyNames", {
			get: function () {
				return ["format", "label"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartTrendline.prototype, "format", {
			get: function () {
				return _createPropertyObject(ExcelOp.ChartTrendlineFormat, this, "Format", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartTrendline.prototype, "label", {
			get: function () {
				_throwIfApiNotSupported("ChartTrendline.label", _defaultApiSetName, "1.8", _hostName);
				return _createPropertyObject(ExcelOp.ChartTrendlineLabel, this, "Label", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		ChartTrendline.prototype.update=function (properties) {
			return _invokeRecursiveUpdate(this, properties);
		};
		ChartTrendline.prototype["delete"]=function () {
			return _invokeMethod(this, "Delete", 0, [], 0);
		};
		ChartTrendline.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		ChartTrendline.prototype.toJSON=function () {
			return {};
		};
		return ChartTrendline;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.ChartTrendline=ChartTrendline;
	var ChartTrendlineCollection=(function (_super) {
		__extends(ChartTrendlineCollection, _super);
		function ChartTrendlineCollection() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(ChartTrendlineCollection.prototype, "_className", {
			get: function () {
				return "ChartTrendlineCollection";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartTrendlineCollection.prototype, "_isCollection", {
			get: function () {
				return true;
			},
			enumerable: true,
			configurable: true
		});
		ChartTrendlineCollection.prototype.add=function (type) {
			return _createAndInstantiateMethodObject(ExcelOp.ChartTrendline, this, "Add", 0, [type], false, true, null, 0);
		};
		ChartTrendlineCollection.prototype.getItem=function (index) {
			return _createIndexerObject(ExcelOp.ChartTrendline, this, [index]);
		};
		ChartTrendlineCollection.prototype.getCount=function () {
			return _invokeMethod(this, "GetCount", 1, [], 4, 0);
		};
		ChartTrendlineCollection.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		ChartTrendlineCollection.prototype.toJSON=function () {
			return {};
		};
		return ChartTrendlineCollection;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.ChartTrendlineCollection=ChartTrendlineCollection;
	var ChartTrendlineFormat=(function (_super) {
		__extends(ChartTrendlineFormat, _super);
		function ChartTrendlineFormat() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(ChartTrendlineFormat.prototype, "_className", {
			get: function () {
				return "ChartTrendlineFormat";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartTrendlineFormat.prototype, "_navigationPropertyNames", {
			get: function () {
				return ["line"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartTrendlineFormat.prototype, "line", {
			get: function () {
				return _createPropertyObject(ExcelOp.ChartLineFormat, this, "Line", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		ChartTrendlineFormat.prototype.toJSON=function () {
			return {};
		};
		return ChartTrendlineFormat;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.ChartTrendlineFormat=ChartTrendlineFormat;
	var ChartTrendlineLabel=(function (_super) {
		__extends(ChartTrendlineLabel, _super);
		function ChartTrendlineLabel() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(ChartTrendlineLabel.prototype, "_className", {
			get: function () {
				return "ChartTrendlineLabel";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartTrendlineLabel.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["top", "left", "width", "height", "formula", "textOrientation", "horizontalAlignment", "verticalAlignment", "text", "autoText", "numberFormat", "linkNumberFormat"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartTrendlineLabel.prototype, "_scalarPropertyUpdateable", {
			get: function () {
				return [true, true, false, false, true, true, true, true, true, true, true, true];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartTrendlineLabel.prototype, "_navigationPropertyNames", {
			get: function () {
				return ["format"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartTrendlineLabel.prototype, "format", {
			get: function () {
				return _createPropertyObject(ExcelOp.ChartTrendlineLabelFormat, this, "Format", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		ChartTrendlineLabel.prototype.update=function (properties) {
			return _invokeRecursiveUpdate(this, properties);
		};
		ChartTrendlineLabel.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		ChartTrendlineLabel.prototype.toJSON=function () {
			return {};
		};
		return ChartTrendlineLabel;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.ChartTrendlineLabel=ChartTrendlineLabel;
	var ChartTrendlineLabelFormat=(function (_super) {
		__extends(ChartTrendlineLabelFormat, _super);
		function ChartTrendlineLabelFormat() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(ChartTrendlineLabelFormat.prototype, "_className", {
			get: function () {
				return "ChartTrendlineLabelFormat";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartTrendlineLabelFormat.prototype, "_navigationPropertyNames", {
			get: function () {
				return ["fill", "border", "font"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartTrendlineLabelFormat.prototype, "border", {
			get: function () {
				return _createPropertyObject(ExcelOp.ChartBorder, this, "Border", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartTrendlineLabelFormat.prototype, "fill", {
			get: function () {
				return _createPropertyObject(ExcelOp.ChartFill, this, "Fill", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartTrendlineLabelFormat.prototype, "font", {
			get: function () {
				return _createPropertyObject(ExcelOp.ChartFont, this, "Font", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		ChartTrendlineLabelFormat.prototype.toJSON=function () {
			return {};
		};
		return ChartTrendlineLabelFormat;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.ChartTrendlineLabelFormat=ChartTrendlineLabelFormat;
	var ChartPlotArea=(function (_super) {
		__extends(ChartPlotArea, _super);
		function ChartPlotArea() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(ChartPlotArea.prototype, "_className", {
			get: function () {
				return "ChartPlotArea";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartPlotArea.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["left", "top", "width", "height", "insideLeft", "insideTop", "insideWidth", "insideHeight", "position"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartPlotArea.prototype, "_scalarPropertyUpdateable", {
			get: function () {
				return [true, true, true, true, true, true, true, true, true];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartPlotArea.prototype, "_navigationPropertyNames", {
			get: function () {
				return ["format"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartPlotArea.prototype, "format", {
			get: function () {
				return _createPropertyObject(ExcelOp.ChartPlotAreaFormat, this, "Format", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		ChartPlotArea.prototype.update=function (properties) {
			return _invokeRecursiveUpdate(this, properties);
		};
		ChartPlotArea.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		ChartPlotArea.prototype.toJSON=function () {
			return {};
		};
		return ChartPlotArea;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.ChartPlotArea=ChartPlotArea;
	var ChartPlotAreaFormat=(function (_super) {
		__extends(ChartPlotAreaFormat, _super);
		function ChartPlotAreaFormat() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(ChartPlotAreaFormat.prototype, "_className", {
			get: function () {
				return "ChartPlotAreaFormat";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartPlotAreaFormat.prototype, "_navigationPropertyNames", {
			get: function () {
				return ["border", "fill"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartPlotAreaFormat.prototype, "border", {
			get: function () {
				return _createPropertyObject(ExcelOp.ChartBorder, this, "Border", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartPlotAreaFormat.prototype, "fill", {
			get: function () {
				return _createPropertyObject(ExcelOp.ChartFill, this, "Fill", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		ChartPlotAreaFormat.prototype.toJSON=function () {
			return {};
		};
		return ChartPlotAreaFormat;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.ChartPlotAreaFormat=ChartPlotAreaFormat;
	var VisualCollection=(function (_super) {
		__extends(VisualCollection, _super);
		function VisualCollection() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(VisualCollection.prototype, "_className", {
			get: function () {
				return "VisualCollection";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(VisualCollection.prototype, "_isCollection", {
			get: function () {
				return true;
			},
			enumerable: true,
			configurable: true
		});
		VisualCollection.prototype.add=function (visualDefinitionGuid, dataSourceType, dataSourceContent) {
			return _createAndInstantiateMethodObject(ExcelOp.Visual, this, "Add", 0, [visualDefinitionGuid, dataSourceType, dataSourceContent], false, true, null, 2);
		};
		VisualCollection.prototype.getSelectedOrNullObject=function () {
			return _createMethodObject(ExcelOp.Visual, this, "GetSelectedOrNullObject", 1, [], false, false, null, 4);
		};
		VisualCollection.prototype._GetItem=function (id) {
			return _createIndexerObject(ExcelOp.Visual, this, [id]);
		};
		VisualCollection.prototype.bootstrapAgaveVisual=function () {
			return _invokeMethod(this, "BootstrapAgaveVisual", 0, [], 2);
		};
		VisualCollection.prototype.getCount=function () {
			return _invokeMethod(this, "GetCount", 1, [], 4, 0);
		};
		VisualCollection.prototype.getDefinitions=function () {
			return _invokeMethod(this, "GetDefinitions", 1, [], 4, 0);
		};
		VisualCollection.prototype.getPreview=function (visualDefinitionGuid, width, height, dpi) {
			return _invokeMethod(this, "GetPreview", 1, [visualDefinitionGuid, width, height, dpi], 4, 0);
		};
		VisualCollection.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		VisualCollection.prototype.toJSON=function () {
			return {};
		};
		return VisualCollection;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.VisualCollection=VisualCollection;
	var Visual=(function (_super) {
		__extends(Visual, _super);
		function Visual() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(Visual.prototype, "_className", {
			get: function () {
				return "Visual";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Visual.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["id"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Visual.prototype, "_navigationPropertyNames", {
			get: function () {
				return ["properties"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Visual.prototype, "properties", {
			get: function () {
				return _createPropertyObject(ExcelOp.VisualPropertyCollection, this, "Properties", true, 4);
			},
			enumerable: true,
			configurable: true
		});
		Visual.prototype.getChildProperties=function (parentPropId) {
			return _createMethodObject(ExcelOp.VisualPropertyCollection, this, "GetChildProperties", 1, [parentPropId], true, false, null, 4);
		};
		Visual.prototype.getDataControllerClient=function () {
			return _createMethodObject(ExcelOp.DataControllerClient, this, "GetDataControllerClient", 1, [], false, false, null, 4);
		};
		Visual.prototype.getElementChildProperties=function (elementId, index) {
			return _createMethodObject(ExcelOp.VisualPropertyCollection, this, "GetElementChildProperties", 1, [elementId, index], true, false, null, 4);
		};
		Visual.prototype.changeDataSource=function (dataSourceType, dataSourceContent) {
			return _invokeMethod(this, "ChangeDataSource", 0, [dataSourceType, dataSourceContent], 2);
		};
		Visual.prototype["delete"]=function () {
			return _invokeMethod(this, "Delete", 0, [], 2);
		};
		Visual.prototype.getDataSource=function () {
			return _invokeMethod(this, "GetDataSource", 1, [], 4, 0);
		};
		Visual.prototype.getProperty=function (propName) {
			return _invokeMethod(this, "GetProperty", 1, [propName], 4, 0);
		};
		Visual.prototype.setProperty=function (propName, value) {
			return _invokeMethod(this, "SetProperty", 0, [propName, value], 2);
		};
		Visual.prototype.setPropertyToDefault=function (propName) {
			return _invokeMethod(this, "SetPropertyToDefault", 0, [propName], 2);
		};
		Visual.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		Visual.prototype.toJSON=function () {
			return {};
		};
		return Visual;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.Visual=Visual;
	var VisualProperty=(function (_super) {
		__extends(VisualProperty, _super);
		function VisualProperty() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(VisualProperty.prototype, "_className", {
			get: function () {
				return "VisualProperty";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(VisualProperty.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["type", "value", "id", "localizedName", "options", "localizedOptions", "hasDefault", "isDefault", "min", "max", "stepSize", "hideUI"];
			},
			enumerable: true,
			configurable: true
		});
		VisualProperty.prototype.getBoolMetaProperty=function (metaProp) {
			return _invokeMethod(this, "GetBoolMetaProperty", 1, [metaProp], 4, 0);
		};
		VisualProperty.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		VisualProperty.prototype.toJSON=function () {
			return {};
		};
		return VisualProperty;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.VisualProperty=VisualProperty;
	var VisualPropertyCollection=(function (_super) {
		__extends(VisualPropertyCollection, _super);
		function VisualPropertyCollection() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(VisualPropertyCollection.prototype, "_className", {
			get: function () {
				return "VisualPropertyCollection";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(VisualPropertyCollection.prototype, "_isCollection", {
			get: function () {
				return true;
			},
			enumerable: true,
			configurable: true
		});
		VisualPropertyCollection.prototype.getItem=function (index) {
			return _createIndexerObject(ExcelOp.VisualProperty, this, [index]);
		};
		VisualPropertyCollection.prototype.getItemAt=function (index) {
			return _createMethodObject(ExcelOp.VisualProperty, this, "GetItemAt", 1, [index], false, false, null, 4);
		};
		VisualPropertyCollection.prototype.getCount=function () {
			return _invokeMethod(this, "GetCount", 1, [], 4, 0);
		};
		VisualPropertyCollection.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		VisualPropertyCollection.prototype.toJSON=function () {
			return {};
		};
		return VisualPropertyCollection;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.VisualPropertyCollection=VisualPropertyCollection;
	var DataControllerClient=(function (_super) {
		__extends(DataControllerClient, _super);
		function DataControllerClient() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(DataControllerClient.prototype, "_className", {
			get: function () {
				return "DataControllerClient";
			},
			enumerable: true,
			configurable: true
		});
		DataControllerClient.prototype.addField=function (wellId, fieldId, position) {
			return _invokeMethod(this, "AddField", 0, [wellId, fieldId, position], 2);
		};
		DataControllerClient.prototype.getAssociatedFields=function (wellId) {
			return _invokeMethod(this, "GetAssociatedFields", 1, [wellId], 4, 0);
		};
		DataControllerClient.prototype.getAvailableFields=function (wellId) {
			return _invokeMethod(this, "GetAvailableFields", 1, [wellId], 4, 0);
		};
		DataControllerClient.prototype.getWells=function () {
			return _invokeMethod(this, "GetWells", 1, [], 4, 0);
		};
		DataControllerClient.prototype.moveField=function (wellId, fromPosition, toPosition) {
			return _invokeMethod(this, "MoveField", 0, [wellId, fromPosition, toPosition], 2);
		};
		DataControllerClient.prototype.removeField=function (wellId, position) {
			return _invokeMethod(this, "RemoveField", 0, [wellId, position], 2);
		};
		DataControllerClient.prototype.toJSON=function () {
			return {};
		};
		return DataControllerClient;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.DataControllerClient=DataControllerClient;
	var RangeSort=(function (_super) {
		__extends(RangeSort, _super);
		function RangeSort() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(RangeSort.prototype, "_className", {
			get: function () {
				return "RangeSort";
			},
			enumerable: true,
			configurable: true
		});
		RangeSort.prototype.apply=function (fields, matchCase, hasHeaders, orientation, method) {
			return _invokeMethod(this, "Apply", 0, [fields, matchCase, hasHeaders, orientation, method], 0);
		};
		RangeSort.prototype.toJSON=function () {
			return {};
		};
		return RangeSort;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.RangeSort=RangeSort;
	var TableSort=(function (_super) {
		__extends(TableSort, _super);
		function TableSort() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(TableSort.prototype, "_className", {
			get: function () {
				return "TableSort";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableSort.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["matchCase", "method", "fields"];
			},
			enumerable: true,
			configurable: true
		});
		TableSort.prototype.apply=function (fields, matchCase, method) {
			return _invokeMethod(this, "Apply", 0, [fields, matchCase, method], 0);
		};
		TableSort.prototype.clear=function () {
			return _invokeMethod(this, "Clear", 0, [], 0);
		};
		TableSort.prototype.reapply=function () {
			return _invokeMethod(this, "Reapply", 0, [], 0);
		};
		TableSort.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		TableSort.prototype.toJSON=function () {
			return {};
		};
		return TableSort;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.TableSort=TableSort;
	var Filter=(function (_super) {
		__extends(Filter, _super);
		function Filter() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(Filter.prototype, "_className", {
			get: function () {
				return "Filter";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Filter.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["criteria"];
			},
			enumerable: true,
			configurable: true
		});
		Filter.prototype.apply=function (criteria) {
			return _invokeMethod(this, "Apply", 0, [criteria], 0);
		};
		Filter.prototype.applyBottomItemsFilter=function (count) {
			return _invokeMethod(this, "ApplyBottomItemsFilter", 0, [count], 0);
		};
		Filter.prototype.applyBottomPercentFilter=function (percent) {
			return _invokeMethod(this, "ApplyBottomPercentFilter", 0, [percent], 0);
		};
		Filter.prototype.applyCellColorFilter=function (color) {
			return _invokeMethod(this, "ApplyCellColorFilter", 0, [color], 0);
		};
		Filter.prototype.applyCustomFilter=function (criteria1, criteria2, oper) {
			return _invokeMethod(this, "ApplyCustomFilter", 0, [criteria1, criteria2, oper], 0);
		};
		Filter.prototype.applyDynamicFilter=function (criteria) {
			return _invokeMethod(this, "ApplyDynamicFilter", 0, [criteria], 0);
		};
		Filter.prototype.applyFontColorFilter=function (color) {
			return _invokeMethod(this, "ApplyFontColorFilter", 0, [color], 0);
		};
		Filter.prototype.applyIconFilter=function (icon) {
			return _invokeMethod(this, "ApplyIconFilter", 0, [icon], 0);
		};
		Filter.prototype.applyTopItemsFilter=function (count) {
			return _invokeMethod(this, "ApplyTopItemsFilter", 0, [count], 0);
		};
		Filter.prototype.applyTopPercentFilter=function (percent) {
			return _invokeMethod(this, "ApplyTopPercentFilter", 0, [percent], 0);
		};
		Filter.prototype.applyValuesFilter=function (values) {
			return _invokeMethod(this, "ApplyValuesFilter", 0, [values], 0);
		};
		Filter.prototype.clear=function () {
			return _invokeMethod(this, "Clear", 0, [], 0);
		};
		Filter.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		Filter.prototype.toJSON=function () {
			return {};
		};
		return Filter;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.Filter=Filter;
	var AutoFilter=(function (_super) {
		__extends(AutoFilter, _super);
		function AutoFilter() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(AutoFilter.prototype, "_className", {
			get: function () {
				return "AutoFilter";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(AutoFilter.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["enabled", "isDataFiltered", "criteria"];
			},
			enumerable: true,
			configurable: true
		});
		AutoFilter.prototype.getRange=function () {
			return _createMethodObject(ExcelOp.Range, this, "GetRange", 1, [], false, true, null, 4);
		};
		AutoFilter.prototype.getRangeOrNullObject=function () {
			return _createMethodObject(ExcelOp.Range, this, "GetRangeOrNullObject", 1, [], false, true, null, 4);
		};
		AutoFilter.prototype.apply=function (range, columnIndex, criteria) {
			return _invokeMethod(this, "Apply", 0, [range, columnIndex, criteria], 0);
		};
		AutoFilter.prototype.clearCriteria=function () {
			return _invokeMethod(this, "ClearCriteria", 0, [], 0);
		};
		AutoFilter.prototype.reapply=function () {
			return _invokeMethod(this, "Reapply", 0, [], 0);
		};
		AutoFilter.prototype.remove=function () {
			return _invokeMethod(this, "Remove", 0, [], 0);
		};
		AutoFilter.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		AutoFilter.prototype.toJSON=function () {
			return {};
		};
		return AutoFilter;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.AutoFilter=AutoFilter;
	var CustomXmlPartScopedCollection=(function (_super) {
		__extends(CustomXmlPartScopedCollection, _super);
		function CustomXmlPartScopedCollection() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(CustomXmlPartScopedCollection.prototype, "_className", {
			get: function () {
				return "CustomXmlPartScopedCollection";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(CustomXmlPartScopedCollection.prototype, "_isCollection", {
			get: function () {
				return true;
			},
			enumerable: true,
			configurable: true
		});
		CustomXmlPartScopedCollection.prototype.getItem=function (id) {
			return _createIndexerObject(ExcelOp.CustomXmlPart, this, [id]);
		};
		CustomXmlPartScopedCollection.prototype.getItemOrNullObject=function (id) {
			return _createMethodObject(ExcelOp.CustomXmlPart, this, "GetItemOrNullObject", 1, [id], false, false, null, 4);
		};
		CustomXmlPartScopedCollection.prototype.getOnlyItem=function () {
			return _createMethodObject(ExcelOp.CustomXmlPart, this, "GetOnlyItem", 1, [], false, false, null, 4);
		};
		CustomXmlPartScopedCollection.prototype.getOnlyItemOrNullObject=function () {
			return _createMethodObject(ExcelOp.CustomXmlPart, this, "GetOnlyItemOrNullObject", 1, [], false, false, null, 4);
		};
		CustomXmlPartScopedCollection.prototype.getCount=function () {
			return _invokeMethod(this, "GetCount", 1, [], 4, 0);
		};
		CustomXmlPartScopedCollection.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		CustomXmlPartScopedCollection.prototype.toJSON=function () {
			return {};
		};
		return CustomXmlPartScopedCollection;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.CustomXmlPartScopedCollection=CustomXmlPartScopedCollection;
	var CustomXmlPartCollection=(function (_super) {
		__extends(CustomXmlPartCollection, _super);
		function CustomXmlPartCollection() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(CustomXmlPartCollection.prototype, "_className", {
			get: function () {
				return "CustomXmlPartCollection";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(CustomXmlPartCollection.prototype, "_isCollection", {
			get: function () {
				return true;
			},
			enumerable: true,
			configurable: true
		});
		CustomXmlPartCollection.prototype.add=function (xml) {
			return _createAndInstantiateMethodObject(ExcelOp.CustomXmlPart, this, "Add", 0, [xml], false, true, null, 0);
		};
		CustomXmlPartCollection.prototype.getByNamespace=function (namespaceUri) {
			return _createMethodObject(ExcelOp.CustomXmlPartScopedCollection, this, "GetByNamespace", 1, [namespaceUri], true, false, null, 4);
		};
		CustomXmlPartCollection.prototype.getItem=function (id) {
			return _createIndexerObject(ExcelOp.CustomXmlPart, this, [id]);
		};
		CustomXmlPartCollection.prototype.getItemOrNullObject=function (id) {
			return _createMethodObject(ExcelOp.CustomXmlPart, this, "GetItemOrNullObject", 1, [id], false, false, null, 4);
		};
		CustomXmlPartCollection.prototype.getCount=function () {
			return _invokeMethod(this, "GetCount", 1, [], 4, 0);
		};
		CustomXmlPartCollection.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		CustomXmlPartCollection.prototype.toJSON=function () {
			return {};
		};
		return CustomXmlPartCollection;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.CustomXmlPartCollection=CustomXmlPartCollection;
	var CustomXmlPart=(function (_super) {
		__extends(CustomXmlPart, _super);
		function CustomXmlPart() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(CustomXmlPart.prototype, "_className", {
			get: function () {
				return "CustomXmlPart";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(CustomXmlPart.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["id", "namespaceUri"];
			},
			enumerable: true,
			configurable: true
		});
		CustomXmlPart.prototype["delete"]=function () {
			return _invokeMethod(this, "Delete", 0, [], 0);
		};
		CustomXmlPart.prototype.getXml=function () {
			return _invokeMethod(this, "GetXml", 1, [], 4, 0);
		};
		CustomXmlPart.prototype.setXml=function (xml) {
			return _invokeMethod(this, "SetXml", 0, [xml], 0);
		};
		CustomXmlPart.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		CustomXmlPart.prototype.toJSON=function () {
			return {};
		};
		return CustomXmlPart;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.CustomXmlPart=CustomXmlPart;
	var _V1Api=(function (_super) {
		__extends(_V1Api, _super);
		function _V1Api() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(_V1Api.prototype, "_className", {
			get: function () {
				return "_V1Api";
			},
			enumerable: true,
			configurable: true
		});
		_V1Api.prototype.bindingAddColumns=function (input) {
			return _invokeMethod(this, "BindingAddColumns", 0, [input], 0, 0);
		};
		_V1Api.prototype.bindingAddFromNamedItem=function (input) {
			return _invokeMethod(this, "BindingAddFromNamedItem", 1, [input], 0, 0);
		};
		_V1Api.prototype.bindingAddFromPrompt=function (input) {
			return _invokeMethod(this, "BindingAddFromPrompt", 1, [input], 0, 0);
		};
		_V1Api.prototype.bindingAddFromSelection=function (input) {
			return _invokeMethod(this, "BindingAddFromSelection", 1, [input], 0, 0);
		};
		_V1Api.prototype.bindingAddRows=function (input) {
			return _invokeMethod(this, "BindingAddRows", 0, [input], 0, 0);
		};
		_V1Api.prototype.bindingClearFormats=function (input) {
			return _invokeMethod(this, "BindingClearFormats", 0, [input], 0, 0);
		};
		_V1Api.prototype.bindingDeleteAllDataValues=function (input) {
			return _invokeMethod(this, "BindingDeleteAllDataValues", 0, [input], 0, 0);
		};
		_V1Api.prototype.bindingGetAll=function () {
			return _invokeMethod(this, "BindingGetAll", 1, [], 4, 0);
		};
		_V1Api.prototype.bindingGetById=function (input) {
			return _invokeMethod(this, "BindingGetById", 1, [input], 4, 0);
		};
		_V1Api.prototype.bindingGetData=function (input) {
			return _invokeMethod(this, "BindingGetData", 1, [input], 4, 0);
		};
		_V1Api.prototype.bindingReleaseById=function (input) {
			return _invokeMethod(this, "BindingReleaseById", 1, [input], 0, 0);
		};
		_V1Api.prototype.bindingSetData=function (input) {
			return _invokeMethod(this, "BindingSetData", 0, [input], 0, 0);
		};
		_V1Api.prototype.bindingSetFormats=function (input) {
			return _invokeMethod(this, "BindingSetFormats", 0, [input], 0, 0);
		};
		_V1Api.prototype.bindingSetTableOptions=function (input) {
			return _invokeMethod(this, "BindingSetTableOptions", 0, [input], 0, 0);
		};
		_V1Api.prototype.getFilePropertiesAsync=function () {
			_throwIfApiNotSupported("_V1Api.getFilePropertiesAsync", _defaultApiSetName, "1.6", _hostName);
			return _invokeMethod(this, "GetFilePropertiesAsync", 1, [], 4, 0);
		};
		_V1Api.prototype.getSelectedData=function (input) {
			return _invokeMethod(this, "GetSelectedData", 1, [input], 4, 0);
		};
		_V1Api.prototype.gotoById=function (input) {
			return _invokeMethod(this, "GotoById", 1, [input], 4, 0);
		};
		_V1Api.prototype.setSelectedData=function (input) {
			return _invokeMethod(this, "SetSelectedData", 0, [input], 0, 0);
		};
		_V1Api.prototype.toJSON=function () {
			return {};
		};
		return _V1Api;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp._V1Api=_V1Api;
	var PivotTableCollection=(function (_super) {
		__extends(PivotTableCollection, _super);
		function PivotTableCollection() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(PivotTableCollection.prototype, "_className", {
			get: function () {
				return "PivotTableCollection";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(PivotTableCollection.prototype, "_isCollection", {
			get: function () {
				return true;
			},
			enumerable: true,
			configurable: true
		});
		PivotTableCollection.prototype.add=function (name, source, destination) {
			_throwIfApiNotSupported("PivotTableCollection.add", _defaultApiSetName, "1.8", _hostName);
			return _createAndInstantiateMethodObject(ExcelOp.PivotTable, this, "Add", 0, [name, source, destination], false, true, null, 0);
		};
		PivotTableCollection.prototype.getItem=function (name) {
			return _createIndexerObject(ExcelOp.PivotTable, this, [name]);
		};
		PivotTableCollection.prototype.getItemOrNullObject=function (name) {
			_throwIfApiNotSupported("PivotTableCollection.getItemOrNullObject", _defaultApiSetName, "1.4", _hostName);
			return _createMethodObject(ExcelOp.PivotTable, this, "GetItemOrNullObject", 1, [name], false, false, null, 4);
		};
		PivotTableCollection.prototype.getCount=function () {
			_throwIfApiNotSupported("PivotTableCollection.getCount", _defaultApiSetName, "1.4", _hostName);
			return _invokeMethod(this, "GetCount", 1, [], 4, 0);
		};
		PivotTableCollection.prototype.refreshAll=function () {
			return _invokeMethod(this, "RefreshAll", 0, [], 0);
		};
		PivotTableCollection.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		PivotTableCollection.prototype.toJSON=function () {
			return {};
		};
		return PivotTableCollection;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.PivotTableCollection=PivotTableCollection;
	var PivotTable=(function (_super) {
		__extends(PivotTable, _super);
		function PivotTable() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(PivotTable.prototype, "_className", {
			get: function () {
				return "PivotTable";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(PivotTable.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["name", "id", "useCustomSortLists", "enableDataValueEditing"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(PivotTable.prototype, "_scalarPropertyUpdateable", {
			get: function () {
				return [true, false, true, true];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(PivotTable.prototype, "_navigationPropertyNames", {
			get: function () {
				return ["worksheet", "hierarchies", "rowHierarchies", "columnHierarchies", "dataHierarchies", "filterHierarchies", "layout"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(PivotTable.prototype, "columnHierarchies", {
			get: function () {
				_throwIfApiNotSupported("PivotTable.columnHierarchies", _defaultApiSetName, "1.8", _hostName);
				return _createPropertyObject(ExcelOp.RowColumnPivotHierarchyCollection, this, "ColumnHierarchies", true, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(PivotTable.prototype, "dataHierarchies", {
			get: function () {
				_throwIfApiNotSupported("PivotTable.dataHierarchies", _defaultApiSetName, "1.8", _hostName);
				return _createPropertyObject(ExcelOp.DataPivotHierarchyCollection, this, "DataHierarchies", true, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(PivotTable.prototype, "filterHierarchies", {
			get: function () {
				_throwIfApiNotSupported("PivotTable.filterHierarchies", _defaultApiSetName, "1.8", _hostName);
				return _createPropertyObject(ExcelOp.FilterPivotHierarchyCollection, this, "FilterHierarchies", true, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(PivotTable.prototype, "hierarchies", {
			get: function () {
				_throwIfApiNotSupported("PivotTable.hierarchies", _defaultApiSetName, "1.8", _hostName);
				return _createPropertyObject(ExcelOp.PivotHierarchyCollection, this, "Hierarchies", true, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(PivotTable.prototype, "layout", {
			get: function () {
				_throwIfApiNotSupported("PivotTable.layout", _defaultApiSetName, "1.8", _hostName);
				return _createPropertyObject(ExcelOp.PivotLayout, this, "Layout", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(PivotTable.prototype, "rowHierarchies", {
			get: function () {
				_throwIfApiNotSupported("PivotTable.rowHierarchies", _defaultApiSetName, "1.8", _hostName);
				return _createPropertyObject(ExcelOp.RowColumnPivotHierarchyCollection, this, "RowHierarchies", true, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(PivotTable.prototype, "worksheet", {
			get: function () {
				return _createPropertyObject(ExcelOp.Worksheet, this, "Worksheet", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		PivotTable.prototype.update=function (properties) {
			return _invokeRecursiveUpdate(this, properties);
		};
		PivotTable.prototype["delete"]=function () {
			_throwIfApiNotSupported("PivotTable.delete", _defaultApiSetName, "1.8", _hostName);
			return _invokeMethod(this, "Delete", 0, [], 0);
		};
		PivotTable.prototype.refresh=function () {
			return _invokeMethod(this, "Refresh", 0, [], 0);
		};
		PivotTable.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		PivotTable.prototype.toJSON=function () {
			return {};
		};
		return PivotTable;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.PivotTable=PivotTable;
	var PivotLayout=(function (_super) {
		__extends(PivotLayout, _super);
		function PivotLayout() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(PivotLayout.prototype, "_className", {
			get: function () {
				return "PivotLayout";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(PivotLayout.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["showColumnGrandTotals", "showRowGrandTotals", "enableFieldList", "subtotalLocation", "layoutType", "autoFormat", "preserveFormatting"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(PivotLayout.prototype, "_scalarPropertyUpdateable", {
			get: function () {
				return [true, true, true, true, true, true, true];
			},
			enumerable: true,
			configurable: true
		});
		PivotLayout.prototype.getCell=function (dataHierarchy, rowItems, columnItems) {
			_throwIfApiNotSupported("PivotLayout.getCell", _defaultApiSetName, "1.9", _hostName);
			return _createAndInstantiateMethodObject(ExcelOp.Range, this, "GetCell", 0, [dataHierarchy, rowItems, columnItems], false, false, null, 0);
		};
		PivotLayout.prototype.getColumnLabelRange=function () {
			return _createAndInstantiateMethodObject(ExcelOp.Range, this, "GetColumnLabelRange", 0, [], false, false, null, 0);
		};
		PivotLayout.prototype.getDataBodyRange=function () {
			return _createAndInstantiateMethodObject(ExcelOp.Range, this, "GetDataBodyRange", 0, [], false, false, null, 0);
		};
		PivotLayout.prototype.getDataHierarchy=function (cell) {
			_throwIfApiNotSupported("PivotLayout.getDataHierarchy", _defaultApiSetName, "1.9", _hostName);
			return _createAndInstantiateMethodObject(ExcelOp.DataPivotHierarchy, this, "GetDataHierarchy", 0, [cell], false, false, null, 0);
		};
		PivotLayout.prototype.getFilterAxisRange=function () {
			return _createAndInstantiateMethodObject(ExcelOp.Range, this, "GetFilterAxisRange", 0, [], false, false, null, 0);
		};
		PivotLayout.prototype.getRange=function () {
			return _createAndInstantiateMethodObject(ExcelOp.Range, this, "GetRange", 0, [], false, false, null, 0);
		};
		PivotLayout.prototype.getRowLabelRange=function () {
			return _createAndInstantiateMethodObject(ExcelOp.Range, this, "GetRowLabelRange", 0, [], false, false, null, 0);
		};
		PivotLayout.prototype.update=function (properties) {
			return _invokeRecursiveUpdate(this, properties);
		};
		PivotLayout.prototype.getPivotItems=function (axis, cell) {
			_throwIfApiNotSupported("PivotLayout.getPivotItems", _defaultApiSetName, "1.9", _hostName);
			return _invokeMethod(this, "GetPivotItems", 0, [axis, cell], 0, 0);
		};
		PivotLayout.prototype.setAutosortOnCell=function (cell, sortby) {
			_throwIfApiNotSupported("PivotLayout.setAutosortOnCell", _defaultApiSetName, "1.9", _hostName);
			return _invokeMethod(this, "SetAutosortOnCell", 0, [cell, sortby], 0);
		};
		PivotLayout.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		PivotLayout.prototype.toJSON=function () {
			return {};
		};
		return PivotLayout;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.PivotLayout=PivotLayout;
	var PivotHierarchyCollection=(function (_super) {
		__extends(PivotHierarchyCollection, _super);
		function PivotHierarchyCollection() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(PivotHierarchyCollection.prototype, "_className", {
			get: function () {
				return "PivotHierarchyCollection";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(PivotHierarchyCollection.prototype, "_isCollection", {
			get: function () {
				return true;
			},
			enumerable: true,
			configurable: true
		});
		PivotHierarchyCollection.prototype.getItem=function (name) {
			return _createIndexerObject(ExcelOp.PivotHierarchy, this, [name]);
		};
		PivotHierarchyCollection.prototype.getItemOrNullObject=function (name) {
			return _createMethodObject(ExcelOp.PivotHierarchy, this, "GetItemOrNullObject", 1, [name], false, false, null, 4);
		};
		PivotHierarchyCollection.prototype.getCount=function () {
			return _invokeMethod(this, "GetCount", 1, [], 4, 0);
		};
		PivotHierarchyCollection.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		PivotHierarchyCollection.prototype.toJSON=function () {
			return {};
		};
		return PivotHierarchyCollection;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.PivotHierarchyCollection=PivotHierarchyCollection;
	var PivotHierarchy=(function (_super) {
		__extends(PivotHierarchy, _super);
		function PivotHierarchy() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(PivotHierarchy.prototype, "_className", {
			get: function () {
				return "PivotHierarchy";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(PivotHierarchy.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["id", "name"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(PivotHierarchy.prototype, "_scalarPropertyUpdateable", {
			get: function () {
				return [false, true];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(PivotHierarchy.prototype, "_navigationPropertyNames", {
			get: function () {
				return ["fields"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(PivotHierarchy.prototype, "fields", {
			get: function () {
				return _createPropertyObject(ExcelOp.PivotFieldCollection, this, "Fields", true, 4);
			},
			enumerable: true,
			configurable: true
		});
		PivotHierarchy.prototype.update=function (properties) {
			return _invokeRecursiveUpdate(this, properties);
		};
		PivotHierarchy.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		PivotHierarchy.prototype.toJSON=function () {
			return {};
		};
		return PivotHierarchy;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.PivotHierarchy=PivotHierarchy;
	var RowColumnPivotHierarchyCollection=(function (_super) {
		__extends(RowColumnPivotHierarchyCollection, _super);
		function RowColumnPivotHierarchyCollection() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(RowColumnPivotHierarchyCollection.prototype, "_className", {
			get: function () {
				return "RowColumnPivotHierarchyCollection";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(RowColumnPivotHierarchyCollection.prototype, "_isCollection", {
			get: function () {
				return true;
			},
			enumerable: true,
			configurable: true
		});
		RowColumnPivotHierarchyCollection.prototype.add=function (pivotHierarchy) {
			return _createAndInstantiateMethodObject(ExcelOp.RowColumnPivotHierarchy, this, "Add", 0, [pivotHierarchy], false, true, null, 0);
		};
		RowColumnPivotHierarchyCollection.prototype.getItem=function (name) {
			return _createIndexerObject(ExcelOp.RowColumnPivotHierarchy, this, [name]);
		};
		RowColumnPivotHierarchyCollection.prototype.getItemOrNullObject=function (name) {
			return _createMethodObject(ExcelOp.RowColumnPivotHierarchy, this, "GetItemOrNullObject", 1, [name], false, false, null, 4);
		};
		RowColumnPivotHierarchyCollection.prototype.getCount=function () {
			return _invokeMethod(this, "GetCount", 1, [], 4, 0);
		};
		RowColumnPivotHierarchyCollection.prototype.remove=function (rowColumnPivotHierarchy) {
			return _invokeMethod(this, "Remove", 0, [rowColumnPivotHierarchy], 0);
		};
		RowColumnPivotHierarchyCollection.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		RowColumnPivotHierarchyCollection.prototype.toJSON=function () {
			return {};
		};
		return RowColumnPivotHierarchyCollection;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.RowColumnPivotHierarchyCollection=RowColumnPivotHierarchyCollection;
	var RowColumnPivotHierarchy=(function (_super) {
		__extends(RowColumnPivotHierarchy, _super);
		function RowColumnPivotHierarchy() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(RowColumnPivotHierarchy.prototype, "_className", {
			get: function () {
				return "RowColumnPivotHierarchy";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(RowColumnPivotHierarchy.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["id", "name", "position"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(RowColumnPivotHierarchy.prototype, "_scalarPropertyUpdateable", {
			get: function () {
				return [false, true, true];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(RowColumnPivotHierarchy.prototype, "_navigationPropertyNames", {
			get: function () {
				return ["fields"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(RowColumnPivotHierarchy.prototype, "fields", {
			get: function () {
				return _createPropertyObject(ExcelOp.PivotFieldCollection, this, "Fields", true, 4);
			},
			enumerable: true,
			configurable: true
		});
		RowColumnPivotHierarchy.prototype.update=function (properties) {
			return _invokeRecursiveUpdate(this, properties);
		};
		RowColumnPivotHierarchy.prototype.setToDefault=function () {
			return _invokeMethod(this, "SetToDefault", 0, [], 0);
		};
		RowColumnPivotHierarchy.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		RowColumnPivotHierarchy.prototype.toJSON=function () {
			return {};
		};
		return RowColumnPivotHierarchy;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.RowColumnPivotHierarchy=RowColumnPivotHierarchy;
	var FilterPivotHierarchyCollection=(function (_super) {
		__extends(FilterPivotHierarchyCollection, _super);
		function FilterPivotHierarchyCollection() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(FilterPivotHierarchyCollection.prototype, "_className", {
			get: function () {
				return "FilterPivotHierarchyCollection";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(FilterPivotHierarchyCollection.prototype, "_isCollection", {
			get: function () {
				return true;
			},
			enumerable: true,
			configurable: true
		});
		FilterPivotHierarchyCollection.prototype.add=function (pivotHierarchy) {
			return _createAndInstantiateMethodObject(ExcelOp.FilterPivotHierarchy, this, "Add", 0, [pivotHierarchy], false, true, null, 0);
		};
		FilterPivotHierarchyCollection.prototype.getItem=function (name) {
			return _createIndexerObject(ExcelOp.FilterPivotHierarchy, this, [name]);
		};
		FilterPivotHierarchyCollection.prototype.getItemOrNullObject=function (name) {
			return _createMethodObject(ExcelOp.FilterPivotHierarchy, this, "GetItemOrNullObject", 1, [name], false, false, null, 4);
		};
		FilterPivotHierarchyCollection.prototype.getCount=function () {
			return _invokeMethod(this, "GetCount", 1, [], 4, 0);
		};
		FilterPivotHierarchyCollection.prototype.remove=function (filterPivotHierarchy) {
			return _invokeMethod(this, "Remove", 0, [filterPivotHierarchy], 0);
		};
		FilterPivotHierarchyCollection.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		FilterPivotHierarchyCollection.prototype.toJSON=function () {
			return {};
		};
		return FilterPivotHierarchyCollection;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.FilterPivotHierarchyCollection=FilterPivotHierarchyCollection;
	var FilterPivotHierarchy=(function (_super) {
		__extends(FilterPivotHierarchy, _super);
		function FilterPivotHierarchy() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(FilterPivotHierarchy.prototype, "_className", {
			get: function () {
				return "FilterPivotHierarchy";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(FilterPivotHierarchy.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["id", "name", "position", "enableMultipleFilterItems"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(FilterPivotHierarchy.prototype, "_scalarPropertyUpdateable", {
			get: function () {
				return [false, true, true, true];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(FilterPivotHierarchy.prototype, "_navigationPropertyNames", {
			get: function () {
				return ["fields"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(FilterPivotHierarchy.prototype, "fields", {
			get: function () {
				return _createPropertyObject(ExcelOp.PivotFieldCollection, this, "Fields", true, 4);
			},
			enumerable: true,
			configurable: true
		});
		FilterPivotHierarchy.prototype.update=function (properties) {
			return _invokeRecursiveUpdate(this, properties);
		};
		FilterPivotHierarchy.prototype.setToDefault=function () {
			return _invokeMethod(this, "SetToDefault", 0, [], 0);
		};
		FilterPivotHierarchy.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		FilterPivotHierarchy.prototype.toJSON=function () {
			return {};
		};
		return FilterPivotHierarchy;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.FilterPivotHierarchy=FilterPivotHierarchy;
	var DataPivotHierarchyCollection=(function (_super) {
		__extends(DataPivotHierarchyCollection, _super);
		function DataPivotHierarchyCollection() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(DataPivotHierarchyCollection.prototype, "_className", {
			get: function () {
				return "DataPivotHierarchyCollection";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(DataPivotHierarchyCollection.prototype, "_isCollection", {
			get: function () {
				return true;
			},
			enumerable: true,
			configurable: true
		});
		DataPivotHierarchyCollection.prototype.add=function (pivotHierarchy) {
			return _createAndInstantiateMethodObject(ExcelOp.DataPivotHierarchy, this, "Add", 0, [pivotHierarchy], false, true, null, 0);
		};
		DataPivotHierarchyCollection.prototype.getItem=function (name) {
			return _createIndexerObject(ExcelOp.DataPivotHierarchy, this, [name]);
		};
		DataPivotHierarchyCollection.prototype.getItemOrNullObject=function (name) {
			return _createMethodObject(ExcelOp.DataPivotHierarchy, this, "GetItemOrNullObject", 1, [name], false, false, null, 4);
		};
		DataPivotHierarchyCollection.prototype.getCount=function () {
			return _invokeMethod(this, "GetCount", 1, [], 4, 0);
		};
		DataPivotHierarchyCollection.prototype.remove=function (DataPivotHierarchy) {
			return _invokeMethod(this, "Remove", 0, [DataPivotHierarchy], 0);
		};
		DataPivotHierarchyCollection.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		DataPivotHierarchyCollection.prototype.toJSON=function () {
			return {};
		};
		return DataPivotHierarchyCollection;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.DataPivotHierarchyCollection=DataPivotHierarchyCollection;
	var DataPivotHierarchy=(function (_super) {
		__extends(DataPivotHierarchy, _super);
		function DataPivotHierarchy() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(DataPivotHierarchy.prototype, "_className", {
			get: function () {
				return "DataPivotHierarchy";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(DataPivotHierarchy.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["id", "name", "position", "numberFormat", "summarizeBy", "showAs"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(DataPivotHierarchy.prototype, "_scalarPropertyUpdateable", {
			get: function () {
				return [false, true, true, true, true, true];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(DataPivotHierarchy.prototype, "_navigationPropertyNames", {
			get: function () {
				return ["field"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(DataPivotHierarchy.prototype, "field", {
			get: function () {
				return _createPropertyObject(ExcelOp.PivotField, this, "Field", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		DataPivotHierarchy.prototype.update=function (properties) {
			return _invokeRecursiveUpdate(this, properties);
		};
		DataPivotHierarchy.prototype.setToDefault=function () {
			return _invokeMethod(this, "SetToDefault", 0, [], 0);
		};
		DataPivotHierarchy.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		DataPivotHierarchy.prototype.toJSON=function () {
			return {};
		};
		return DataPivotHierarchy;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.DataPivotHierarchy=DataPivotHierarchy;
	var PivotFieldCollection=(function (_super) {
		__extends(PivotFieldCollection, _super);
		function PivotFieldCollection() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(PivotFieldCollection.prototype, "_className", {
			get: function () {
				return "PivotFieldCollection";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(PivotFieldCollection.prototype, "_isCollection", {
			get: function () {
				return true;
			},
			enumerable: true,
			configurable: true
		});
		PivotFieldCollection.prototype.getItem=function (name) {
			return _createIndexerObject(ExcelOp.PivotField, this, [name]);
		};
		PivotFieldCollection.prototype.getItemOrNullObject=function (name) {
			return _createMethodObject(ExcelOp.PivotField, this, "GetItemOrNullObject", 1, [name], false, false, null, 4);
		};
		PivotFieldCollection.prototype.getCount=function () {
			return _invokeMethod(this, "GetCount", 1, [], 4, 0);
		};
		PivotFieldCollection.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		PivotFieldCollection.prototype.toJSON=function () {
			return {};
		};
		return PivotFieldCollection;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.PivotFieldCollection=PivotFieldCollection;
	var PivotField=(function (_super) {
		__extends(PivotField, _super);
		function PivotField() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(PivotField.prototype, "_className", {
			get: function () {
				return "PivotField";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(PivotField.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["id", "name", "subtotals", "showAllItems"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(PivotField.prototype, "_scalarPropertyUpdateable", {
			get: function () {
				return [false, true, true, true];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(PivotField.prototype, "_navigationPropertyNames", {
			get: function () {
				return ["items"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(PivotField.prototype, "items", {
			get: function () {
				return _createPropertyObject(ExcelOp.PivotItemCollection, this, "Items", true, 4);
			},
			enumerable: true,
			configurable: true
		});
		PivotField.prototype.update=function (properties) {
			return _invokeRecursiveUpdate(this, properties);
		};
		PivotField.prototype.sortByLabels=function (sortby) {
			return _invokeMethod(this, "SortByLabels", 0, [sortby], 0);
		};
		PivotField.prototype.sortByValues=function (sortby, valuesHierarchy, pivotItemScope) {
			_throwIfApiNotSupported("PivotField.sortByValues", _defaultApiSetName, "1.9", _hostName);
			return _invokeMethod(this, "SortByValues", 0, [sortby, valuesHierarchy, pivotItemScope], 0);
		};
		PivotField.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		PivotField.prototype.toJSON=function () {
			return {};
		};
		return PivotField;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.PivotField=PivotField;
	var PivotItemCollection=(function (_super) {
		__extends(PivotItemCollection, _super);
		function PivotItemCollection() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(PivotItemCollection.prototype, "_className", {
			get: function () {
				return "PivotItemCollection";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(PivotItemCollection.prototype, "_isCollection", {
			get: function () {
				return true;
			},
			enumerable: true,
			configurable: true
		});
		PivotItemCollection.prototype.getItem=function (name) {
			return _createIndexerObject(ExcelOp.PivotItem, this, [name]);
		};
		PivotItemCollection.prototype.getItemOrNullObject=function (name) {
			return _createMethodObject(ExcelOp.PivotItem, this, "GetItemOrNullObject", 1, [name], false, false, null, 4);
		};
		PivotItemCollection.prototype.getCount=function () {
			return _invokeMethod(this, "GetCount", 1, [], 4, 0);
		};
		PivotItemCollection.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		PivotItemCollection.prototype.toJSON=function () {
			return {};
		};
		return PivotItemCollection;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.PivotItemCollection=PivotItemCollection;
	var PivotItem=(function (_super) {
		__extends(PivotItem, _super);
		function PivotItem() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(PivotItem.prototype, "_className", {
			get: function () {
				return "PivotItem";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(PivotItem.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["id", "name", "isExpanded", "visible"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(PivotItem.prototype, "_scalarPropertyUpdateable", {
			get: function () {
				return [false, true, true, true];
			},
			enumerable: true,
			configurable: true
		});
		PivotItem.prototype.update=function (properties) {
			return _invokeRecursiveUpdate(this, properties);
		};
		PivotItem.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		PivotItem.prototype.toJSON=function () {
			return {};
		};
		return PivotItem;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.PivotItem=PivotItem;
	var PivotFilterTopBottomCriterion;
	(function (PivotFilterTopBottomCriterion) {
		PivotFilterTopBottomCriterion["invalid"]="Invalid";
		PivotFilterTopBottomCriterion["topItems"]="TopItems";
		PivotFilterTopBottomCriterion["topPercent"]="TopPercent";
		PivotFilterTopBottomCriterion["topSum"]="TopSum";
		PivotFilterTopBottomCriterion["bottomItems"]="BottomItems";
		PivotFilterTopBottomCriterion["bottomPercent"]="BottomPercent";
		PivotFilterTopBottomCriterion["bottomSum"]="BottomSum";
	})(PivotFilterTopBottomCriterion=ExcelOp.PivotFilterTopBottomCriterion || (ExcelOp.PivotFilterTopBottomCriterion={}));
	var SortBy;
	(function (SortBy) {
		SortBy["ascending"]="Ascending";
		SortBy["descending"]="Descending";
	})(SortBy=ExcelOp.SortBy || (ExcelOp.SortBy={}));
	var AggregationFunction;
	(function (AggregationFunction) {
		AggregationFunction["unknown"]="Unknown";
		AggregationFunction["automatic"]="Automatic";
		AggregationFunction["sum"]="Sum";
		AggregationFunction["count"]="Count";
		AggregationFunction["average"]="Average";
		AggregationFunction["max"]="Max";
		AggregationFunction["min"]="Min";
		AggregationFunction["product"]="Product";
		AggregationFunction["countNumbers"]="CountNumbers";
		AggregationFunction["standardDeviation"]="StandardDeviation";
		AggregationFunction["standardDeviationP"]="StandardDeviationP";
		AggregationFunction["variance"]="Variance";
		AggregationFunction["varianceP"]="VarianceP";
	})(AggregationFunction=ExcelOp.AggregationFunction || (ExcelOp.AggregationFunction={}));
	var ShowAsCalculation;
	(function (ShowAsCalculation) {
		ShowAsCalculation["unknown"]="Unknown";
		ShowAsCalculation["none"]="None";
		ShowAsCalculation["percentOfGrandTotal"]="PercentOfGrandTotal";
		ShowAsCalculation["percentOfRowTotal"]="PercentOfRowTotal";
		ShowAsCalculation["percentOfColumnTotal"]="PercentOfColumnTotal";
		ShowAsCalculation["percentOfParentRowTotal"]="PercentOfParentRowTotal";
		ShowAsCalculation["percentOfParentColumnTotal"]="PercentOfParentColumnTotal";
		ShowAsCalculation["percentOfParentTotal"]="PercentOfParentTotal";
		ShowAsCalculation["percentOf"]="PercentOf";
		ShowAsCalculation["runningTotal"]="RunningTotal";
		ShowAsCalculation["percentRunningTotal"]="PercentRunningTotal";
		ShowAsCalculation["differenceFrom"]="DifferenceFrom";
		ShowAsCalculation["percentDifferenceFrom"]="PercentDifferenceFrom";
		ShowAsCalculation["rankAscending"]="RankAscending";
		ShowAsCalculation["rankDecending"]="RankDecending";
		ShowAsCalculation["index"]="Index";
	})(ShowAsCalculation=ExcelOp.ShowAsCalculation || (ExcelOp.ShowAsCalculation={}));
	var PivotAxis;
	(function (PivotAxis) {
		PivotAxis["unknown"]="Unknown";
		PivotAxis["row"]="Row";
		PivotAxis["column"]="Column";
		PivotAxis["data"]="Data";
		PivotAxis["filter"]="Filter";
	})(PivotAxis=ExcelOp.PivotAxis || (ExcelOp.PivotAxis={}));
	var DocumentProperties=(function (_super) {
		__extends(DocumentProperties, _super);
		function DocumentProperties() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(DocumentProperties.prototype, "_className", {
			get: function () {
				return "DocumentProperties";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(DocumentProperties.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["title", "subject", "author", "keywords", "comments", "lastAuthor", "revisionNumber", "creationDate", "category", "manager", "company"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(DocumentProperties.prototype, "_scalarPropertyUpdateable", {
			get: function () {
				return [true, true, true, true, true, false, true, false, true, true, true];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(DocumentProperties.prototype, "_navigationPropertyNames", {
			get: function () {
				return ["custom"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(DocumentProperties.prototype, "custom", {
			get: function () {
				return _createPropertyObject(ExcelOp.CustomPropertyCollection, this, "Custom", true, 4);
			},
			enumerable: true,
			configurable: true
		});
		DocumentProperties.prototype.update=function (properties) {
			return _invokeRecursiveUpdate(this, properties);
		};
		DocumentProperties.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		DocumentProperties.prototype.toJSON=function () {
			return {};
		};
		return DocumentProperties;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.DocumentProperties=DocumentProperties;
	var CustomProperty=(function (_super) {
		__extends(CustomProperty, _super);
		function CustomProperty() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(CustomProperty.prototype, "_className", {
			get: function () {
				return "CustomProperty";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(CustomProperty.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["key", "value", "type"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(CustomProperty.prototype, "_scalarPropertyUpdateable", {
			get: function () {
				return [false, true, false];
			},
			enumerable: true,
			configurable: true
		});
		CustomProperty.prototype.update=function (properties) {
			return _invokeRecursiveUpdate(this, properties);
		};
		CustomProperty.prototype["delete"]=function () {
			return _invokeMethod(this, "Delete", 0, [], 0);
		};
		CustomProperty.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		CustomProperty.prototype.toJSON=function () {
			return {};
		};
		return CustomProperty;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.CustomProperty=CustomProperty;
	var CustomPropertyCollection=(function (_super) {
		__extends(CustomPropertyCollection, _super);
		function CustomPropertyCollection() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(CustomPropertyCollection.prototype, "_className", {
			get: function () {
				return "CustomPropertyCollection";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(CustomPropertyCollection.prototype, "_isCollection", {
			get: function () {
				return true;
			},
			enumerable: true,
			configurable: true
		});
		CustomPropertyCollection.prototype.add=function (key, value) {
			return _createAndInstantiateMethodObject(ExcelOp.CustomProperty, this, "Add", 0, [key, value], false, true, null, 0);
		};
		CustomPropertyCollection.prototype.getItem=function (key) {
			return _createIndexerObject(ExcelOp.CustomProperty, this, [key]);
		};
		CustomPropertyCollection.prototype.getItemOrNullObject=function (key) {
			return _createMethodObject(ExcelOp.CustomProperty, this, "GetItemOrNullObject", 1, [key], false, false, null, 4);
		};
		CustomPropertyCollection.prototype.deleteAll=function () {
			return _invokeMethod(this, "DeleteAll", 0, [], 0);
		};
		CustomPropertyCollection.prototype.getCount=function () {
			return _invokeMethod(this, "GetCount", 1, [], 4, 0);
		};
		CustomPropertyCollection.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		CustomPropertyCollection.prototype.toJSON=function () {
			return {};
		};
		return CustomPropertyCollection;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.CustomPropertyCollection=CustomPropertyCollection;
	var ConditionalFormatCollection=(function (_super) {
		__extends(ConditionalFormatCollection, _super);
		function ConditionalFormatCollection() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(ConditionalFormatCollection.prototype, "_className", {
			get: function () {
				return "ConditionalFormatCollection";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ConditionalFormatCollection.prototype, "_isCollection", {
			get: function () {
				return true;
			},
			enumerable: true,
			configurable: true
		});
		ConditionalFormatCollection.prototype.add=function (type) {
			return _createAndInstantiateMethodObject(ExcelOp.ConditionalFormat, this, "Add", 0, [type], false, true, null, 0);
		};
		ConditionalFormatCollection.prototype.getItem=function (id) {
			return _createIndexerObject(ExcelOp.ConditionalFormat, this, [id]);
		};
		ConditionalFormatCollection.prototype.getItemAt=function (index) {
			return _createMethodObject(ExcelOp.ConditionalFormat, this, "GetItemAt", 1, [index], false, false, null, 4);
		};
		ConditionalFormatCollection.prototype.clearAll=function () {
			return _invokeMethod(this, "ClearAll", 0, [], 0);
		};
		ConditionalFormatCollection.prototype.getCount=function () {
			return _invokeMethod(this, "GetCount", 1, [], 4, 0);
		};
		ConditionalFormatCollection.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		ConditionalFormatCollection.prototype.toJSON=function () {
			return {};
		};
		return ConditionalFormatCollection;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.ConditionalFormatCollection=ConditionalFormatCollection;
	var ConditionalFormat=(function (_super) {
		__extends(ConditionalFormat, _super);
		function ConditionalFormat() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(ConditionalFormat.prototype, "_className", {
			get: function () {
				return "ConditionalFormat";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ConditionalFormat.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["stopIfTrue", "priority", "type", "id"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ConditionalFormat.prototype, "_scalarPropertyUpdateable", {
			get: function () {
				return [true, true, false, false];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ConditionalFormat.prototype, "_navigationPropertyNames", {
			get: function () {
				return ["dataBarOrNullObject", "dataBar", "customOrNullObject", "custom", "iconSet", "iconSetOrNullObject", "colorScale", "colorScaleOrNullObject", "topBottom", "topBottomOrNullObject", "preset", "presetOrNullObject", "textComparison", "textComparisonOrNullObject", "cellValue", "cellValueOrNullObject"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ConditionalFormat.prototype, "cellValue", {
			get: function () {
				return _createPropertyObject(ExcelOp.CellValueConditionalFormat, this, "CellValue", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ConditionalFormat.prototype, "cellValueOrNullObject", {
			get: function () {
				return _createPropertyObject(ExcelOp.CellValueConditionalFormat, this, "CellValueOrNullObject", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ConditionalFormat.prototype, "colorScale", {
			get: function () {
				return _createPropertyObject(ExcelOp.ColorScaleConditionalFormat, this, "ColorScale", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ConditionalFormat.prototype, "colorScaleOrNullObject", {
			get: function () {
				return _createPropertyObject(ExcelOp.ColorScaleConditionalFormat, this, "ColorScaleOrNullObject", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ConditionalFormat.prototype, "custom", {
			get: function () {
				return _createPropertyObject(ExcelOp.CustomConditionalFormat, this, "Custom", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ConditionalFormat.prototype, "customOrNullObject", {
			get: function () {
				return _createPropertyObject(ExcelOp.CustomConditionalFormat, this, "CustomOrNullObject", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ConditionalFormat.prototype, "dataBar", {
			get: function () {
				return _createPropertyObject(ExcelOp.DataBarConditionalFormat, this, "DataBar", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ConditionalFormat.prototype, "dataBarOrNullObject", {
			get: function () {
				return _createPropertyObject(ExcelOp.DataBarConditionalFormat, this, "DataBarOrNullObject", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ConditionalFormat.prototype, "iconSet", {
			get: function () {
				return _createPropertyObject(ExcelOp.IconSetConditionalFormat, this, "IconSet", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ConditionalFormat.prototype, "iconSetOrNullObject", {
			get: function () {
				return _createPropertyObject(ExcelOp.IconSetConditionalFormat, this, "IconSetOrNullObject", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ConditionalFormat.prototype, "preset", {
			get: function () {
				return _createPropertyObject(ExcelOp.PresetCriteriaConditionalFormat, this, "Preset", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ConditionalFormat.prototype, "presetOrNullObject", {
			get: function () {
				return _createPropertyObject(ExcelOp.PresetCriteriaConditionalFormat, this, "PresetOrNullObject", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ConditionalFormat.prototype, "textComparison", {
			get: function () {
				return _createPropertyObject(ExcelOp.TextConditionalFormat, this, "TextComparison", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ConditionalFormat.prototype, "textComparisonOrNullObject", {
			get: function () {
				return _createPropertyObject(ExcelOp.TextConditionalFormat, this, "TextComparisonOrNullObject", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ConditionalFormat.prototype, "topBottom", {
			get: function () {
				return _createPropertyObject(ExcelOp.TopBottomConditionalFormat, this, "TopBottom", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ConditionalFormat.prototype, "topBottomOrNullObject", {
			get: function () {
				return _createPropertyObject(ExcelOp.TopBottomConditionalFormat, this, "TopBottomOrNullObject", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		ConditionalFormat.prototype.getRange=function () {
			return _createMethodObject(ExcelOp.Range, this, "GetRange", 1, [], false, true, null, 4);
		};
		ConditionalFormat.prototype.getRangeOrNullObject=function () {
			return _createMethodObject(ExcelOp.Range, this, "GetRangeOrNullObject", 1, [], false, true, null, 4);
		};
		ConditionalFormat.prototype.getRanges=function () {
			_throwIfApiNotSupported("ConditionalFormat.getRanges", _defaultApiSetName, "1.9", _hostName);
			return _createMethodObject(ExcelOp.RangeAreas, this, "GetRanges", 1, [], false, true, null, 4);
		};
		ConditionalFormat.prototype.update=function (properties) {
			return _invokeRecursiveUpdate(this, properties);
		};
		ConditionalFormat.prototype["delete"]=function () {
			return _invokeMethod(this, "Delete", 0, [], 0);
		};
		ConditionalFormat.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		ConditionalFormat.prototype.toJSON=function () {
			return {};
		};
		return ConditionalFormat;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.ConditionalFormat=ConditionalFormat;
	var DataBarConditionalFormat=(function (_super) {
		__extends(DataBarConditionalFormat, _super);
		function DataBarConditionalFormat() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(DataBarConditionalFormat.prototype, "_className", {
			get: function () {
				return "DataBarConditionalFormat";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(DataBarConditionalFormat.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["showDataBarOnly", "barDirection", "axisFormat", "axisColor", "lowerBoundRule", "upperBoundRule"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(DataBarConditionalFormat.prototype, "_scalarPropertyUpdateable", {
			get: function () {
				return [true, true, true, true, true, true];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(DataBarConditionalFormat.prototype, "_navigationPropertyNames", {
			get: function () {
				return ["positiveFormat", "negativeFormat"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(DataBarConditionalFormat.prototype, "negativeFormat", {
			get: function () {
				return _createPropertyObject(ExcelOp.ConditionalDataBarNegativeFormat, this, "NegativeFormat", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(DataBarConditionalFormat.prototype, "positiveFormat", {
			get: function () {
				return _createPropertyObject(ExcelOp.ConditionalDataBarPositiveFormat, this, "PositiveFormat", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		DataBarConditionalFormat.prototype.update=function (properties) {
			return _invokeRecursiveUpdate(this, properties);
		};
		DataBarConditionalFormat.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		DataBarConditionalFormat.prototype.toJSON=function () {
			return {};
		};
		return DataBarConditionalFormat;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.DataBarConditionalFormat=DataBarConditionalFormat;
	var ConditionalDataBarPositiveFormat=(function (_super) {
		__extends(ConditionalDataBarPositiveFormat, _super);
		function ConditionalDataBarPositiveFormat() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(ConditionalDataBarPositiveFormat.prototype, "_className", {
			get: function () {
				return "ConditionalDataBarPositiveFormat";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ConditionalDataBarPositiveFormat.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["fillColor", "gradientFill", "borderColor"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ConditionalDataBarPositiveFormat.prototype, "_scalarPropertyUpdateable", {
			get: function () {
				return [true, true, true];
			},
			enumerable: true,
			configurable: true
		});
		ConditionalDataBarPositiveFormat.prototype.update=function (properties) {
			return _invokeRecursiveUpdate(this, properties);
		};
		ConditionalDataBarPositiveFormat.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		ConditionalDataBarPositiveFormat.prototype.toJSON=function () {
			return {};
		};
		return ConditionalDataBarPositiveFormat;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.ConditionalDataBarPositiveFormat=ConditionalDataBarPositiveFormat;
	var ConditionalDataBarNegativeFormat=(function (_super) {
		__extends(ConditionalDataBarNegativeFormat, _super);
		function ConditionalDataBarNegativeFormat() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(ConditionalDataBarNegativeFormat.prototype, "_className", {
			get: function () {
				return "ConditionalDataBarNegativeFormat";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ConditionalDataBarNegativeFormat.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["fillColor", "matchPositiveFillColor", "borderColor", "matchPositiveBorderColor"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ConditionalDataBarNegativeFormat.prototype, "_scalarPropertyUpdateable", {
			get: function () {
				return [true, true, true, true];
			},
			enumerable: true,
			configurable: true
		});
		ConditionalDataBarNegativeFormat.prototype.update=function (properties) {
			return _invokeRecursiveUpdate(this, properties);
		};
		ConditionalDataBarNegativeFormat.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		ConditionalDataBarNegativeFormat.prototype.toJSON=function () {
			return {};
		};
		return ConditionalDataBarNegativeFormat;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.ConditionalDataBarNegativeFormat=ConditionalDataBarNegativeFormat;
	var CustomConditionalFormat=(function (_super) {
		__extends(CustomConditionalFormat, _super);
		function CustomConditionalFormat() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(CustomConditionalFormat.prototype, "_className", {
			get: function () {
				return "CustomConditionalFormat";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(CustomConditionalFormat.prototype, "_navigationPropertyNames", {
			get: function () {
				return ["rule", "format"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(CustomConditionalFormat.prototype, "format", {
			get: function () {
				return _createPropertyObject(ExcelOp.ConditionalRangeFormat, this, "Format", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(CustomConditionalFormat.prototype, "rule", {
			get: function () {
				return _createPropertyObject(ExcelOp.ConditionalFormatRule, this, "Rule", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		CustomConditionalFormat.prototype.toJSON=function () {
			return {};
		};
		return CustomConditionalFormat;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.CustomConditionalFormat=CustomConditionalFormat;
	var ConditionalFormatRule=(function (_super) {
		__extends(ConditionalFormatRule, _super);
		function ConditionalFormatRule() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(ConditionalFormatRule.prototype, "_className", {
			get: function () {
				return "ConditionalFormatRule";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ConditionalFormatRule.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["formula", "formulaLocal", "formulaR1C1"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ConditionalFormatRule.prototype, "_scalarPropertyUpdateable", {
			get: function () {
				return [true, true, true];
			},
			enumerable: true,
			configurable: true
		});
		ConditionalFormatRule.prototype.update=function (properties) {
			return _invokeRecursiveUpdate(this, properties);
		};
		ConditionalFormatRule.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		ConditionalFormatRule.prototype.toJSON=function () {
			return {};
		};
		return ConditionalFormatRule;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.ConditionalFormatRule=ConditionalFormatRule;
	var IconSetConditionalFormat=(function (_super) {
		__extends(IconSetConditionalFormat, _super);
		function IconSetConditionalFormat() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(IconSetConditionalFormat.prototype, "_className", {
			get: function () {
				return "IconSetConditionalFormat";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(IconSetConditionalFormat.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["reverseIconOrder", "showIconOnly", "style", "criteria"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(IconSetConditionalFormat.prototype, "_scalarPropertyUpdateable", {
			get: function () {
				return [true, true, true, true];
			},
			enumerable: true,
			configurable: true
		});
		IconSetConditionalFormat.prototype.update=function (properties) {
			return _invokeRecursiveUpdate(this, properties);
		};
		IconSetConditionalFormat.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		IconSetConditionalFormat.prototype.toJSON=function () {
			return {};
		};
		return IconSetConditionalFormat;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.IconSetConditionalFormat=IconSetConditionalFormat;
	var ColorScaleConditionalFormat=(function (_super) {
		__extends(ColorScaleConditionalFormat, _super);
		function ColorScaleConditionalFormat() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(ColorScaleConditionalFormat.prototype, "_className", {
			get: function () {
				return "ColorScaleConditionalFormat";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ColorScaleConditionalFormat.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["threeColorScale", "criteria"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ColorScaleConditionalFormat.prototype, "_scalarPropertyUpdateable", {
			get: function () {
				return [false, true];
			},
			enumerable: true,
			configurable: true
		});
		ColorScaleConditionalFormat.prototype.update=function (properties) {
			return _invokeRecursiveUpdate(this, properties);
		};
		ColorScaleConditionalFormat.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		ColorScaleConditionalFormat.prototype.toJSON=function () {
			return {};
		};
		return ColorScaleConditionalFormat;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.ColorScaleConditionalFormat=ColorScaleConditionalFormat;
	var TopBottomConditionalFormat=(function (_super) {
		__extends(TopBottomConditionalFormat, _super);
		function TopBottomConditionalFormat() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(TopBottomConditionalFormat.prototype, "_className", {
			get: function () {
				return "TopBottomConditionalFormat";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TopBottomConditionalFormat.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["rule"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TopBottomConditionalFormat.prototype, "_scalarPropertyUpdateable", {
			get: function () {
				return [true];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TopBottomConditionalFormat.prototype, "_navigationPropertyNames", {
			get: function () {
				return ["format"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TopBottomConditionalFormat.prototype, "format", {
			get: function () {
				return _createPropertyObject(ExcelOp.ConditionalRangeFormat, this, "Format", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		TopBottomConditionalFormat.prototype.update=function (properties) {
			return _invokeRecursiveUpdate(this, properties);
		};
		TopBottomConditionalFormat.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		TopBottomConditionalFormat.prototype.toJSON=function () {
			return {};
		};
		return TopBottomConditionalFormat;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.TopBottomConditionalFormat=TopBottomConditionalFormat;
	var PresetCriteriaConditionalFormat=(function (_super) {
		__extends(PresetCriteriaConditionalFormat, _super);
		function PresetCriteriaConditionalFormat() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(PresetCriteriaConditionalFormat.prototype, "_className", {
			get: function () {
				return "PresetCriteriaConditionalFormat";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(PresetCriteriaConditionalFormat.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["rule"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(PresetCriteriaConditionalFormat.prototype, "_scalarPropertyUpdateable", {
			get: function () {
				return [true];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(PresetCriteriaConditionalFormat.prototype, "_navigationPropertyNames", {
			get: function () {
				return ["format"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(PresetCriteriaConditionalFormat.prototype, "format", {
			get: function () {
				return _createPropertyObject(ExcelOp.ConditionalRangeFormat, this, "Format", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		PresetCriteriaConditionalFormat.prototype.update=function (properties) {
			return _invokeRecursiveUpdate(this, properties);
		};
		PresetCriteriaConditionalFormat.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		PresetCriteriaConditionalFormat.prototype.toJSON=function () {
			return {};
		};
		return PresetCriteriaConditionalFormat;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.PresetCriteriaConditionalFormat=PresetCriteriaConditionalFormat;
	var TextConditionalFormat=(function (_super) {
		__extends(TextConditionalFormat, _super);
		function TextConditionalFormat() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(TextConditionalFormat.prototype, "_className", {
			get: function () {
				return "TextConditionalFormat";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TextConditionalFormat.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["rule"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TextConditionalFormat.prototype, "_scalarPropertyUpdateable", {
			get: function () {
				return [true];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TextConditionalFormat.prototype, "_navigationPropertyNames", {
			get: function () {
				return ["format"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TextConditionalFormat.prototype, "format", {
			get: function () {
				return _createPropertyObject(ExcelOp.ConditionalRangeFormat, this, "Format", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		TextConditionalFormat.prototype.update=function (properties) {
			return _invokeRecursiveUpdate(this, properties);
		};
		TextConditionalFormat.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		TextConditionalFormat.prototype.toJSON=function () {
			return {};
		};
		return TextConditionalFormat;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.TextConditionalFormat=TextConditionalFormat;
	var CellValueConditionalFormat=(function (_super) {
		__extends(CellValueConditionalFormat, _super);
		function CellValueConditionalFormat() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(CellValueConditionalFormat.prototype, "_className", {
			get: function () {
				return "CellValueConditionalFormat";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(CellValueConditionalFormat.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["rule"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(CellValueConditionalFormat.prototype, "_scalarPropertyUpdateable", {
			get: function () {
				return [true];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(CellValueConditionalFormat.prototype, "_navigationPropertyNames", {
			get: function () {
				return ["format"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(CellValueConditionalFormat.prototype, "format", {
			get: function () {
				return _createPropertyObject(ExcelOp.ConditionalRangeFormat, this, "Format", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		CellValueConditionalFormat.prototype.update=function (properties) {
			return _invokeRecursiveUpdate(this, properties);
		};
		CellValueConditionalFormat.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		CellValueConditionalFormat.prototype.toJSON=function () {
			return {};
		};
		return CellValueConditionalFormat;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.CellValueConditionalFormat=CellValueConditionalFormat;
	var ConditionalRangeFormat=(function (_super) {
		__extends(ConditionalRangeFormat, _super);
		function ConditionalRangeFormat() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(ConditionalRangeFormat.prototype, "_className", {
			get: function () {
				return "ConditionalRangeFormat";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ConditionalRangeFormat.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["numberFormat"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ConditionalRangeFormat.prototype, "_scalarPropertyUpdateable", {
			get: function () {
				return [true];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ConditionalRangeFormat.prototype, "_navigationPropertyNames", {
			get: function () {
				return ["fill", "font", "borders"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ConditionalRangeFormat.prototype, "borders", {
			get: function () {
				return _createPropertyObject(ExcelOp.ConditionalRangeBorderCollection, this, "Borders", true, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ConditionalRangeFormat.prototype, "fill", {
			get: function () {
				return _createPropertyObject(ExcelOp.ConditionalRangeFill, this, "Fill", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ConditionalRangeFormat.prototype, "font", {
			get: function () {
				return _createPropertyObject(ExcelOp.ConditionalRangeFont, this, "Font", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		ConditionalRangeFormat.prototype.update=function (properties) {
			return _invokeRecursiveUpdate(this, properties);
		};
		ConditionalRangeFormat.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		ConditionalRangeFormat.prototype.toJSON=function () {
			return {};
		};
		return ConditionalRangeFormat;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.ConditionalRangeFormat=ConditionalRangeFormat;
	var ConditionalRangeFont=(function (_super) {
		__extends(ConditionalRangeFont, _super);
		function ConditionalRangeFont() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(ConditionalRangeFont.prototype, "_className", {
			get: function () {
				return "ConditionalRangeFont";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ConditionalRangeFont.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["color", "italic", "bold", "underline", "strikethrough"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ConditionalRangeFont.prototype, "_scalarPropertyUpdateable", {
			get: function () {
				return [true, true, true, true, true];
			},
			enumerable: true,
			configurable: true
		});
		ConditionalRangeFont.prototype.update=function (properties) {
			return _invokeRecursiveUpdate(this, properties);
		};
		ConditionalRangeFont.prototype.clear=function () {
			return _invokeMethod(this, "Clear", 0, [], 0);
		};
		ConditionalRangeFont.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		ConditionalRangeFont.prototype.toJSON=function () {
			return {};
		};
		return ConditionalRangeFont;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.ConditionalRangeFont=ConditionalRangeFont;
	var ConditionalRangeFill=(function (_super) {
		__extends(ConditionalRangeFill, _super);
		function ConditionalRangeFill() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(ConditionalRangeFill.prototype, "_className", {
			get: function () {
				return "ConditionalRangeFill";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ConditionalRangeFill.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["color"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ConditionalRangeFill.prototype, "_scalarPropertyUpdateable", {
			get: function () {
				return [true];
			},
			enumerable: true,
			configurable: true
		});
		ConditionalRangeFill.prototype.update=function (properties) {
			return _invokeRecursiveUpdate(this, properties);
		};
		ConditionalRangeFill.prototype.clear=function () {
			return _invokeMethod(this, "Clear", 0, [], 0);
		};
		ConditionalRangeFill.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		ConditionalRangeFill.prototype.toJSON=function () {
			return {};
		};
		return ConditionalRangeFill;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.ConditionalRangeFill=ConditionalRangeFill;
	var ConditionalRangeBorder=(function (_super) {
		__extends(ConditionalRangeBorder, _super);
		function ConditionalRangeBorder() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(ConditionalRangeBorder.prototype, "_className", {
			get: function () {
				return "ConditionalRangeBorder";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ConditionalRangeBorder.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["sideIndex", "style", "color"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ConditionalRangeBorder.prototype, "_scalarPropertyUpdateable", {
			get: function () {
				return [false, true, true];
			},
			enumerable: true,
			configurable: true
		});
		ConditionalRangeBorder.prototype.update=function (properties) {
			return _invokeRecursiveUpdate(this, properties);
		};
		ConditionalRangeBorder.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		ConditionalRangeBorder.prototype.toJSON=function () {
			return {};
		};
		return ConditionalRangeBorder;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.ConditionalRangeBorder=ConditionalRangeBorder;
	var ConditionalRangeBorderCollection=(function (_super) {
		__extends(ConditionalRangeBorderCollection, _super);
		function ConditionalRangeBorderCollection() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(ConditionalRangeBorderCollection.prototype, "_className", {
			get: function () {
				return "ConditionalRangeBorderCollection";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ConditionalRangeBorderCollection.prototype, "_isCollection", {
			get: function () {
				return true;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ConditionalRangeBorderCollection.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["count"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ConditionalRangeBorderCollection.prototype, "_navigationPropertyNames", {
			get: function () {
				return ["top", "bottom", "left", "right"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ConditionalRangeBorderCollection.prototype, "bottom", {
			get: function () {
				return _createPropertyObject(ExcelOp.ConditionalRangeBorder, this, "Bottom", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ConditionalRangeBorderCollection.prototype, "left", {
			get: function () {
				return _createPropertyObject(ExcelOp.ConditionalRangeBorder, this, "Left", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ConditionalRangeBorderCollection.prototype, "right", {
			get: function () {
				return _createPropertyObject(ExcelOp.ConditionalRangeBorder, this, "Right", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ConditionalRangeBorderCollection.prototype, "top", {
			get: function () {
				return _createPropertyObject(ExcelOp.ConditionalRangeBorder, this, "Top", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		ConditionalRangeBorderCollection.prototype.getItem=function (index) {
			return _createIndexerObject(ExcelOp.ConditionalRangeBorder, this, [index]);
		};
		ConditionalRangeBorderCollection.prototype.getItemAt=function (index) {
			return _createMethodObject(ExcelOp.ConditionalRangeBorder, this, "GetItemAt", 1, [index], false, false, null, 4);
		};
		ConditionalRangeBorderCollection.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		ConditionalRangeBorderCollection.prototype.toJSON=function () {
			return {};
		};
		return ConditionalRangeBorderCollection;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.ConditionalRangeBorderCollection=ConditionalRangeBorderCollection;
	var NumberFormattingService=(function (_super) {
		__extends(NumberFormattingService, _super);
		function NumberFormattingService() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(NumberFormattingService.prototype, "_className", {
			get: function () {
				return "NumberFormattingService";
			},
			enumerable: true,
			configurable: true
		});
		NumberFormattingService.prototype.getFormatter=function (format) {
			return _createAndInstantiateMethodObject(ExcelOp.NumberFormatter, this, "GetFormatter", 0, [format], false, false, null, 0);
		};
		NumberFormattingService.prototype.toJSON=function () {
			return {};
		};
		return NumberFormattingService;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.NumberFormattingService=NumberFormattingService;
	ExcelOp.numberFormattingService=_createTopLevelServiceObject(NumberFormattingService, _localDocumentContext, "Microsoft.ExcelServices.NumberFormattingService", false, 4);
	var NumberFormatter=(function (_super) {
		__extends(NumberFormatter, _super);
		function NumberFormatter() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(NumberFormatter.prototype, "_className", {
			get: function () {
				return "NumberFormatter";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(NumberFormatter.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["isDateTime", "isPercent", "isCurrency", "isNumeric", "isText", "hasYear", "hasMonth", "hasDayOfWeek"];
			},
			enumerable: true,
			configurable: true
		});
		NumberFormatter.prototype.format=function (value) {
			return _invokeMethod(this, "Format", 0, [value], 0, 0);
		};
		NumberFormatter.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		NumberFormatter.prototype.toJSON=function () {
			return {};
		};
		return NumberFormatter;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.NumberFormatter=NumberFormatter;
	var CustomFunctionManager=(function (_super) {
		__extends(CustomFunctionManager, _super);
		function CustomFunctionManager() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(CustomFunctionManager.prototype, "_className", {
			get: function () {
				return "CustomFunctionManager";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(CustomFunctionManager.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["status"];
			},
			enumerable: true,
			configurable: true
		});
		CustomFunctionManager.prototype.register=function (metadata, javascript) {
			return _invokeMethod(this, "Register", 0, [metadata, javascript], 0);
		};
		CustomFunctionManager.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		CustomFunctionManager.prototype.toJSON=function () {
			return {};
		};
		return CustomFunctionManager;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.CustomFunctionManager=CustomFunctionManager;
	ExcelOp.customFunctionManager=_createTopLevelServiceObject(CustomFunctionManager, _localDocumentContext, "Microsoft.ExcelServices.CustomFunctionManager", false, 4);
	var Style=(function (_super) {
		__extends(Style, _super);
		function Style() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(Style.prototype, "_className", {
			get: function () {
				return "Style";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Style.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["builtIn", "formulaHidden", "horizontalAlignment", "includeAlignment", "includeBorder", "includeFont", "includeNumber", "includePatterns", "includeProtection", "indentLevel", "locked", "name", "numberFormat", "numberFormatLocal", "readingOrder", "shrinkToFit", "verticalAlignment", "wrapText", "textOrientation", "autoIndent"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Style.prototype, "_scalarPropertyUpdateable", {
			get: function () {
				return [false, true, true, true, true, true, true, true, true, true, true, false, true, true, true, true, true, true, true, true];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Style.prototype, "_navigationPropertyNames", {
			get: function () {
				return ["borders", "font", "fill"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Style.prototype, "borders", {
			get: function () {
				return _createPropertyObject(ExcelOp.RangeBorderCollection, this, "Borders", true, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Style.prototype, "fill", {
			get: function () {
				return _createPropertyObject(ExcelOp.RangeFill, this, "Fill", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Style.prototype, "font", {
			get: function () {
				return _createPropertyObject(ExcelOp.RangeFont, this, "Font", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Style.prototype.update=function (properties) {
			return _invokeRecursiveUpdate(this, properties);
		};
		Style.prototype["delete"]=function () {
			return _invokeMethod(this, "Delete", 0, [], 0);
		};
		Style.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		Style.prototype.toJSON=function () {
			return {};
		};
		return Style;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.Style=Style;
	var StyleCollection=(function (_super) {
		__extends(StyleCollection, _super);
		function StyleCollection() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(StyleCollection.prototype, "_className", {
			get: function () {
				return "StyleCollection";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(StyleCollection.prototype, "_isCollection", {
			get: function () {
				return true;
			},
			enumerable: true,
			configurable: true
		});
		StyleCollection.prototype.getItem=function (name) {
			return _createIndexerObject(ExcelOp.Style, this, [name]);
		};
		StyleCollection.prototype.getItemAt=function (index) {
			_throwIfApiNotSupported("StyleCollection.getItemAt", _defaultApiSetName, "1.9", _hostName);
			return _createMethodObject(ExcelOp.Style, this, "GetItemAt", 1, [index], false, false, null, 4);
		};
		StyleCollection.prototype.add=function (name) {
			return _invokeMethod(this, "Add", 0, [name], 0);
		};
		StyleCollection.prototype.getCount=function () {
			_throwIfApiNotSupported("StyleCollection.getCount", _defaultApiSetName, "1.9", _hostName);
			return _invokeMethod(this, "GetCount", 1, [], 4, 0);
		};
		StyleCollection.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		StyleCollection.prototype.toJSON=function () {
			return {};
		};
		return StyleCollection;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.StyleCollection=StyleCollection;
	var InternalTest=(function (_super) {
		__extends(InternalTest, _super);
		function InternalTest() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(InternalTest.prototype, "_className", {
			get: function () {
				return "InternalTest";
			},
			enumerable: true,
			configurable: true
		});
		InternalTest.prototype.compareTempFilesAreIdentical=function (filename1, filename2) {
			_throwIfApiNotSupported("InternalTest.compareTempFilesAreIdentical", _defaultApiSetName, "99.9", _hostName);
			return _invokeMethod(this, "CompareTempFilesAreIdentical", 1, [filename1, filename2], 0, 0);
		};
		InternalTest.prototype.delay=function (seconds) {
			return _invokeMethod(this, "Delay", 0, [seconds], 0, 0);
		};
		InternalTest.prototype.enterCellEdit=function (duration) {
			_throwIfApiNotSupported("InternalTest.enterCellEdit", _defaultApiSetName, "1.9", _hostName);
			return _invokeMethod(this, "EnterCellEdit", 0, [duration], 0);
		};
		InternalTest.prototype.firstPartyMethod=function () {
			_throwIfApiNotSupported("InternalTest.firstPartyMethod", _defaultApiSetName, "1.7", _hostName);
			return _invokeMethod(this, "FirstPartyMethod", 1, [], 4 | 1);
		};
		InternalTest.prototype.installCustomFunctionsFromCache=function () {
			_throwIfApiNotSupported("InternalTest.installCustomFunctionsFromCache", _defaultApiSetName, "1.9", _hostName);
			return _invokeMethod(this, "InstallCustomFunctionsFromCache", 0, [], 0);
		};
		InternalTest.prototype.recalc=function (force, allFormulas) {
			_throwIfApiNotSupported("InternalTest.recalc", _defaultApiSetName, "1.9", _hostName);
			return _invokeMethod(this, "Recalc", 0, [force, allFormulas], 0);
		};
		InternalTest.prototype.recalcBySolutionId=function (solutionId) {
			_throwIfApiNotSupported("InternalTest.recalcBySolutionId", _defaultApiSetName, "1.9", _hostName);
			return _invokeMethod(this, "RecalcBySolutionId", 0, [solutionId], 0);
		};
		InternalTest.prototype.saveWorkbookToTempFile=function (filename) {
			_throwIfApiNotSupported("InternalTest.saveWorkbookToTempFile", _defaultApiSetName, "99.9", _hostName);
			return _invokeMethod(this, "SaveWorkbookToTempFile", 1, [filename], 0);
		};
		InternalTest.prototype.triggerMessage=function (messageCategory, messageType, targetId, message) {
			_throwIfApiNotSupported("InternalTest.triggerMessage", _defaultApiSetName, "1.7", _hostName);
			return _invokeMethod(this, "TriggerMessage", 0, [messageCategory, messageType, targetId, message], 0);
		};
		InternalTest.prototype.triggerPostProcess=function () {
			_throwIfApiNotSupported("InternalTest.triggerPostProcess", _defaultApiSetName, "1.7", _hostName);
			return _invokeMethod(this, "TriggerPostProcess", 0, [], 0);
		};
		InternalTest.prototype.triggerTestEvent=function (prop1, worksheet) {
			_throwIfApiNotSupported("InternalTest.triggerTestEvent", _defaultApiSetName, "1.7", _hostName);
			return _invokeMethod(this, "TriggerTestEvent", 0, [prop1, worksheet], 0);
		};
		InternalTest.prototype.triggerTestEventWithFilter=function (prop1, msgType, worksheet) {
			_throwIfApiNotSupported("InternalTest.triggerTestEventWithFilter", _defaultApiSetName, "1.7", _hostName);
			return _invokeMethod(this, "TriggerTestEventWithFilter", 0, [prop1, msgType, worksheet], 0);
		};
		InternalTest.prototype.triggerUserRedo=function () {
			_throwIfApiNotSupported("InternalTest.triggerUserRedo", _defaultApiSetName, "99.9", _hostName);
			return _invokeMethod(this, "TriggerUserRedo", 1, [], 0);
		};
		InternalTest.prototype.triggerUserUndo=function () {
			_throwIfApiNotSupported("InternalTest.triggerUserUndo", _defaultApiSetName, "99.9", _hostName);
			return _invokeMethod(this, "TriggerUserUndo", 1, [], 0);
		};
		InternalTest.prototype.unregisterAllCustomFunctionExecutionEvents=function () {
			_throwIfApiNotSupported("InternalTest.unregisterAllCustomFunctionExecutionEvents", "CustomFunctions", "1.1", _hostName);
			return _invokeMethod(this, "UnregisterAllCustomFunctionExecutionEvents", 0, [], 0);
		};
		InternalTest.prototype.updateRangeValueOnCurrentSheet=function (address, values) {
			_throwIfApiNotSupported("InternalTest.updateRangeValueOnCurrentSheet", _defaultApiSetName, "99.9", _hostName);
			return _invokeMethod(this, "UpdateRangeValueOnCurrentSheet", 0, [address, values], 2);
		};
		InternalTest.prototype.toJSON=function () {
			return {};
		};
		return InternalTest;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.InternalTest=InternalTest;
	var PageLayout=(function (_super) {
		__extends(PageLayout, _super);
		function PageLayout() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(PageLayout.prototype, "_className", {
			get: function () {
				return "PageLayout";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(PageLayout.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["orientation", "paperSize", "blackAndWhite", "printErrors", "zoom", "centerHorizontally", "centerVertically", "printHeadings", "printGridlines", "leftMargin", "rightMargin", "topMargin", "bottomMargin", "headerMargin", "footerMargin", "printComments", "draftMode", "firstPageNumber", "printOrder"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(PageLayout.prototype, "_scalarPropertyUpdateable", {
			get: function () {
				return [true, true, true, true, true, true, true, true, true, true, true, true, true, true, true, true, true, true, true];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(PageLayout.prototype, "_navigationPropertyNames", {
			get: function () {
				return ["headersFooters"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(PageLayout.prototype, "headersFooters", {
			get: function () {
				return _createPropertyObject(ExcelOp.HeaderFooterGroup, this, "HeadersFooters", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		PageLayout.prototype.getPrintArea=function () {
			return _createMethodObject(ExcelOp.RangeAreas, this, "GetPrintArea", 1, [], false, true, null, 4);
		};
		PageLayout.prototype.getPrintAreaOrNullObject=function () {
			return _createMethodObject(ExcelOp.RangeAreas, this, "GetPrintAreaOrNullObject", 1, [], false, true, null, 4);
		};
		PageLayout.prototype.getPrintTitleColumns=function () {
			return _createMethodObject(ExcelOp.Range, this, "GetPrintTitleColumns", 1, [], false, true, null, 4);
		};
		PageLayout.prototype.getPrintTitleColumnsOrNullObject=function () {
			return _createMethodObject(ExcelOp.Range, this, "GetPrintTitleColumnsOrNullObject", 1, [], false, true, null, 4);
		};
		PageLayout.prototype.getPrintTitleRows=function () {
			return _createMethodObject(ExcelOp.Range, this, "GetPrintTitleRows", 1, [], false, true, null, 4);
		};
		PageLayout.prototype.getPrintTitleRowsOrNullObject=function () {
			return _createMethodObject(ExcelOp.Range, this, "GetPrintTitleRowsOrNullObject", 1, [], false, true, null, 4);
		};
		PageLayout.prototype.update=function (properties) {
			return _invokeRecursiveUpdate(this, properties);
		};
		PageLayout.prototype.setPrintArea=function (printArea) {
			return _invokeMethod(this, "SetPrintArea", 0, [printArea], 0);
		};
		PageLayout.prototype.setPrintMargins=function (unit, marginOptions) {
			return _invokeMethod(this, "SetPrintMargins", 0, [unit, marginOptions], 0);
		};
		PageLayout.prototype.setPrintTitleColumns=function (printTitleColumns) {
			return _invokeMethod(this, "SetPrintTitleColumns", 0, [printTitleColumns], 0);
		};
		PageLayout.prototype.setPrintTitleRows=function (printTitleRows) {
			return _invokeMethod(this, "SetPrintTitleRows", 0, [printTitleRows], 0);
		};
		PageLayout.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		PageLayout.prototype.toJSON=function () {
			return {};
		};
		return PageLayout;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.PageLayout=PageLayout;
	var HeaderFooter=(function (_super) {
		__extends(HeaderFooter, _super);
		function HeaderFooter() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(HeaderFooter.prototype, "_className", {
			get: function () {
				return "HeaderFooter";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(HeaderFooter.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["leftHeader", "centerHeader", "rightHeader", "leftFooter", "centerFooter", "rightFooter"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(HeaderFooter.prototype, "_scalarPropertyUpdateable", {
			get: function () {
				return [true, true, true, true, true, true];
			},
			enumerable: true,
			configurable: true
		});
		HeaderFooter.prototype.update=function (properties) {
			return _invokeRecursiveUpdate(this, properties);
		};
		HeaderFooter.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		HeaderFooter.prototype.toJSON=function () {
			return {};
		};
		return HeaderFooter;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.HeaderFooter=HeaderFooter;
	var HeaderFooterGroup=(function (_super) {
		__extends(HeaderFooterGroup, _super);
		function HeaderFooterGroup() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(HeaderFooterGroup.prototype, "_className", {
			get: function () {
				return "HeaderFooterGroup";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(HeaderFooterGroup.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["state", "useSheetMargins", "useSheetScale"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(HeaderFooterGroup.prototype, "_scalarPropertyUpdateable", {
			get: function () {
				return [true, true, true];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(HeaderFooterGroup.prototype, "_navigationPropertyNames", {
			get: function () {
				return ["defaultForAllPages", "firstPage", "evenPages", "oddPages"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(HeaderFooterGroup.prototype, "defaultForAllPages", {
			get: function () {
				return _createPropertyObject(ExcelOp.HeaderFooter, this, "DefaultForAllPages", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(HeaderFooterGroup.prototype, "evenPages", {
			get: function () {
				return _createPropertyObject(ExcelOp.HeaderFooter, this, "EvenPages", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(HeaderFooterGroup.prototype, "firstPage", {
			get: function () {
				return _createPropertyObject(ExcelOp.HeaderFooter, this, "FirstPage", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(HeaderFooterGroup.prototype, "oddPages", {
			get: function () {
				return _createPropertyObject(ExcelOp.HeaderFooter, this, "OddPages", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		HeaderFooterGroup.prototype.update=function (properties) {
			return _invokeRecursiveUpdate(this, properties);
		};
		HeaderFooterGroup.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		HeaderFooterGroup.prototype.toJSON=function () {
			return {};
		};
		return HeaderFooterGroup;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.HeaderFooterGroup=HeaderFooterGroup;
	var PageBreak=(function (_super) {
		__extends(PageBreak, _super);
		function PageBreak() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(PageBreak.prototype, "_className", {
			get: function () {
				return "PageBreak";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(PageBreak.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["_Id", "columnIndex", "rowIndex"];
			},
			enumerable: true,
			configurable: true
		});
		PageBreak.prototype.getStartCell=function () {
			return _createMethodObject(ExcelOp.Range, this, "GetStartCell", 1, [], false, true, null, 4);
		};
		PageBreak.prototype["delete"]=function () {
			return _invokeMethod(this, "Delete", 0, [], 0);
		};
		PageBreak.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		PageBreak.prototype.toJSON=function () {
			return {};
		};
		return PageBreak;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.PageBreak=PageBreak;
	var PageBreakCollection=(function (_super) {
		__extends(PageBreakCollection, _super);
		function PageBreakCollection() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(PageBreakCollection.prototype, "_className", {
			get: function () {
				return "PageBreakCollection";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(PageBreakCollection.prototype, "_isCollection", {
			get: function () {
				return true;
			},
			enumerable: true,
			configurable: true
		});
		PageBreakCollection.prototype.add=function (pageBreakRange) {
			return _createAndInstantiateMethodObject(ExcelOp.PageBreak, this, "Add", 0, [pageBreakRange], false, true, null, 0);
		};
		PageBreakCollection.prototype.getItem=function (index) {
			return _createIndexerObject(ExcelOp.PageBreak, this, [index]);
		};
		PageBreakCollection.prototype.getCount=function () {
			return _invokeMethod(this, "GetCount", 1, [], 4, 0);
		};
		PageBreakCollection.prototype.removePageBreaks=function () {
			return _invokeMethod(this, "RemovePageBreaks", 0, [], 0);
		};
		PageBreakCollection.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		PageBreakCollection.prototype.toJSON=function () {
			return {};
		};
		return PageBreakCollection;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.PageBreakCollection=PageBreakCollection;
	var DataConnectionCollection=(function (_super) {
		__extends(DataConnectionCollection, _super);
		function DataConnectionCollection() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(DataConnectionCollection.prototype, "_className", {
			get: function () {
				return "DataConnectionCollection";
			},
			enumerable: true,
			configurable: true
		});
		DataConnectionCollection.prototype.refreshAll=function () {
			return _invokeMethod(this, "RefreshAll", 0, [], 0);
		};
		DataConnectionCollection.prototype.toJSON=function () {
			return {};
		};
		return DataConnectionCollection;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.DataConnectionCollection=DataConnectionCollection;
	var RangeCollection=(function (_super) {
		__extends(RangeCollection, _super);
		function RangeCollection() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(RangeCollection.prototype, "_className", {
			get: function () {
				return "RangeCollection";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(RangeCollection.prototype, "_isCollection", {
			get: function () {
				return true;
			},
			enumerable: true,
			configurable: true
		});
		RangeCollection.prototype.getItemAt=function (index) {
			return _createMethodObject(ExcelOp.Range, this, "GetItemAt", 1, [index], false, false, null, 4);
		};
		RangeCollection.prototype.getCount=function () {
			return _invokeMethod(this, "GetCount", 1, [], 4, 0);
		};
		RangeCollection.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		RangeCollection.prototype.toJSON=function () {
			return {};
		};
		return RangeCollection;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.RangeCollection=RangeCollection;
	var CommentCollection=(function (_super) {
		__extends(CommentCollection, _super);
		function CommentCollection() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(CommentCollection.prototype, "_className", {
			get: function () {
				return "CommentCollection";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(CommentCollection.prototype, "_isCollection", {
			get: function () {
				return true;
			},
			enumerable: true,
			configurable: true
		});
		CommentCollection.prototype.add=function (content, cellAddress, contentType) {
			return _createAndInstantiateMethodObject(ExcelOp.Comment, this, "Add", 0, [content, cellAddress, contentType], false, true, null, 0);
		};
		CommentCollection.prototype.getItem=function (commentId) {
			return _createIndexerObject(ExcelOp.Comment, this, [commentId]);
		};
		CommentCollection.prototype.getItemAt=function (index) {
			return _createMethodObject(ExcelOp.Comment, this, "GetItemAt", 1, [index], false, false, null, 4);
		};
		CommentCollection.prototype.getItemByCell=function (cellAddress) {
			return _createMethodObject(ExcelOp.Comment, this, "GetItemByCell", 1, [cellAddress], false, false, null, 4);
		};
		CommentCollection.prototype.getItemByReplyId=function (replyId) {
			return _createMethodObject(ExcelOp.Comment, this, "GetItemByReplyId", 1, [replyId], false, false, null, 4);
		};
		CommentCollection.prototype.getCount=function () {
			return _invokeMethod(this, "GetCount", 1, [], 4, 0);
		};
		CommentCollection.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		CommentCollection.prototype.toJSON=function () {
			return {};
		};
		return CommentCollection;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.CommentCollection=CommentCollection;
	var Comment=(function (_super) {
		__extends(Comment, _super);
		function Comment() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(Comment.prototype, "_className", {
			get: function () {
				return "Comment";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Comment.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["id", "isParent", "content", "authorName", "authorEmail", "creationDate"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Comment.prototype, "_scalarPropertyUpdateable", {
			get: function () {
				return [false, false, true, false, false, false];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Comment.prototype, "_navigationPropertyNames", {
			get: function () {
				return ["replies"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Comment.prototype, "replies", {
			get: function () {
				return _createPropertyObject(ExcelOp.CommentReplyCollection, this, "Replies", true, 4);
			},
			enumerable: true,
			configurable: true
		});
		Comment.prototype.getLocation=function () {
			return _createMethodObject(ExcelOp.Range, this, "GetLocation", 1, [], false, true, null, 4);
		};
		Comment.prototype.update=function (properties) {
			return _invokeRecursiveUpdate(this, properties);
		};
		Comment.prototype["delete"]=function () {
			return _invokeMethod(this, "Delete", 0, [], 0);
		};
		Comment.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		Comment.prototype.toJSON=function () {
			return {};
		};
		return Comment;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.Comment=Comment;
	var CommentReplyCollection=(function (_super) {
		__extends(CommentReplyCollection, _super);
		function CommentReplyCollection() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(CommentReplyCollection.prototype, "_className", {
			get: function () {
				return "CommentReplyCollection";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(CommentReplyCollection.prototype, "_isCollection", {
			get: function () {
				return true;
			},
			enumerable: true,
			configurable: true
		});
		CommentReplyCollection.prototype.add=function (content, contentType) {
			return _createAndInstantiateMethodObject(ExcelOp.CommentReply, this, "Add", 0, [content, contentType], false, true, null, 0);
		};
		CommentReplyCollection.prototype.getItem=function (commentReplyId) {
			return _createIndexerObject(ExcelOp.CommentReply, this, [commentReplyId]);
		};
		CommentReplyCollection.prototype.getItemAt=function (index) {
			return _createMethodObject(ExcelOp.CommentReply, this, "GetItemAt", 1, [index], false, false, null, 4);
		};
		CommentReplyCollection.prototype.getCount=function () {
			return _invokeMethod(this, "GetCount", 1, [], 4, 0);
		};
		CommentReplyCollection.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		CommentReplyCollection.prototype.toJSON=function () {
			return {};
		};
		return CommentReplyCollection;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.CommentReplyCollection=CommentReplyCollection;
	var CommentReply=(function (_super) {
		__extends(CommentReply, _super);
		function CommentReply() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(CommentReply.prototype, "_className", {
			get: function () {
				return "CommentReply";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(CommentReply.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["id", "isParent", "content", "authorName", "authorEmail", "creationDate"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(CommentReply.prototype, "_scalarPropertyUpdateable", {
			get: function () {
				return [false, false, true, false, false, false];
			},
			enumerable: true,
			configurable: true
		});
		CommentReply.prototype.getLocation=function () {
			return _createMethodObject(ExcelOp.Range, this, "GetLocation", 1, [], false, true, null, 4);
		};
		CommentReply.prototype.getParentComment=function () {
			return _createAndInstantiateMethodObject(ExcelOp.Comment, this, "GetParentComment", 0, [], false, false, null, 0);
		};
		CommentReply.prototype.update=function (properties) {
			return _invokeRecursiveUpdate(this, properties);
		};
		CommentReply.prototype["delete"]=function () {
			return _invokeMethod(this, "Delete", 0, [], 0);
		};
		CommentReply.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		CommentReply.prototype.toJSON=function () {
			return {};
		};
		return CommentReply;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.CommentReply=CommentReply;
	var ShapeCollection=(function (_super) {
		__extends(ShapeCollection, _super);
		function ShapeCollection() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(ShapeCollection.prototype, "_className", {
			get: function () {
				return "ShapeCollection";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ShapeCollection.prototype, "_isCollection", {
			get: function () {
				return true;
			},
			enumerable: true,
			configurable: true
		});
		ShapeCollection.prototype.addGeometricShape=function (geometricShapeType) {
			return _createAndInstantiateMethodObject(ExcelOp.Shape, this, "AddGeometricShape", 0, [geometricShapeType], false, false, null, 0);
		};
		ShapeCollection.prototype.addGroup=function (values) {
			return _createAndInstantiateMethodObject(ExcelOp.Shape, this, "AddGroup", 0, [values], false, false, null, 0);
		};
		ShapeCollection.prototype.addImage=function (base64ImageString) {
			return _createAndInstantiateMethodObject(ExcelOp.Shape, this, "AddImage", 0, [base64ImageString], false, false, null, 0);
		};
		ShapeCollection.prototype.addLine=function (startLeft, startTop, endLeft, endTop, connectorType) {
			return _createAndInstantiateMethodObject(ExcelOp.Shape, this, "AddLine", 0, [startLeft, startTop, endLeft, endTop, connectorType], false, false, null, 0);
		};
		ShapeCollection.prototype.addSVG=function (xmlImageString) {
			return _createAndInstantiateMethodObject(ExcelOp.Shape, this, "AddSVG", 0, [xmlImageString], false, false, null, 0);
		};
		ShapeCollection.prototype.addTextBox=function (text) {
			return _createAndInstantiateMethodObject(ExcelOp.Shape, this, "AddTextBox", 0, [text], false, false, null, 0);
		};
		ShapeCollection.prototype.getItem=function (name) {
			return _createMethodObject(ExcelOp.Shape, this, "GetItem", 1, [name], false, false, null, 4);
		};
		ShapeCollection.prototype.getItemAt=function (index) {
			return _createMethodObject(ExcelOp.Shape, this, "GetItemAt", 1, [index], false, false, null, 4);
		};
		ShapeCollection.prototype._GetItem=function (shapeId) {
			return _createIndexerObject(ExcelOp.Shape, this, [shapeId]);
		};
		ShapeCollection.prototype.getCount=function () {
			return _invokeMethod(this, "GetCount", 1, [], 4, 0);
		};
		ShapeCollection.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		ShapeCollection.prototype.toJSON=function () {
			return {};
		};
		return ShapeCollection;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.ShapeCollection=ShapeCollection;
	var Shape=(function (_super) {
		__extends(Shape, _super);
		function Shape() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(Shape.prototype, "_className", {
			get: function () {
				return "Shape";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Shape.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["id", "name", "left", "top", "width", "height", "rotation", "zorderPosition", "altTextTitle", "altTextDescription", "type", "lockAspectRatio", "placement", "geometricShapeType", "visible", "level", "connectionSiteCount"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Shape.prototype, "_scalarPropertyUpdateable", {
			get: function () {
				return [false, true, true, true, true, true, true, false, true, true, false, true, true, true, true, false, false];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Shape.prototype, "_navigationPropertyNames", {
			get: function () {
				return ["geometricShape", "image", "textFrame", "fill", "group", "parentGroup", "line", "lineFormat"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Shape.prototype, "fill", {
			get: function () {
				return _createPropertyObject(ExcelOp.ShapeFill, this, "Fill", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Shape.prototype, "geometricShape", {
			get: function () {
				return _createPropertyObject(ExcelOp.GeometricShape, this, "GeometricShape", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Shape.prototype, "group", {
			get: function () {
				return _createPropertyObject(ExcelOp.ShapeGroup, this, "Group", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Shape.prototype, "image", {
			get: function () {
				return _createPropertyObject(ExcelOp.Image, this, "Image", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Shape.prototype, "line", {
			get: function () {
				return _createPropertyObject(ExcelOp.Line, this, "Line", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Shape.prototype, "lineFormat", {
			get: function () {
				return _createPropertyObject(ExcelOp.ShapeLineFormat, this, "LineFormat", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Shape.prototype, "parentGroup", {
			get: function () {
				return _createPropertyObject(ExcelOp.Shape, this, "ParentGroup", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Shape.prototype, "textFrame", {
			get: function () {
				return _createPropertyObject(ExcelOp.TextFrame, this, "TextFrame", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Shape.prototype.update=function (properties) {
			return _invokeRecursiveUpdate(this, properties);
		};
		Shape.prototype["delete"]=function () {
			return _invokeMethod(this, "Delete", 0, [], 0);
		};
		Shape.prototype.incrementLeft=function (increment) {
			return _invokeMethod(this, "IncrementLeft", 0, [increment], 0);
		};
		Shape.prototype.incrementRotation=function (increment) {
			return _invokeMethod(this, "IncrementRotation", 0, [increment], 0);
		};
		Shape.prototype.incrementTop=function (increment) {
			return _invokeMethod(this, "IncrementTop", 0, [increment], 0);
		};
		Shape.prototype.saveAsPicture=function (format) {
			return _invokeMethod(this, "SaveAsPicture", 0, [format], 0, 0);
		};
		Shape.prototype.scaleHeight=function (scaleFactor, scaleType, scaleFrom) {
			return _invokeMethod(this, "ScaleHeight", 0, [scaleFactor, scaleType, scaleFrom], 0);
		};
		Shape.prototype.scaleWidth=function (scaleFactor, scaleType, scaleFrom) {
			return _invokeMethod(this, "ScaleWidth", 0, [scaleFactor, scaleType, scaleFrom], 0);
		};
		Shape.prototype.setZOrder=function (value) {
			return _invokeMethod(this, "SetZOrder", 0, [value], 0);
		};
		Shape.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		Shape.prototype.toJSON=function () {
			return {};
		};
		return Shape;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.Shape=Shape;
	var GeometricShape=(function (_super) {
		__extends(GeometricShape, _super);
		function GeometricShape() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(GeometricShape.prototype, "_className", {
			get: function () {
				return "GeometricShape";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(GeometricShape.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["id"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(GeometricShape.prototype, "_navigationPropertyNames", {
			get: function () {
				return ["shape"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(GeometricShape.prototype, "shape", {
			get: function () {
				return _createPropertyObject(ExcelOp.Shape, this, "Shape", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		GeometricShape.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		GeometricShape.prototype.toJSON=function () {
			return {};
		};
		return GeometricShape;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.GeometricShape=GeometricShape;
	var Image=(function (_super) {
		__extends(Image, _super);
		function Image() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(Image.prototype, "_className", {
			get: function () {
				return "Image";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Image.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["id", "format"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Image.prototype, "_navigationPropertyNames", {
			get: function () {
				return ["shape"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Image.prototype, "shape", {
			get: function () {
				return _createPropertyObject(ExcelOp.Shape, this, "Shape", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Image.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		Image.prototype.toJSON=function () {
			return {};
		};
		return Image;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.Image=Image;
	var ShapeGroup=(function (_super) {
		__extends(ShapeGroup, _super);
		function ShapeGroup() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(ShapeGroup.prototype, "_className", {
			get: function () {
				return "ShapeGroup";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ShapeGroup.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["id"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ShapeGroup.prototype, "_navigationPropertyNames", {
			get: function () {
				return ["shapes", "shape"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ShapeGroup.prototype, "shape", {
			get: function () {
				return _createPropertyObject(ExcelOp.Shape, this, "Shape", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ShapeGroup.prototype, "shapes", {
			get: function () {
				return _createPropertyObject(ExcelOp.GroupShapeCollection, this, "Shapes", true, 4);
			},
			enumerable: true,
			configurable: true
		});
		ShapeGroup.prototype.ungroup=function () {
			return _invokeMethod(this, "Ungroup", 0, [], 0);
		};
		ShapeGroup.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		ShapeGroup.prototype.toJSON=function () {
			return {};
		};
		return ShapeGroup;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.ShapeGroup=ShapeGroup;
	var GroupShapeCollection=(function (_super) {
		__extends(GroupShapeCollection, _super);
		function GroupShapeCollection() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(GroupShapeCollection.prototype, "_className", {
			get: function () {
				return "GroupShapeCollection";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(GroupShapeCollection.prototype, "_isCollection", {
			get: function () {
				return true;
			},
			enumerable: true,
			configurable: true
		});
		GroupShapeCollection.prototype.getItem=function (name) {
			return _createMethodObject(ExcelOp.Shape, this, "GetItem", 1, [name], false, false, null, 4);
		};
		GroupShapeCollection.prototype.getItemAt=function (index) {
			return _createMethodObject(ExcelOp.Shape, this, "GetItemAt", 1, [index], false, false, null, 4);
		};
		GroupShapeCollection.prototype._GetItem=function (shapeId) {
			return _createIndexerObject(ExcelOp.Shape, this, [shapeId]);
		};
		GroupShapeCollection.prototype.getCount=function () {
			return _invokeMethod(this, "GetCount", 1, [], 4, 0);
		};
		GroupShapeCollection.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		GroupShapeCollection.prototype.toJSON=function () {
			return {};
		};
		return GroupShapeCollection;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.GroupShapeCollection=GroupShapeCollection;
	var Line=(function (_super) {
		__extends(Line, _super);
		function Line() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(Line.prototype, "_className", {
			get: function () {
				return "Line";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Line.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["id", "connectorType", "beginArrowHeadLength", "beginArrowHeadStyle", "beginArrowHeadWidth", "endArrowHeadLength", "endArrowHeadStyle", "endArrowHeadWidth", "isBeginConnected", "beginConnectedSite", "isEndConnected", "endConnectedSite"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Line.prototype, "_scalarPropertyUpdateable", {
			get: function () {
				return [false, true, true, true, true, true, true, true, false, false, false, false];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Line.prototype, "_navigationPropertyNames", {
			get: function () {
				return ["shape", "beginConnectedShape", "endConnectedShape"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Line.prototype, "beginConnectedShape", {
			get: function () {
				return _createPropertyObject(ExcelOp.Shape, this, "BeginConnectedShape", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Line.prototype, "endConnectedShape", {
			get: function () {
				return _createPropertyObject(ExcelOp.Shape, this, "EndConnectedShape", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Line.prototype, "shape", {
			get: function () {
				return _createPropertyObject(ExcelOp.Shape, this, "Shape", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Line.prototype.update=function (properties) {
			return _invokeRecursiveUpdate(this, properties);
		};
		Line.prototype.beginConnect=function (shape, connectionSite) {
			return _invokeMethod(this, "BeginConnect", 0, [shape, connectionSite], 0);
		};
		Line.prototype.beginDisconnect=function () {
			return _invokeMethod(this, "BeginDisconnect", 0, [], 0);
		};
		Line.prototype.endConnect=function (shape, connectionSite) {
			return _invokeMethod(this, "EndConnect", 0, [shape, connectionSite], 0);
		};
		Line.prototype.endDisconnect=function () {
			return _invokeMethod(this, "EndDisconnect", 0, [], 0);
		};
		Line.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		Line.prototype.toJSON=function () {
			return {};
		};
		return Line;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.Line=Line;
	var ShapeFill=(function (_super) {
		__extends(ShapeFill, _super);
		function ShapeFill() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(ShapeFill.prototype, "_className", {
			get: function () {
				return "ShapeFill";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ShapeFill.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["foreColor", "type", "transparency"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ShapeFill.prototype, "_scalarPropertyUpdateable", {
			get: function () {
				return [true, false, true];
			},
			enumerable: true,
			configurable: true
		});
		ShapeFill.prototype.update=function (properties) {
			return _invokeRecursiveUpdate(this, properties);
		};
		ShapeFill.prototype.clear=function () {
			return _invokeMethod(this, "Clear", 0, [], 0);
		};
		ShapeFill.prototype.setSolidColor=function (color) {
			return _invokeMethod(this, "SetSolidColor", 0, [color], 0);
		};
		ShapeFill.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		ShapeFill.prototype.toJSON=function () {
			return {};
		};
		return ShapeFill;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.ShapeFill=ShapeFill;
	var ShapeLineFormat=(function (_super) {
		__extends(ShapeLineFormat, _super);
		function ShapeLineFormat() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(ShapeLineFormat.prototype, "_className", {
			get: function () {
				return "ShapeLineFormat";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ShapeLineFormat.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["visible", "color", "style", "weight", "dashStyle", "transparency"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ShapeLineFormat.prototype, "_scalarPropertyUpdateable", {
			get: function () {
				return [true, true, true, true, true, true];
			},
			enumerable: true,
			configurable: true
		});
		ShapeLineFormat.prototype.update=function (properties) {
			return _invokeRecursiveUpdate(this, properties);
		};
		ShapeLineFormat.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		ShapeLineFormat.prototype.toJSON=function () {
			return {};
		};
		return ShapeLineFormat;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.ShapeLineFormat=ShapeLineFormat;
	var TextFrame=(function (_super) {
		__extends(TextFrame, _super);
		function TextFrame() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(TextFrame.prototype, "_className", {
			get: function () {
				return "TextFrame";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TextFrame.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["leftMargin", "rightMargin", "topMargin", "bottomMargin", "horizontalAlignment", "horizontalOverflow", "verticalAlignment", "verticalOverflow", "orientation", "readingOrder", "hasText", "autoSize"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TextFrame.prototype, "_scalarPropertyUpdateable", {
			get: function () {
				return [true, true, true, true, true, true, true, true, true, true, false, true];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TextFrame.prototype, "_navigationPropertyNames", {
			get: function () {
				return ["textRange"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TextFrame.prototype, "textRange", {
			get: function () {
				return _createPropertyObject(ExcelOp.TextRange, this, "TextRange", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		TextFrame.prototype.update=function (properties) {
			return _invokeRecursiveUpdate(this, properties);
		};
		TextFrame.prototype.deleteText=function () {
			return _invokeMethod(this, "DeleteText", 0, [], 0);
		};
		TextFrame.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		TextFrame.prototype.toJSON=function () {
			return {};
		};
		return TextFrame;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.TextFrame=TextFrame;
	var TextRange=(function (_super) {
		__extends(TextRange, _super);
		function TextRange() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(TextRange.prototype, "_className", {
			get: function () {
				return "TextRange";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TextRange.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["text"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TextRange.prototype, "_scalarPropertyUpdateable", {
			get: function () {
				return [true];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TextRange.prototype, "_navigationPropertyNames", {
			get: function () {
				return ["font"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TextRange.prototype, "font", {
			get: function () {
				return _createPropertyObject(ExcelOp.ShapeFont, this, "Font", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		TextRange.prototype.getCharacters=function (start, length) {
			return _createAndInstantiateMethodObject(ExcelOp.TextRange, this, "GetCharacters", 0, [start, length], false, false, null, 0);
		};
		TextRange.prototype.update=function (properties) {
			return _invokeRecursiveUpdate(this, properties);
		};
		TextRange.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		TextRange.prototype.toJSON=function () {
			return {};
		};
		return TextRange;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.TextRange=TextRange;
	var ShapeFont=(function (_super) {
		__extends(ShapeFont, _super);
		function ShapeFont() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(ShapeFont.prototype, "_className", {
			get: function () {
				return "ShapeFont";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ShapeFont.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["size", "name", "color", "bold", "italic", "underline"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ShapeFont.prototype, "_scalarPropertyUpdateable", {
			get: function () {
				return [true, true, true, true, true, true];
			},
			enumerable: true,
			configurable: true
		});
		ShapeFont.prototype.update=function (properties) {
			return _invokeRecursiveUpdate(this, properties);
		};
		ShapeFont.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		ShapeFont.prototype.toJSON=function () {
			return {};
		};
		return ShapeFont;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.ShapeFont=ShapeFont;
	var Slicer=(function (_super) {
		__extends(Slicer, _super);
		function Slicer() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(Slicer.prototype, "_className", {
			get: function () {
				return "Slicer";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Slicer.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["id", "name", "caption", "left", "top", "width", "height", "nameInFormula", "sourceFieldName", "isFilterCleared", "style", "columnWidth", "sortBy"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Slicer.prototype, "_scalarPropertyUpdateable", {
			get: function () {
				return [false, true, true, true, true, true, true, true, false, false, true, true, true];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Slicer.prototype, "_navigationPropertyNames", {
			get: function () {
				return ["slicerItems", "worksheet"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Slicer.prototype, "slicerItems", {
			get: function () {
				return _createPropertyObject(ExcelOp.SlicerItemCollection, this, "SlicerItems", true, 4);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Slicer.prototype, "worksheet", {
			get: function () {
				return _createPropertyObject(ExcelOp.Worksheet, this, "Worksheet", false, 4);
			},
			enumerable: true,
			configurable: true
		});
		Slicer.prototype.update=function (properties) {
			return _invokeRecursiveUpdate(this, properties);
		};
		Slicer.prototype.activate=function () {
			_throwIfApiNotSupported("Slicer.activate", _defaultApiSetName, "99.9", _hostName);
			return _invokeMethod(this, "Activate", 1, [], 0);
		};
		Slicer.prototype.clearFilters=function () {
			return _invokeMethod(this, "ClearFilters", 0, [], 0);
		};
		Slicer.prototype["delete"]=function () {
			return _invokeMethod(this, "Delete", 0, [], 0);
		};
		Slicer.prototype.getSelectedItems=function () {
			return _invokeMethod(this, "GetSelectedItems", 0, [], 0, 0);
		};
		Slicer.prototype.selectItems=function (items) {
			return _invokeMethod(this, "SelectItems", 0, [items], 0);
		};
		Slicer.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		Slicer.prototype.toJSON=function () {
			return {};
		};
		return Slicer;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.Slicer=Slicer;
	var SlicerCollection=(function (_super) {
		__extends(SlicerCollection, _super);
		function SlicerCollection() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(SlicerCollection.prototype, "_className", {
			get: function () {
				return "SlicerCollection";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(SlicerCollection.prototype, "_isCollection", {
			get: function () {
				return true;
			},
			enumerable: true,
			configurable: true
		});
		SlicerCollection.prototype.add=function (slicerSource, sourceField, slicerDestination) {
			return _createAndInstantiateMethodObject(ExcelOp.Slicer, this, "Add", 0, [slicerSource, sourceField, slicerDestination], false, true, null, 0);
		};
		SlicerCollection.prototype.getItem=function (key) {
			return _createIndexerObject(ExcelOp.Slicer, this, [key]);
		};
		SlicerCollection.prototype.getItemAt=function (index) {
			return _createMethodObject(ExcelOp.Slicer, this, "GetItemAt", 1, [index], false, false, null, 4);
		};
		SlicerCollection.prototype.getItemOrNullObject=function (key) {
			return _createMethodObject(ExcelOp.Slicer, this, "GetItemOrNullObject", 1, [key], false, false, null, 4);
		};
		SlicerCollection.prototype.getCount=function () {
			return _invokeMethod(this, "GetCount", 1, [], 4, 0);
		};
		SlicerCollection.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		SlicerCollection.prototype.toJSON=function () {
			return {};
		};
		return SlicerCollection;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.SlicerCollection=SlicerCollection;
	var SlicerItem=(function (_super) {
		__extends(SlicerItem, _super);
		function SlicerItem() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(SlicerItem.prototype, "_className", {
			get: function () {
				return "SlicerItem";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(SlicerItem.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["key", "name", "isSelected", "hasData"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(SlicerItem.prototype, "_scalarPropertyUpdateable", {
			get: function () {
				return [false, false, true, false];
			},
			enumerable: true,
			configurable: true
		});
		SlicerItem.prototype.update=function (properties) {
			return _invokeRecursiveUpdate(this, properties);
		};
		SlicerItem.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		SlicerItem.prototype.toJSON=function () {
			return {};
		};
		return SlicerItem;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.SlicerItem=SlicerItem;
	var SlicerItemCollection=(function (_super) {
		__extends(SlicerItemCollection, _super);
		function SlicerItemCollection() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(SlicerItemCollection.prototype, "_className", {
			get: function () {
				return "SlicerItemCollection";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(SlicerItemCollection.prototype, "_isCollection", {
			get: function () {
				return true;
			},
			enumerable: true,
			configurable: true
		});
		SlicerItemCollection.prototype.getItem=function (key) {
			return _createIndexerObject(ExcelOp.SlicerItem, this, [key]);
		};
		SlicerItemCollection.prototype.getItemAt=function (index) {
			return _createMethodObject(ExcelOp.SlicerItem, this, "GetItemAt", 1, [index], false, false, null, 4);
		};
		SlicerItemCollection.prototype.getItemOrNullObject=function (key) {
			return _createMethodObject(ExcelOp.SlicerItem, this, "GetItemOrNullObject", 1, [key], false, false, null, 4);
		};
		SlicerItemCollection.prototype.getCount=function () {
			return _invokeMethod(this, "GetCount", 1, [], 4, 0);
		};
		SlicerItemCollection.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		SlicerItemCollection.prototype.toJSON=function () {
			return {};
		};
		return SlicerItemCollection;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.SlicerItemCollection=SlicerItemCollection;
	var Ribbon=(function (_super) {
		__extends(Ribbon, _super);
		function Ribbon() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(Ribbon.prototype, "_className", {
			get: function () {
				return "Ribbon";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Ribbon.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["activeTab"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Ribbon.prototype, "_scalarPropertyUpdateable", {
			get: function () {
				return [true];
			},
			enumerable: true,
			configurable: true
		});
		Ribbon.prototype.update=function (properties) {
			return _invokeRecursiveUpdate(this, properties);
		};
		Ribbon.prototype.executeCommand=function (tcid, mouseClick) {
			return _invokeMethod(this, "ExecuteCommand", 0, [tcid, mouseClick], 0);
		};
		Ribbon.prototype.showTeachingCallout=function (tcid, title, message) {
			return _invokeMethod(this, "ShowTeachingCallout", 0, [tcid, title, message], 0);
		};
		Ribbon.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		Ribbon.prototype.toJSON=function () {
			return {};
		};
		return Ribbon;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.Ribbon=Ribbon;
	var AxisType;
	(function (AxisType) {
		AxisType["invalid"]="Invalid";
		AxisType["category"]="Category";
		AxisType["value"]="Value";
		AxisType["series"]="Series";
	})(AxisType=ExcelOp.AxisType || (ExcelOp.AxisType={}));
	var AxisGroup;
	(function (AxisGroup) {
		AxisGroup["primary"]="Primary";
		AxisGroup["secondary"]="Secondary";
	})(AxisGroup=ExcelOp.AxisGroup || (ExcelOp.AxisGroup={}));
	var AxisScaleType;
	(function (AxisScaleType) {
		AxisScaleType["linear"]="Linear";
		AxisScaleType["logarithmic"]="Logarithmic";
	})(AxisScaleType=ExcelOp.AxisScaleType || (ExcelOp.AxisScaleType={}));
	var AxisCrosses;
	(function (AxisCrosses) {
		AxisCrosses["automatic"]="Automatic";
		AxisCrosses["maximum"]="Maximum";
		AxisCrosses["minimum"]="Minimum";
		AxisCrosses["custom"]="Custom";
	})(AxisCrosses=ExcelOp.AxisCrosses || (ExcelOp.AxisCrosses={}));
	var AxisTickMark;
	(function (AxisTickMark) {
		AxisTickMark["none"]="None";
		AxisTickMark["cross"]="Cross";
		AxisTickMark["inside"]="Inside";
		AxisTickMark["outside"]="Outside";
	})(AxisTickMark=ExcelOp.AxisTickMark || (ExcelOp.AxisTickMark={}));
	var AxisTickLabelPosition;
	(function (AxisTickLabelPosition) {
		AxisTickLabelPosition["nextToAxis"]="NextToAxis";
		AxisTickLabelPosition["high"]="High";
		AxisTickLabelPosition["low"]="Low";
		AxisTickLabelPosition["none"]="None";
	})(AxisTickLabelPosition=ExcelOp.AxisTickLabelPosition || (ExcelOp.AxisTickLabelPosition={}));
	var TrendlineType;
	(function (TrendlineType) {
		TrendlineType["linear"]="Linear";
		TrendlineType["exponential"]="Exponential";
		TrendlineType["logarithmic"]="Logarithmic";
		TrendlineType["movingAverage"]="MovingAverage";
		TrendlineType["polynomial"]="Polynomial";
		TrendlineType["power"]="Power";
	})(TrendlineType=ExcelOp.TrendlineType || (ExcelOp.TrendlineType={}));
	var ChartAxisType;
	(function (ChartAxisType) {
		ChartAxisType["invalid"]="Invalid";
		ChartAxisType["category"]="Category";
		ChartAxisType["value"]="Value";
		ChartAxisType["series"]="Series";
	})(ChartAxisType=ExcelOp.ChartAxisType || (ExcelOp.ChartAxisType={}));
	var ChartAxisGroup;
	(function (ChartAxisGroup) {
		ChartAxisGroup["primary"]="Primary";
		ChartAxisGroup["secondary"]="Secondary";
	})(ChartAxisGroup=ExcelOp.ChartAxisGroup || (ExcelOp.ChartAxisGroup={}));
	var ChartAxisScaleType;
	(function (ChartAxisScaleType) {
		ChartAxisScaleType["linear"]="Linear";
		ChartAxisScaleType["logarithmic"]="Logarithmic";
	})(ChartAxisScaleType=ExcelOp.ChartAxisScaleType || (ExcelOp.ChartAxisScaleType={}));
	var ChartAxisPosition;
	(function (ChartAxisPosition) {
		ChartAxisPosition["automatic"]="Automatic";
		ChartAxisPosition["maximum"]="Maximum";
		ChartAxisPosition["minimum"]="Minimum";
		ChartAxisPosition["custom"]="Custom";
	})(ChartAxisPosition=ExcelOp.ChartAxisPosition || (ExcelOp.ChartAxisPosition={}));
	var ChartAxisTickMark;
	(function (ChartAxisTickMark) {
		ChartAxisTickMark["none"]="None";
		ChartAxisTickMark["cross"]="Cross";
		ChartAxisTickMark["inside"]="Inside";
		ChartAxisTickMark["outside"]="Outside";
	})(ChartAxisTickMark=ExcelOp.ChartAxisTickMark || (ExcelOp.ChartAxisTickMark={}));
	var CalculationState;
	(function (CalculationState) {
		CalculationState["done"]="Done";
		CalculationState["calculating"]="Calculating";
		CalculationState["pending"]="Pending";
	})(CalculationState=ExcelOp.CalculationState || (ExcelOp.CalculationState={}));
	var ChartAxisTickLabelPosition;
	(function (ChartAxisTickLabelPosition) {
		ChartAxisTickLabelPosition["nextToAxis"]="NextToAxis";
		ChartAxisTickLabelPosition["high"]="High";
		ChartAxisTickLabelPosition["low"]="Low";
		ChartAxisTickLabelPosition["none"]="None";
	})(ChartAxisTickLabelPosition=ExcelOp.ChartAxisTickLabelPosition || (ExcelOp.ChartAxisTickLabelPosition={}));
	var ChartAxisDisplayUnit;
	(function (ChartAxisDisplayUnit) {
		ChartAxisDisplayUnit["none"]="None";
		ChartAxisDisplayUnit["hundreds"]="Hundreds";
		ChartAxisDisplayUnit["thousands"]="Thousands";
		ChartAxisDisplayUnit["tenThousands"]="TenThousands";
		ChartAxisDisplayUnit["hundredThousands"]="HundredThousands";
		ChartAxisDisplayUnit["millions"]="Millions";
		ChartAxisDisplayUnit["tenMillions"]="TenMillions";
		ChartAxisDisplayUnit["hundredMillions"]="HundredMillions";
		ChartAxisDisplayUnit["billions"]="Billions";
		ChartAxisDisplayUnit["trillions"]="Trillions";
		ChartAxisDisplayUnit["custom"]="Custom";
	})(ChartAxisDisplayUnit=ExcelOp.ChartAxisDisplayUnit || (ExcelOp.ChartAxisDisplayUnit={}));
	var ChartAxisTimeUnit;
	(function (ChartAxisTimeUnit) {
		ChartAxisTimeUnit["days"]="Days";
		ChartAxisTimeUnit["months"]="Months";
		ChartAxisTimeUnit["years"]="Years";
	})(ChartAxisTimeUnit=ExcelOp.ChartAxisTimeUnit || (ExcelOp.ChartAxisTimeUnit={}));
	var ChartBoxQuartileCalculation;
	(function (ChartBoxQuartileCalculation) {
		ChartBoxQuartileCalculation["inclusive"]="Inclusive";
		ChartBoxQuartileCalculation["exclusive"]="Exclusive";
	})(ChartBoxQuartileCalculation=ExcelOp.ChartBoxQuartileCalculation || (ExcelOp.ChartBoxQuartileCalculation={}));
	var ChartAxisCategoryType;
	(function (ChartAxisCategoryType) {
		ChartAxisCategoryType["automatic"]="Automatic";
		ChartAxisCategoryType["textAxis"]="TextAxis";
		ChartAxisCategoryType["dateAxis"]="DateAxis";
	})(ChartAxisCategoryType=ExcelOp.ChartAxisCategoryType || (ExcelOp.ChartAxisCategoryType={}));
	var ChartBinType;
	(function (ChartBinType) {
		ChartBinType["category"]="Category";
		ChartBinType["auto"]="Auto";
		ChartBinType["binWidth"]="BinWidth";
		ChartBinType["binCount"]="BinCount";
	})(ChartBinType=ExcelOp.ChartBinType || (ExcelOp.ChartBinType={}));
	var ChartLineStyle;
	(function (ChartLineStyle) {
		ChartLineStyle["none"]="None";
		ChartLineStyle["continuous"]="Continuous";
		ChartLineStyle["dash"]="Dash";
		ChartLineStyle["dashDot"]="DashDot";
		ChartLineStyle["dashDotDot"]="DashDotDot";
		ChartLineStyle["dot"]="Dot";
		ChartLineStyle["grey25"]="Grey25";
		ChartLineStyle["grey50"]="Grey50";
		ChartLineStyle["grey75"]="Grey75";
		ChartLineStyle["automatic"]="Automatic";
		ChartLineStyle["roundDot"]="RoundDot";
	})(ChartLineStyle=ExcelOp.ChartLineStyle || (ExcelOp.ChartLineStyle={}));
	var ChartDataLabelPosition;
	(function (ChartDataLabelPosition) {
		ChartDataLabelPosition["invalid"]="Invalid";
		ChartDataLabelPosition["none"]="None";
		ChartDataLabelPosition["center"]="Center";
		ChartDataLabelPosition["insideEnd"]="InsideEnd";
		ChartDataLabelPosition["insideBase"]="InsideBase";
		ChartDataLabelPosition["outsideEnd"]="OutsideEnd";
		ChartDataLabelPosition["left"]="Left";
		ChartDataLabelPosition["right"]="Right";
		ChartDataLabelPosition["top"]="Top";
		ChartDataLabelPosition["bottom"]="Bottom";
		ChartDataLabelPosition["bestFit"]="BestFit";
		ChartDataLabelPosition["callout"]="Callout";
	})(ChartDataLabelPosition=ExcelOp.ChartDataLabelPosition || (ExcelOp.ChartDataLabelPosition={}));
	var ChartErrorBarsInclude;
	(function (ChartErrorBarsInclude) {
		ChartErrorBarsInclude["both"]="Both";
		ChartErrorBarsInclude["minusValues"]="MinusValues";
		ChartErrorBarsInclude["plusValues"]="PlusValues";
	})(ChartErrorBarsInclude=ExcelOp.ChartErrorBarsInclude || (ExcelOp.ChartErrorBarsInclude={}));
	var ChartErrorBarsType;
	(function (ChartErrorBarsType) {
		ChartErrorBarsType["fixedValue"]="FixedValue";
		ChartErrorBarsType["percent"]="Percent";
		ChartErrorBarsType["stDev"]="StDev";
		ChartErrorBarsType["stError"]="StError";
		ChartErrorBarsType["custom"]="Custom";
	})(ChartErrorBarsType=ExcelOp.ChartErrorBarsType || (ExcelOp.ChartErrorBarsType={}));
	var ChartMapAreaLevel;
	(function (ChartMapAreaLevel) {
		ChartMapAreaLevel["automatic"]="Automatic";
		ChartMapAreaLevel["dataOnly"]="DataOnly";
		ChartMapAreaLevel["city"]="City";
		ChartMapAreaLevel["county"]="County";
		ChartMapAreaLevel["state"]="State";
		ChartMapAreaLevel["country"]="Country";
		ChartMapAreaLevel["continent"]="Continent";
		ChartMapAreaLevel["world"]="World";
	})(ChartMapAreaLevel=ExcelOp.ChartMapAreaLevel || (ExcelOp.ChartMapAreaLevel={}));
	var ChartGradientStyle;
	(function (ChartGradientStyle) {
		ChartGradientStyle["twoPhaseColor"]="TwoPhaseColor";
		ChartGradientStyle["threePhaseColor"]="ThreePhaseColor";
	})(ChartGradientStyle=ExcelOp.ChartGradientStyle || (ExcelOp.ChartGradientStyle={}));
	var ChartGradientStyleType;
	(function (ChartGradientStyleType) {
		ChartGradientStyleType["extremeValue"]="ExtremeValue";
		ChartGradientStyleType["number"]="Number";
		ChartGradientStyleType["percent"]="Percent";
	})(ChartGradientStyleType=ExcelOp.ChartGradientStyleType || (ExcelOp.ChartGradientStyleType={}));
	var ChartTitlePosition;
	(function (ChartTitlePosition) {
		ChartTitlePosition["automatic"]="Automatic";
		ChartTitlePosition["top"]="Top";
		ChartTitlePosition["bottom"]="Bottom";
		ChartTitlePosition["left"]="Left";
		ChartTitlePosition["right"]="Right";
	})(ChartTitlePosition=ExcelOp.ChartTitlePosition || (ExcelOp.ChartTitlePosition={}));
	var ChartLegendPosition;
	(function (ChartLegendPosition) {
		ChartLegendPosition["invalid"]="Invalid";
		ChartLegendPosition["top"]="Top";
		ChartLegendPosition["bottom"]="Bottom";
		ChartLegendPosition["left"]="Left";
		ChartLegendPosition["right"]="Right";
		ChartLegendPosition["corner"]="Corner";
		ChartLegendPosition["custom"]="Custom";
	})(ChartLegendPosition=ExcelOp.ChartLegendPosition || (ExcelOp.ChartLegendPosition={}));
	var ChartMarkerStyle;
	(function (ChartMarkerStyle) {
		ChartMarkerStyle["invalid"]="Invalid";
		ChartMarkerStyle["automatic"]="Automatic";
		ChartMarkerStyle["none"]="None";
		ChartMarkerStyle["square"]="Square";
		ChartMarkerStyle["diamond"]="Diamond";
		ChartMarkerStyle["triangle"]="Triangle";
		ChartMarkerStyle["x"]="X";
		ChartMarkerStyle["star"]="Star";
		ChartMarkerStyle["dot"]="Dot";
		ChartMarkerStyle["dash"]="Dash";
		ChartMarkerStyle["circle"]="Circle";
		ChartMarkerStyle["plus"]="Plus";
		ChartMarkerStyle["picture"]="Picture";
	})(ChartMarkerStyle=ExcelOp.ChartMarkerStyle || (ExcelOp.ChartMarkerStyle={}));
	var ChartPlotAreaPosition;
	(function (ChartPlotAreaPosition) {
		ChartPlotAreaPosition["automatic"]="Automatic";
		ChartPlotAreaPosition["custom"]="Custom";
	})(ChartPlotAreaPosition=ExcelOp.ChartPlotAreaPosition || (ExcelOp.ChartPlotAreaPosition={}));
	var ChartMapLabelStrategy;
	(function (ChartMapLabelStrategy) {
		ChartMapLabelStrategy["none"]="None";
		ChartMapLabelStrategy["bestFit"]="BestFit";
		ChartMapLabelStrategy["showAll"]="ShowAll";
	})(ChartMapLabelStrategy=ExcelOp.ChartMapLabelStrategy || (ExcelOp.ChartMapLabelStrategy={}));
	var ChartMapProjectionType;
	(function (ChartMapProjectionType) {
		ChartMapProjectionType["automatic"]="Automatic";
		ChartMapProjectionType["mercator"]="Mercator";
		ChartMapProjectionType["miller"]="Miller";
		ChartMapProjectionType["robinson"]="Robinson";
		ChartMapProjectionType["albers"]="Albers";
	})(ChartMapProjectionType=ExcelOp.ChartMapProjectionType || (ExcelOp.ChartMapProjectionType={}));
	var ChartParentLabelStrategy;
	(function (ChartParentLabelStrategy) {
		ChartParentLabelStrategy["none"]="None";
		ChartParentLabelStrategy["banner"]="Banner";
		ChartParentLabelStrategy["overlapping"]="Overlapping";
	})(ChartParentLabelStrategy=ExcelOp.ChartParentLabelStrategy || (ExcelOp.ChartParentLabelStrategy={}));
	var ChartSeriesBy;
	(function (ChartSeriesBy) {
		ChartSeriesBy["auto"]="Auto";
		ChartSeriesBy["columns"]="Columns";
		ChartSeriesBy["rows"]="Rows";
	})(ChartSeriesBy=ExcelOp.ChartSeriesBy || (ExcelOp.ChartSeriesBy={}));
	var ChartTextHorizontalAlignment;
	(function (ChartTextHorizontalAlignment) {
		ChartTextHorizontalAlignment["center"]="Center";
		ChartTextHorizontalAlignment["left"]="Left";
		ChartTextHorizontalAlignment["right"]="Right";
		ChartTextHorizontalAlignment["justify"]="Justify";
		ChartTextHorizontalAlignment["distributed"]="Distributed";
	})(ChartTextHorizontalAlignment=ExcelOp.ChartTextHorizontalAlignment || (ExcelOp.ChartTextHorizontalAlignment={}));
	var ChartTextVerticalAlignment;
	(function (ChartTextVerticalAlignment) {
		ChartTextVerticalAlignment["center"]="Center";
		ChartTextVerticalAlignment["bottom"]="Bottom";
		ChartTextVerticalAlignment["top"]="Top";
		ChartTextVerticalAlignment["justify"]="Justify";
		ChartTextVerticalAlignment["distributed"]="Distributed";
	})(ChartTextVerticalAlignment=ExcelOp.ChartTextVerticalAlignment || (ExcelOp.ChartTextVerticalAlignment={}));
	var ChartTickLabelAlignment;
	(function (ChartTickLabelAlignment) {
		ChartTickLabelAlignment["center"]="Center";
		ChartTickLabelAlignment["left"]="Left";
		ChartTickLabelAlignment["right"]="Right";
	})(ChartTickLabelAlignment=ExcelOp.ChartTickLabelAlignment || (ExcelOp.ChartTickLabelAlignment={}));
	var ChartType;
	(function (ChartType) {
		ChartType["invalid"]="Invalid";
		ChartType["columnClustered"]="ColumnClustered";
		ChartType["columnStacked"]="ColumnStacked";
		ChartType["columnStacked100"]="ColumnStacked100";
		ChartType["_3DColumnClustered"]="3DColumnClustered";
		ChartType["_3DColumnStacked"]="3DColumnStacked";
		ChartType["_3DColumnStacked100"]="3DColumnStacked100";
		ChartType["barClustered"]="BarClustered";
		ChartType["barStacked"]="BarStacked";
		ChartType["barStacked100"]="BarStacked100";
		ChartType["_3DBarClustered"]="3DBarClustered";
		ChartType["_3DBarStacked"]="3DBarStacked";
		ChartType["_3DBarStacked100"]="3DBarStacked100";
		ChartType["lineStacked"]="LineStacked";
		ChartType["lineStacked100"]="LineStacked100";
		ChartType["lineMarkers"]="LineMarkers";
		ChartType["lineMarkersStacked"]="LineMarkersStacked";
		ChartType["lineMarkersStacked100"]="LineMarkersStacked100";
		ChartType["pieOfPie"]="PieOfPie";
		ChartType["pieExploded"]="PieExploded";
		ChartType["_3DPieExploded"]="3DPieExploded";
		ChartType["barOfPie"]="BarOfPie";
		ChartType["xyscatterSmooth"]="XYScatterSmooth";
		ChartType["xyscatterSmoothNoMarkers"]="XYScatterSmoothNoMarkers";
		ChartType["xyscatterLines"]="XYScatterLines";
		ChartType["xyscatterLinesNoMarkers"]="XYScatterLinesNoMarkers";
		ChartType["areaStacked"]="AreaStacked";
		ChartType["areaStacked100"]="AreaStacked100";
		ChartType["_3DAreaStacked"]="3DAreaStacked";
		ChartType["_3DAreaStacked100"]="3DAreaStacked100";
		ChartType["doughnutExploded"]="DoughnutExploded";
		ChartType["radarMarkers"]="RadarMarkers";
		ChartType["radarFilled"]="RadarFilled";
		ChartType["surface"]="Surface";
		ChartType["surfaceWireframe"]="SurfaceWireframe";
		ChartType["surfaceTopView"]="SurfaceTopView";
		ChartType["surfaceTopViewWireframe"]="SurfaceTopViewWireframe";
		ChartType["bubble"]="Bubble";
		ChartType["bubble3DEffect"]="Bubble3DEffect";
		ChartType["stockHLC"]="StockHLC";
		ChartType["stockOHLC"]="StockOHLC";
		ChartType["stockVHLC"]="StockVHLC";
		ChartType["stockVOHLC"]="StockVOHLC";
		ChartType["cylinderColClustered"]="CylinderColClustered";
		ChartType["cylinderColStacked"]="CylinderColStacked";
		ChartType["cylinderColStacked100"]="CylinderColStacked100";
		ChartType["cylinderBarClustered"]="CylinderBarClustered";
		ChartType["cylinderBarStacked"]="CylinderBarStacked";
		ChartType["cylinderBarStacked100"]="CylinderBarStacked100";
		ChartType["cylinderCol"]="CylinderCol";
		ChartType["coneColClustered"]="ConeColClustered";
		ChartType["coneColStacked"]="ConeColStacked";
		ChartType["coneColStacked100"]="ConeColStacked100";
		ChartType["coneBarClustered"]="ConeBarClustered";
		ChartType["coneBarStacked"]="ConeBarStacked";
		ChartType["coneBarStacked100"]="ConeBarStacked100";
		ChartType["coneCol"]="ConeCol";
		ChartType["pyramidColClustered"]="PyramidColClustered";
		ChartType["pyramidColStacked"]="PyramidColStacked";
		ChartType["pyramidColStacked100"]="PyramidColStacked100";
		ChartType["pyramidBarClustered"]="PyramidBarClustered";
		ChartType["pyramidBarStacked"]="PyramidBarStacked";
		ChartType["pyramidBarStacked100"]="PyramidBarStacked100";
		ChartType["pyramidCol"]="PyramidCol";
		ChartType["_3DColumn"]="3DColumn";
		ChartType["line"]="Line";
		ChartType["_3DLine"]="3DLine";
		ChartType["_3DPie"]="3DPie";
		ChartType["pie"]="Pie";
		ChartType["xyscatter"]="XYScatter";
		ChartType["_3DArea"]="3DArea";
		ChartType["area"]="Area";
		ChartType["doughnut"]="Doughnut";
		ChartType["radar"]="Radar";
		ChartType["histogram"]="Histogram";
		ChartType["boxwhisker"]="Boxwhisker";
		ChartType["pareto"]="Pareto";
		ChartType["regionMap"]="RegionMap";
		ChartType["treemap"]="Treemap";
		ChartType["waterfall"]="Waterfall";
		ChartType["sunburst"]="Sunburst";
		ChartType["funnel"]="Funnel";
	})(ChartType=ExcelOp.ChartType || (ExcelOp.ChartType={}));
	var ChartUnderlineStyle;
	(function (ChartUnderlineStyle) {
		ChartUnderlineStyle["none"]="None";
		ChartUnderlineStyle["single"]="Single";
	})(ChartUnderlineStyle=ExcelOp.ChartUnderlineStyle || (ExcelOp.ChartUnderlineStyle={}));
	var ChartDisplayBlanksAs;
	(function (ChartDisplayBlanksAs) {
		ChartDisplayBlanksAs["notPlotted"]="NotPlotted";
		ChartDisplayBlanksAs["zero"]="Zero";
		ChartDisplayBlanksAs["interplotted"]="Interplotted";
	})(ChartDisplayBlanksAs=ExcelOp.ChartDisplayBlanksAs || (ExcelOp.ChartDisplayBlanksAs={}));
	var ChartPlotBy;
	(function (ChartPlotBy) {
		ChartPlotBy["rows"]="Rows";
		ChartPlotBy["columns"]="Columns";
	})(ChartPlotBy=ExcelOp.ChartPlotBy || (ExcelOp.ChartPlotBy={}));
	var ChartSplitType;
	(function (ChartSplitType) {
		ChartSplitType["splitByPosition"]="SplitByPosition";
		ChartSplitType["splitByValue"]="SplitByValue";
		ChartSplitType["splitByPercentValue"]="SplitByPercentValue";
		ChartSplitType["splitByCustomSplit"]="SplitByCustomSplit";
	})(ChartSplitType=ExcelOp.ChartSplitType || (ExcelOp.ChartSplitType={}));
	var ChartColorScheme;
	(function (ChartColorScheme) {
		ChartColorScheme["colorfulPalette1"]="ColorfulPalette1";
		ChartColorScheme["colorfulPalette2"]="ColorfulPalette2";
		ChartColorScheme["colorfulPalette3"]="ColorfulPalette3";
		ChartColorScheme["colorfulPalette4"]="ColorfulPalette4";
		ChartColorScheme["monochromaticPalette1"]="MonochromaticPalette1";
		ChartColorScheme["monochromaticPalette2"]="MonochromaticPalette2";
		ChartColorScheme["monochromaticPalette3"]="MonochromaticPalette3";
		ChartColorScheme["monochromaticPalette4"]="MonochromaticPalette4";
		ChartColorScheme["monochromaticPalette5"]="MonochromaticPalette5";
		ChartColorScheme["monochromaticPalette6"]="MonochromaticPalette6";
		ChartColorScheme["monochromaticPalette7"]="MonochromaticPalette7";
		ChartColorScheme["monochromaticPalette8"]="MonochromaticPalette8";
		ChartColorScheme["monochromaticPalette9"]="MonochromaticPalette9";
		ChartColorScheme["monochromaticPalette10"]="MonochromaticPalette10";
		ChartColorScheme["monochromaticPalette11"]="MonochromaticPalette11";
		ChartColorScheme["monochromaticPalette12"]="MonochromaticPalette12";
		ChartColorScheme["monochromaticPalette13"]="MonochromaticPalette13";
	})(ChartColorScheme=ExcelOp.ChartColorScheme || (ExcelOp.ChartColorScheme={}));
	var ChartTrendlineType;
	(function (ChartTrendlineType) {
		ChartTrendlineType["linear"]="Linear";
		ChartTrendlineType["exponential"]="Exponential";
		ChartTrendlineType["logarithmic"]="Logarithmic";
		ChartTrendlineType["movingAverage"]="MovingAverage";
		ChartTrendlineType["polynomial"]="Polynomial";
		ChartTrendlineType["power"]="Power";
	})(ChartTrendlineType=ExcelOp.ChartTrendlineType || (ExcelOp.ChartTrendlineType={}));
	var ShapeZOrder;
	(function (ShapeZOrder) {
		ShapeZOrder["bringToFront"]="BringToFront";
		ShapeZOrder["bringForward"]="BringForward";
		ShapeZOrder["sendToBack"]="SendToBack";
		ShapeZOrder["sendBackward"]="SendBackward";
	})(ShapeZOrder=ExcelOp.ShapeZOrder || (ExcelOp.ShapeZOrder={}));
	var ShapeType;
	(function (ShapeType) {
		ShapeType["unknown"]="Unknown";
		ShapeType["image"]="Image";
		ShapeType["geometricShape"]="GeometricShape";
		ShapeType["group"]="Group";
		ShapeType["line"]="Line";
	})(ShapeType=ExcelOp.ShapeType || (ExcelOp.ShapeType={}));
	var ShapeScaleType;
	(function (ShapeScaleType) {
		ShapeScaleType["currentSize"]="CurrentSize";
		ShapeScaleType["originalSize"]="OriginalSize";
	})(ShapeScaleType=ExcelOp.ShapeScaleType || (ExcelOp.ShapeScaleType={}));
	var ShapeScaleFrom;
	(function (ShapeScaleFrom) {
		ShapeScaleFrom["scaleFromTopLeft"]="ScaleFromTopLeft";
		ShapeScaleFrom["scaleFromMiddle"]="ScaleFromMiddle";
		ShapeScaleFrom["scaleFromBottomRight"]="ScaleFromBottomRight";
	})(ShapeScaleFrom=ExcelOp.ShapeScaleFrom || (ExcelOp.ShapeScaleFrom={}));
	var ShapeFillType;
	(function (ShapeFillType) {
		ShapeFillType["noFill"]="NoFill";
		ShapeFillType["solid"]="Solid";
		ShapeFillType["gradient"]="Gradient";
		ShapeFillType["pattern"]="Pattern";
		ShapeFillType["pictureAndTexture"]="PictureAndTexture";
		ShapeFillType["mixed"]="Mixed";
	})(ShapeFillType=ExcelOp.ShapeFillType || (ExcelOp.ShapeFillType={}));
	var ShapeFontUnderlineStyle;
	(function (ShapeFontUnderlineStyle) {
		ShapeFontUnderlineStyle["none"]="None";
		ShapeFontUnderlineStyle["single"]="Single";
		ShapeFontUnderlineStyle["double"]="Double";
		ShapeFontUnderlineStyle["heavy"]="Heavy";
		ShapeFontUnderlineStyle["dotted"]="Dotted";
		ShapeFontUnderlineStyle["dottedHeavy"]="DottedHeavy";
		ShapeFontUnderlineStyle["dash"]="Dash";
		ShapeFontUnderlineStyle["dashHeavy"]="DashHeavy";
		ShapeFontUnderlineStyle["dashLong"]="DashLong";
		ShapeFontUnderlineStyle["dashLongHeavy"]="DashLongHeavy";
		ShapeFontUnderlineStyle["dotDash"]="DotDash";
		ShapeFontUnderlineStyle["dotDashHeavy"]="DotDashHeavy";
		ShapeFontUnderlineStyle["dotDotDash"]="DotDotDash";
		ShapeFontUnderlineStyle["dotDotDashHeavy"]="DotDotDashHeavy";
		ShapeFontUnderlineStyle["wavy"]="Wavy";
		ShapeFontUnderlineStyle["wavyHeavy"]="WavyHeavy";
		ShapeFontUnderlineStyle["wavyDouble"]="WavyDouble";
	})(ShapeFontUnderlineStyle=ExcelOp.ShapeFontUnderlineStyle || (ExcelOp.ShapeFontUnderlineStyle={}));
	var PictureFormat;
	(function (PictureFormat) {
		PictureFormat["unknown"]="UNKNOWN";
		PictureFormat["bmp"]="BMP";
		PictureFormat["jpeg"]="JPEG";
		PictureFormat["gif"]="GIF";
		PictureFormat["png"]="PNG";
		PictureFormat["svg"]="SVG";
	})(PictureFormat=ExcelOp.PictureFormat || (ExcelOp.PictureFormat={}));
	var ShapeLineStyle;
	(function (ShapeLineStyle) {
		ShapeLineStyle["single"]="Single";
		ShapeLineStyle["thickBetweenThin"]="ThickBetweenThin";
		ShapeLineStyle["thickThin"]="ThickThin";
		ShapeLineStyle["thinThick"]="ThinThick";
		ShapeLineStyle["thinThin"]="ThinThin";
	})(ShapeLineStyle=ExcelOp.ShapeLineStyle || (ExcelOp.ShapeLineStyle={}));
	var ShapeLineDashStyle;
	(function (ShapeLineDashStyle) {
		ShapeLineDashStyle["dash"]="Dash";
		ShapeLineDashStyle["dashDot"]="DashDot";
		ShapeLineDashStyle["dashDotDot"]="DashDotDot";
		ShapeLineDashStyle["longDash"]="LongDash";
		ShapeLineDashStyle["longDashDot"]="LongDashDot";
		ShapeLineDashStyle["roundDot"]="RoundDot";
		ShapeLineDashStyle["solid"]="Solid";
		ShapeLineDashStyle["squareDot"]="SquareDot";
		ShapeLineDashStyle["longDashDotDot"]="LongDashDotDot";
		ShapeLineDashStyle["systemDash"]="SystemDash";
		ShapeLineDashStyle["systemDot"]="SystemDot";
		ShapeLineDashStyle["systemDashDot"]="SystemDashDot";
	})(ShapeLineDashStyle=ExcelOp.ShapeLineDashStyle || (ExcelOp.ShapeLineDashStyle={}));
	var ArrowHeadLength;
	(function (ArrowHeadLength) {
		ArrowHeadLength["short"]="Short";
		ArrowHeadLength["medium"]="Medium";
		ArrowHeadLength["long"]="Long";
	})(ArrowHeadLength=ExcelOp.ArrowHeadLength || (ExcelOp.ArrowHeadLength={}));
	var ArrowHeadStyle;
	(function (ArrowHeadStyle) {
		ArrowHeadStyle["none"]="None";
		ArrowHeadStyle["triangle"]="Triangle";
		ArrowHeadStyle["stealth"]="Stealth";
		ArrowHeadStyle["diamond"]="Diamond";
		ArrowHeadStyle["oval"]="Oval";
		ArrowHeadStyle["open"]="Open";
	})(ArrowHeadStyle=ExcelOp.ArrowHeadStyle || (ExcelOp.ArrowHeadStyle={}));
	var ArrowHeadWidth;
	(function (ArrowHeadWidth) {
		ArrowHeadWidth["narrow"]="Narrow";
		ArrowHeadWidth["medium"]="Medium";
		ArrowHeadWidth["wide"]="Wide";
	})(ArrowHeadWidth=ExcelOp.ArrowHeadWidth || (ExcelOp.ArrowHeadWidth={}));
	var BindingType;
	(function (BindingType) {
		BindingType["range"]="Range";
		BindingType["table"]="Table";
		BindingType["text"]="Text";
	})(BindingType=ExcelOp.BindingType || (ExcelOp.BindingType={}));
	var BorderIndex;
	(function (BorderIndex) {
		BorderIndex["edgeTop"]="EdgeTop";
		BorderIndex["edgeBottom"]="EdgeBottom";
		BorderIndex["edgeLeft"]="EdgeLeft";
		BorderIndex["edgeRight"]="EdgeRight";
		BorderIndex["insideVertical"]="InsideVertical";
		BorderIndex["insideHorizontal"]="InsideHorizontal";
		BorderIndex["diagonalDown"]="DiagonalDown";
		BorderIndex["diagonalUp"]="DiagonalUp";
	})(BorderIndex=ExcelOp.BorderIndex || (ExcelOp.BorderIndex={}));
	var BorderLineStyle;
	(function (BorderLineStyle) {
		BorderLineStyle["none"]="None";
		BorderLineStyle["continuous"]="Continuous";
		BorderLineStyle["dash"]="Dash";
		BorderLineStyle["dashDot"]="DashDot";
		BorderLineStyle["dashDotDot"]="DashDotDot";
		BorderLineStyle["dot"]="Dot";
		BorderLineStyle["double"]="Double";
		BorderLineStyle["slantDashDot"]="SlantDashDot";
	})(BorderLineStyle=ExcelOp.BorderLineStyle || (ExcelOp.BorderLineStyle={}));
	var BorderWeight;
	(function (BorderWeight) {
		BorderWeight["hairline"]="Hairline";
		BorderWeight["thin"]="Thin";
		BorderWeight["medium"]="Medium";
		BorderWeight["thick"]="Thick";
	})(BorderWeight=ExcelOp.BorderWeight || (ExcelOp.BorderWeight={}));
	var CalculationMode;
	(function (CalculationMode) {
		CalculationMode["automatic"]="Automatic";
		CalculationMode["automaticExceptTables"]="AutomaticExceptTables";
		CalculationMode["manual"]="Manual";
	})(CalculationMode=ExcelOp.CalculationMode || (ExcelOp.CalculationMode={}));
	var CalculationType;
	(function (CalculationType) {
		CalculationType["recalculate"]="Recalculate";
		CalculationType["full"]="Full";
		CalculationType["fullRebuild"]="FullRebuild";
	})(CalculationType=ExcelOp.CalculationType || (ExcelOp.CalculationType={}));
	var ClearApplyTo;
	(function (ClearApplyTo) {
		ClearApplyTo["all"]="All";
		ClearApplyTo["formats"]="Formats";
		ClearApplyTo["contents"]="Contents";
		ClearApplyTo["hyperlinks"]="Hyperlinks";
		ClearApplyTo["removeHyperlinks"]="RemoveHyperlinks";
	})(ClearApplyTo=ExcelOp.ClearApplyTo || (ExcelOp.ClearApplyTo={}));
	var VisualCategory;
	(function (VisualCategory) {
		VisualCategory["column"]="Column";
		VisualCategory["bar"]="Bar";
		VisualCategory["line"]="Line";
		VisualCategory["area"]="Area";
		VisualCategory["pie"]="Pie";
		VisualCategory["donut"]="Donut";
		VisualCategory["scatter"]="Scatter";
		VisualCategory["bubble"]="Bubble";
		VisualCategory["statistical"]="Statistical";
		VisualCategory["stock"]="Stock";
		VisualCategory["combo"]="Combo";
		VisualCategory["hierarchy"]="Hierarchy";
		VisualCategory["surface"]="Surface";
		VisualCategory["map"]="Map";
		VisualCategory["funnel"]="Funnel";
		VisualCategory["radar"]="Radar";
		VisualCategory["waterfall"]="Waterfall";
		VisualCategory["threeD"]="ThreeD";
		VisualCategory["other"]="Other";
	})(VisualCategory=ExcelOp.VisualCategory || (ExcelOp.VisualCategory={}));
	var VisualPropertyType;
	(function (VisualPropertyType) {
		VisualPropertyType["object"]="Object";
		VisualPropertyType["collection"]="Collection";
		VisualPropertyType["string"]="String";
		VisualPropertyType["double"]="Double";
		VisualPropertyType["int"]="Int";
		VisualPropertyType["bool"]="Bool";
		VisualPropertyType["enum"]="Enum";
		VisualPropertyType["color"]="Color";
	})(VisualPropertyType=ExcelOp.VisualPropertyType || (ExcelOp.VisualPropertyType={}));
	var VisualChangeType;
	(function (VisualChangeType) {
		VisualChangeType["dataChange"]="DataChange";
		VisualChangeType["propertyChange"]="PropertyChange";
		VisualChangeType["genericChange"]="GenericChange";
		VisualChangeType["selectionChange"]="SelectionChange";
	})(VisualChangeType=ExcelOp.VisualChangeType || (ExcelOp.VisualChangeType={}));
	var BoolMetaPropertyType;
	(function (BoolMetaPropertyType) {
		BoolMetaPropertyType["writeOnly"]="WriteOnly";
		BoolMetaPropertyType["readOnly"]="ReadOnly";
		BoolMetaPropertyType["hideUI"]="HideUI";
		BoolMetaPropertyType["nextPropOnSameLine"]="NextPropOnSameLine";
		BoolMetaPropertyType["hideLabel"]="HideLabel";
		BoolMetaPropertyType["showResetUI"]="ShowResetUI";
		BoolMetaPropertyType["hasOwnExpandableSection"]="HasOwnExpandableSection";
		BoolMetaPropertyType["untransferable"]="Untransferable";
	})(BoolMetaPropertyType=ExcelOp.BoolMetaPropertyType || (ExcelOp.BoolMetaPropertyType={}));
	var ConditionalDataBarAxisFormat;
	(function (ConditionalDataBarAxisFormat) {
		ConditionalDataBarAxisFormat["automatic"]="Automatic";
		ConditionalDataBarAxisFormat["none"]="None";
		ConditionalDataBarAxisFormat["cellMidPoint"]="CellMidPoint";
	})(ConditionalDataBarAxisFormat=ExcelOp.ConditionalDataBarAxisFormat || (ExcelOp.ConditionalDataBarAxisFormat={}));
	var ConditionalDataBarDirection;
	(function (ConditionalDataBarDirection) {
		ConditionalDataBarDirection["context"]="Context";
		ConditionalDataBarDirection["leftToRight"]="LeftToRight";
		ConditionalDataBarDirection["rightToLeft"]="RightToLeft";
	})(ConditionalDataBarDirection=ExcelOp.ConditionalDataBarDirection || (ExcelOp.ConditionalDataBarDirection={}));
	var ConditionalFormatDirection;
	(function (ConditionalFormatDirection) {
		ConditionalFormatDirection["top"]="Top";
		ConditionalFormatDirection["bottom"]="Bottom";
	})(ConditionalFormatDirection=ExcelOp.ConditionalFormatDirection || (ExcelOp.ConditionalFormatDirection={}));
	var ConditionalFormatType;
	(function (ConditionalFormatType) {
		ConditionalFormatType["custom"]="Custom";
		ConditionalFormatType["dataBar"]="DataBar";
		ConditionalFormatType["colorScale"]="ColorScale";
		ConditionalFormatType["iconSet"]="IconSet";
		ConditionalFormatType["topBottom"]="TopBottom";
		ConditionalFormatType["presetCriteria"]="PresetCriteria";
		ConditionalFormatType["containsText"]="ContainsText";
		ConditionalFormatType["cellValue"]="CellValue";
	})(ConditionalFormatType=ExcelOp.ConditionalFormatType || (ExcelOp.ConditionalFormatType={}));
	var ConditionalFormatRuleType;
	(function (ConditionalFormatRuleType) {
		ConditionalFormatRuleType["invalid"]="Invalid";
		ConditionalFormatRuleType["automatic"]="Automatic";
		ConditionalFormatRuleType["lowestValue"]="LowestValue";
		ConditionalFormatRuleType["highestValue"]="HighestValue";
		ConditionalFormatRuleType["number"]="Number";
		ConditionalFormatRuleType["percent"]="Percent";
		ConditionalFormatRuleType["formula"]="Formula";
		ConditionalFormatRuleType["percentile"]="Percentile";
	})(ConditionalFormatRuleType=ExcelOp.ConditionalFormatRuleType || (ExcelOp.ConditionalFormatRuleType={}));
	var ConditionalFormatIconRuleType;
	(function (ConditionalFormatIconRuleType) {
		ConditionalFormatIconRuleType["invalid"]="Invalid";
		ConditionalFormatIconRuleType["number"]="Number";
		ConditionalFormatIconRuleType["percent"]="Percent";
		ConditionalFormatIconRuleType["formula"]="Formula";
		ConditionalFormatIconRuleType["percentile"]="Percentile";
	})(ConditionalFormatIconRuleType=ExcelOp.ConditionalFormatIconRuleType || (ExcelOp.ConditionalFormatIconRuleType={}));
	var ConditionalFormatColorCriterionType;
	(function (ConditionalFormatColorCriterionType) {
		ConditionalFormatColorCriterionType["invalid"]="Invalid";
		ConditionalFormatColorCriterionType["lowestValue"]="LowestValue";
		ConditionalFormatColorCriterionType["highestValue"]="HighestValue";
		ConditionalFormatColorCriterionType["number"]="Number";
		ConditionalFormatColorCriterionType["percent"]="Percent";
		ConditionalFormatColorCriterionType["formula"]="Formula";
		ConditionalFormatColorCriterionType["percentile"]="Percentile";
	})(ConditionalFormatColorCriterionType=ExcelOp.ConditionalFormatColorCriterionType || (ExcelOp.ConditionalFormatColorCriterionType={}));
	var ConditionalTopBottomCriterionType;
	(function (ConditionalTopBottomCriterionType) {
		ConditionalTopBottomCriterionType["invalid"]="Invalid";
		ConditionalTopBottomCriterionType["topItems"]="TopItems";
		ConditionalTopBottomCriterionType["topPercent"]="TopPercent";
		ConditionalTopBottomCriterionType["bottomItems"]="BottomItems";
		ConditionalTopBottomCriterionType["bottomPercent"]="BottomPercent";
	})(ConditionalTopBottomCriterionType=ExcelOp.ConditionalTopBottomCriterionType || (ExcelOp.ConditionalTopBottomCriterionType={}));
	var ConditionalFormatPresetCriterion;
	(function (ConditionalFormatPresetCriterion) {
		ConditionalFormatPresetCriterion["invalid"]="Invalid";
		ConditionalFormatPresetCriterion["blanks"]="Blanks";
		ConditionalFormatPresetCriterion["nonBlanks"]="NonBlanks";
		ConditionalFormatPresetCriterion["errors"]="Errors";
		ConditionalFormatPresetCriterion["nonErrors"]="NonErrors";
		ConditionalFormatPresetCriterion["yesterday"]="Yesterday";
		ConditionalFormatPresetCriterion["today"]="Today";
		ConditionalFormatPresetCriterion["tomorrow"]="Tomorrow";
		ConditionalFormatPresetCriterion["lastSevenDays"]="LastSevenDays";
		ConditionalFormatPresetCriterion["lastWeek"]="LastWeek";
		ConditionalFormatPresetCriterion["thisWeek"]="ThisWeek";
		ConditionalFormatPresetCriterion["nextWeek"]="NextWeek";
		ConditionalFormatPresetCriterion["lastMonth"]="LastMonth";
		ConditionalFormatPresetCriterion["thisMonth"]="ThisMonth";
		ConditionalFormatPresetCriterion["nextMonth"]="NextMonth";
		ConditionalFormatPresetCriterion["aboveAverage"]="AboveAverage";
		ConditionalFormatPresetCriterion["belowAverage"]="BelowAverage";
		ConditionalFormatPresetCriterion["equalOrAboveAverage"]="EqualOrAboveAverage";
		ConditionalFormatPresetCriterion["equalOrBelowAverage"]="EqualOrBelowAverage";
		ConditionalFormatPresetCriterion["oneStdDevAboveAverage"]="OneStdDevAboveAverage";
		ConditionalFormatPresetCriterion["oneStdDevBelowAverage"]="OneStdDevBelowAverage";
		ConditionalFormatPresetCriterion["twoStdDevAboveAverage"]="TwoStdDevAboveAverage";
		ConditionalFormatPresetCriterion["twoStdDevBelowAverage"]="TwoStdDevBelowAverage";
		ConditionalFormatPresetCriterion["threeStdDevAboveAverage"]="ThreeStdDevAboveAverage";
		ConditionalFormatPresetCriterion["threeStdDevBelowAverage"]="ThreeStdDevBelowAverage";
		ConditionalFormatPresetCriterion["uniqueValues"]="UniqueValues";
		ConditionalFormatPresetCriterion["duplicateValues"]="DuplicateValues";
	})(ConditionalFormatPresetCriterion=ExcelOp.ConditionalFormatPresetCriterion || (ExcelOp.ConditionalFormatPresetCriterion={}));
	var ConditionalTextOperator;
	(function (ConditionalTextOperator) {
		ConditionalTextOperator["invalid"]="Invalid";
		ConditionalTextOperator["contains"]="Contains";
		ConditionalTextOperator["notContains"]="NotContains";
		ConditionalTextOperator["beginsWith"]="BeginsWith";
		ConditionalTextOperator["endsWith"]="EndsWith";
	})(ConditionalTextOperator=ExcelOp.ConditionalTextOperator || (ExcelOp.ConditionalTextOperator={}));
	var ConditionalCellValueOperator;
	(function (ConditionalCellValueOperator) {
		ConditionalCellValueOperator["invalid"]="Invalid";
		ConditionalCellValueOperator["between"]="Between";
		ConditionalCellValueOperator["notBetween"]="NotBetween";
		ConditionalCellValueOperator["equalTo"]="EqualTo";
		ConditionalCellValueOperator["notEqualTo"]="NotEqualTo";
		ConditionalCellValueOperator["greaterThan"]="GreaterThan";
		ConditionalCellValueOperator["lessThan"]="LessThan";
		ConditionalCellValueOperator["greaterThanOrEqual"]="GreaterThanOrEqual";
		ConditionalCellValueOperator["lessThanOrEqual"]="LessThanOrEqual";
	})(ConditionalCellValueOperator=ExcelOp.ConditionalCellValueOperator || (ExcelOp.ConditionalCellValueOperator={}));
	var ConditionalIconCriterionOperator;
	(function (ConditionalIconCriterionOperator) {
		ConditionalIconCriterionOperator["invalid"]="Invalid";
		ConditionalIconCriterionOperator["greaterThan"]="GreaterThan";
		ConditionalIconCriterionOperator["greaterThanOrEqual"]="GreaterThanOrEqual";
	})(ConditionalIconCriterionOperator=ExcelOp.ConditionalIconCriterionOperator || (ExcelOp.ConditionalIconCriterionOperator={}));
	var ConditionalRangeBorderIndex;
	(function (ConditionalRangeBorderIndex) {
		ConditionalRangeBorderIndex["edgeTop"]="EdgeTop";
		ConditionalRangeBorderIndex["edgeBottom"]="EdgeBottom";
		ConditionalRangeBorderIndex["edgeLeft"]="EdgeLeft";
		ConditionalRangeBorderIndex["edgeRight"]="EdgeRight";
	})(ConditionalRangeBorderIndex=ExcelOp.ConditionalRangeBorderIndex || (ExcelOp.ConditionalRangeBorderIndex={}));
	var ConditionalRangeBorderLineStyle;
	(function (ConditionalRangeBorderLineStyle) {
		ConditionalRangeBorderLineStyle["none"]="None";
		ConditionalRangeBorderLineStyle["continuous"]="Continuous";
		ConditionalRangeBorderLineStyle["dash"]="Dash";
		ConditionalRangeBorderLineStyle["dashDot"]="DashDot";
		ConditionalRangeBorderLineStyle["dashDotDot"]="DashDotDot";
		ConditionalRangeBorderLineStyle["dot"]="Dot";
	})(ConditionalRangeBorderLineStyle=ExcelOp.ConditionalRangeBorderLineStyle || (ExcelOp.ConditionalRangeBorderLineStyle={}));
	var ConditionalRangeFontUnderlineStyle;
	(function (ConditionalRangeFontUnderlineStyle) {
		ConditionalRangeFontUnderlineStyle["none"]="None";
		ConditionalRangeFontUnderlineStyle["single"]="Single";
		ConditionalRangeFontUnderlineStyle["double"]="Double";
	})(ConditionalRangeFontUnderlineStyle=ExcelOp.ConditionalRangeFontUnderlineStyle || (ExcelOp.ConditionalRangeFontUnderlineStyle={}));
	var CustomFunctionType;
	(function (CustomFunctionType) {
		CustomFunctionType["invalid"]="Invalid";
		CustomFunctionType["script"]="Script";
		CustomFunctionType["webService"]="WebService";
	})(CustomFunctionType=ExcelOp.CustomFunctionType || (ExcelOp.CustomFunctionType={}));
	var CustomFunctionMetadataFormat;
	(function (CustomFunctionMetadataFormat) {
		CustomFunctionMetadataFormat["invalid"]="Invalid";
		CustomFunctionMetadataFormat["openApi"]="OpenApi";
	})(CustomFunctionMetadataFormat=ExcelOp.CustomFunctionMetadataFormat || (ExcelOp.CustomFunctionMetadataFormat={}));
	var DataValidationType;
	(function (DataValidationType) {
		DataValidationType["none"]="None";
		DataValidationType["wholeNumber"]="WholeNumber";
		DataValidationType["decimal"]="Decimal";
		DataValidationType["list"]="List";
		DataValidationType["date"]="Date";
		DataValidationType["time"]="Time";
		DataValidationType["textLength"]="TextLength";
		DataValidationType["custom"]="Custom";
		DataValidationType["inconsistent"]="Inconsistent";
		DataValidationType["mixedCriteria"]="MixedCriteria";
	})(DataValidationType=ExcelOp.DataValidationType || (ExcelOp.DataValidationType={}));
	var DataValidationOperator;
	(function (DataValidationOperator) {
		DataValidationOperator["between"]="Between";
		DataValidationOperator["notBetween"]="NotBetween";
		DataValidationOperator["equalTo"]="EqualTo";
		DataValidationOperator["notEqualTo"]="NotEqualTo";
		DataValidationOperator["greaterThan"]="GreaterThan";
		DataValidationOperator["lessThan"]="LessThan";
		DataValidationOperator["greaterThanOrEqualTo"]="GreaterThanOrEqualTo";
		DataValidationOperator["lessThanOrEqualTo"]="LessThanOrEqualTo";
	})(DataValidationOperator=ExcelOp.DataValidationOperator || (ExcelOp.DataValidationOperator={}));
	var DataValidationAlertStyle;
	(function (DataValidationAlertStyle) {
		DataValidationAlertStyle["stop"]="Stop";
		DataValidationAlertStyle["warning"]="Warning";
		DataValidationAlertStyle["information"]="Information";
	})(DataValidationAlertStyle=ExcelOp.DataValidationAlertStyle || (ExcelOp.DataValidationAlertStyle={}));
	var DeleteShiftDirection;
	(function (DeleteShiftDirection) {
		DeleteShiftDirection["up"]="Up";
		DeleteShiftDirection["left"]="Left";
	})(DeleteShiftDirection=ExcelOp.DeleteShiftDirection || (ExcelOp.DeleteShiftDirection={}));
	var DynamicFilterCriteria;
	(function (DynamicFilterCriteria) {
		DynamicFilterCriteria["unknown"]="Unknown";
		DynamicFilterCriteria["aboveAverage"]="AboveAverage";
		DynamicFilterCriteria["allDatesInPeriodApril"]="AllDatesInPeriodApril";
		DynamicFilterCriteria["allDatesInPeriodAugust"]="AllDatesInPeriodAugust";
		DynamicFilterCriteria["allDatesInPeriodDecember"]="AllDatesInPeriodDecember";
		DynamicFilterCriteria["allDatesInPeriodFebruray"]="AllDatesInPeriodFebruray";
		DynamicFilterCriteria["allDatesInPeriodJanuary"]="AllDatesInPeriodJanuary";
		DynamicFilterCriteria["allDatesInPeriodJuly"]="AllDatesInPeriodJuly";
		DynamicFilterCriteria["allDatesInPeriodJune"]="AllDatesInPeriodJune";
		DynamicFilterCriteria["allDatesInPeriodMarch"]="AllDatesInPeriodMarch";
		DynamicFilterCriteria["allDatesInPeriodMay"]="AllDatesInPeriodMay";
		DynamicFilterCriteria["allDatesInPeriodNovember"]="AllDatesInPeriodNovember";
		DynamicFilterCriteria["allDatesInPeriodOctober"]="AllDatesInPeriodOctober";
		DynamicFilterCriteria["allDatesInPeriodQuarter1"]="AllDatesInPeriodQuarter1";
		DynamicFilterCriteria["allDatesInPeriodQuarter2"]="AllDatesInPeriodQuarter2";
		DynamicFilterCriteria["allDatesInPeriodQuarter3"]="AllDatesInPeriodQuarter3";
		DynamicFilterCriteria["allDatesInPeriodQuarter4"]="AllDatesInPeriodQuarter4";
		DynamicFilterCriteria["allDatesInPeriodSeptember"]="AllDatesInPeriodSeptember";
		DynamicFilterCriteria["belowAverage"]="BelowAverage";
		DynamicFilterCriteria["lastMonth"]="LastMonth";
		DynamicFilterCriteria["lastQuarter"]="LastQuarter";
		DynamicFilterCriteria["lastWeek"]="LastWeek";
		DynamicFilterCriteria["lastYear"]="LastYear";
		DynamicFilterCriteria["nextMonth"]="NextMonth";
		DynamicFilterCriteria["nextQuarter"]="NextQuarter";
		DynamicFilterCriteria["nextWeek"]="NextWeek";
		DynamicFilterCriteria["nextYear"]="NextYear";
		DynamicFilterCriteria["thisMonth"]="ThisMonth";
		DynamicFilterCriteria["thisQuarter"]="ThisQuarter";
		DynamicFilterCriteria["thisWeek"]="ThisWeek";
		DynamicFilterCriteria["thisYear"]="ThisYear";
		DynamicFilterCriteria["today"]="Today";
		DynamicFilterCriteria["tomorrow"]="Tomorrow";
		DynamicFilterCriteria["yearToDate"]="YearToDate";
		DynamicFilterCriteria["yesterday"]="Yesterday";
	})(DynamicFilterCriteria=ExcelOp.DynamicFilterCriteria || (ExcelOp.DynamicFilterCriteria={}));
	var FilterDatetimeSpecificity;
	(function (FilterDatetimeSpecificity) {
		FilterDatetimeSpecificity["year"]="Year";
		FilterDatetimeSpecificity["month"]="Month";
		FilterDatetimeSpecificity["day"]="Day";
		FilterDatetimeSpecificity["hour"]="Hour";
		FilterDatetimeSpecificity["minute"]="Minute";
		FilterDatetimeSpecificity["second"]="Second";
	})(FilterDatetimeSpecificity=ExcelOp.FilterDatetimeSpecificity || (ExcelOp.FilterDatetimeSpecificity={}));
	var FilterOn;
	(function (FilterOn) {
		FilterOn["bottomItems"]="BottomItems";
		FilterOn["bottomPercent"]="BottomPercent";
		FilterOn["cellColor"]="CellColor";
		FilterOn["dynamic"]="Dynamic";
		FilterOn["fontColor"]="FontColor";
		FilterOn["values"]="Values";
		FilterOn["topItems"]="TopItems";
		FilterOn["topPercent"]="TopPercent";
		FilterOn["icon"]="Icon";
		FilterOn["custom"]="Custom";
	})(FilterOn=ExcelOp.FilterOn || (ExcelOp.FilterOn={}));
	var FilterOperator;
	(function (FilterOperator) {
		FilterOperator["and"]="And";
		FilterOperator["or"]="Or";
	})(FilterOperator=ExcelOp.FilterOperator || (ExcelOp.FilterOperator={}));
	var HorizontalAlignment;
	(function (HorizontalAlignment) {
		HorizontalAlignment["general"]="General";
		HorizontalAlignment["left"]="Left";
		HorizontalAlignment["center"]="Center";
		HorizontalAlignment["right"]="Right";
		HorizontalAlignment["fill"]="Fill";
		HorizontalAlignment["justify"]="Justify";
		HorizontalAlignment["centerAcrossSelection"]="CenterAcrossSelection";
		HorizontalAlignment["distributed"]="Distributed";
	})(HorizontalAlignment=ExcelOp.HorizontalAlignment || (ExcelOp.HorizontalAlignment={}));
	var IconSet;
	(function (IconSet) {
		IconSet["invalid"]="Invalid";
		IconSet["threeArrows"]="ThreeArrows";
		IconSet["threeArrowsGray"]="ThreeArrowsGray";
		IconSet["threeFlags"]="ThreeFlags";
		IconSet["threeTrafficLights1"]="ThreeTrafficLights1";
		IconSet["threeTrafficLights2"]="ThreeTrafficLights2";
		IconSet["threeSigns"]="ThreeSigns";
		IconSet["threeSymbols"]="ThreeSymbols";
		IconSet["threeSymbols2"]="ThreeSymbols2";
		IconSet["fourArrows"]="FourArrows";
		IconSet["fourArrowsGray"]="FourArrowsGray";
		IconSet["fourRedToBlack"]="FourRedToBlack";
		IconSet["fourRating"]="FourRating";
		IconSet["fourTrafficLights"]="FourTrafficLights";
		IconSet["fiveArrows"]="FiveArrows";
		IconSet["fiveArrowsGray"]="FiveArrowsGray";
		IconSet["fiveRating"]="FiveRating";
		IconSet["fiveQuarters"]="FiveQuarters";
		IconSet["threeStars"]="ThreeStars";
		IconSet["threeTriangles"]="ThreeTriangles";
		IconSet["fiveBoxes"]="FiveBoxes";
		IconSet["linkedEntityFinanceIcon"]="LinkedEntityFinanceIcon";
		IconSet["linkedEntityMapIcon"]="LinkedEntityMapIcon";
	})(IconSet=ExcelOp.IconSet || (ExcelOp.IconSet={}));
	var ImageFittingMode;
	(function (ImageFittingMode) {
		ImageFittingMode["fit"]="Fit";
		ImageFittingMode["fitAndCenter"]="FitAndCenter";
		ImageFittingMode["fill"]="Fill";
	})(ImageFittingMode=ExcelOp.ImageFittingMode || (ExcelOp.ImageFittingMode={}));
	var InsertShiftDirection;
	(function (InsertShiftDirection) {
		InsertShiftDirection["down"]="Down";
		InsertShiftDirection["right"]="Right";
	})(InsertShiftDirection=ExcelOp.InsertShiftDirection || (ExcelOp.InsertShiftDirection={}));
	var NamedItemScope;
	(function (NamedItemScope) {
		NamedItemScope["worksheet"]="Worksheet";
		NamedItemScope["workbook"]="Workbook";
	})(NamedItemScope=ExcelOp.NamedItemScope || (ExcelOp.NamedItemScope={}));
	var NamedItemType;
	(function (NamedItemType) {
		NamedItemType["string"]="String";
		NamedItemType["integer"]="Integer";
		NamedItemType["double"]="Double";
		NamedItemType["boolean"]="Boolean";
		NamedItemType["range"]="Range";
		NamedItemType["error"]="Error";
		NamedItemType["array"]="Array";
	})(NamedItemType=ExcelOp.NamedItemType || (ExcelOp.NamedItemType={}));
	var RangeUnderlineStyle;
	(function (RangeUnderlineStyle) {
		RangeUnderlineStyle["none"]="None";
		RangeUnderlineStyle["single"]="Single";
		RangeUnderlineStyle["double"]="Double";
		RangeUnderlineStyle["singleAccountant"]="SingleAccountant";
		RangeUnderlineStyle["doubleAccountant"]="DoubleAccountant";
	})(RangeUnderlineStyle=ExcelOp.RangeUnderlineStyle || (ExcelOp.RangeUnderlineStyle={}));
	var SheetVisibility;
	(function (SheetVisibility) {
		SheetVisibility["visible"]="Visible";
		SheetVisibility["hidden"]="Hidden";
		SheetVisibility["veryHidden"]="VeryHidden";
	})(SheetVisibility=ExcelOp.SheetVisibility || (ExcelOp.SheetVisibility={}));
	var RangeValueType;
	(function (RangeValueType) {
		RangeValueType["unknown"]="Unknown";
		RangeValueType["empty"]="Empty";
		RangeValueType["string"]="String";
		RangeValueType["integer"]="Integer";
		RangeValueType["double"]="Double";
		RangeValueType["boolean"]="Boolean";
		RangeValueType["error"]="Error";
		RangeValueType["richValue"]="RichValue";
	})(RangeValueType=ExcelOp.RangeValueType || (ExcelOp.RangeValueType={}));
	var SearchDirection;
	(function (SearchDirection) {
		SearchDirection["forward"]="Forward";
		SearchDirection["backwards"]="Backwards";
	})(SearchDirection=ExcelOp.SearchDirection || (ExcelOp.SearchDirection={}));
	var SortOrientation;
	(function (SortOrientation) {
		SortOrientation["rows"]="Rows";
		SortOrientation["columns"]="Columns";
	})(SortOrientation=ExcelOp.SortOrientation || (ExcelOp.SortOrientation={}));
	var SortOn;
	(function (SortOn) {
		SortOn["value"]="Value";
		SortOn["cellColor"]="CellColor";
		SortOn["fontColor"]="FontColor";
		SortOn["icon"]="Icon";
	})(SortOn=ExcelOp.SortOn || (ExcelOp.SortOn={}));
	var SortDataOption;
	(function (SortDataOption) {
		SortDataOption["normal"]="Normal";
		SortDataOption["textAsNumber"]="TextAsNumber";
	})(SortDataOption=ExcelOp.SortDataOption || (ExcelOp.SortDataOption={}));
	var SortMethod;
	(function (SortMethod) {
		SortMethod["pinYin"]="PinYin";
		SortMethod["strokeCount"]="StrokeCount";
	})(SortMethod=ExcelOp.SortMethod || (ExcelOp.SortMethod={}));
	var VerticalAlignment;
	(function (VerticalAlignment) {
		VerticalAlignment["top"]="Top";
		VerticalAlignment["center"]="Center";
		VerticalAlignment["bottom"]="Bottom";
		VerticalAlignment["justify"]="Justify";
		VerticalAlignment["distributed"]="Distributed";
	})(VerticalAlignment=ExcelOp.VerticalAlignment || (ExcelOp.VerticalAlignment={}));
	var DocumentPropertyType;
	(function (DocumentPropertyType) {
		DocumentPropertyType["number"]="Number";
		DocumentPropertyType["boolean"]="Boolean";
		DocumentPropertyType["date"]="Date";
		DocumentPropertyType["string"]="String";
		DocumentPropertyType["float"]="Float";
	})(DocumentPropertyType=ExcelOp.DocumentPropertyType || (ExcelOp.DocumentPropertyType={}));
	var EventSource;
	(function (EventSource) {
		EventSource["local"]="Local";
		EventSource["remote"]="Remote";
	})(EventSource=ExcelOp.EventSource || (ExcelOp.EventSource={}));
	var DataChangeType;
	(function (DataChangeType) {
		DataChangeType["unknown"]="Unknown";
		DataChangeType["rangeEdited"]="RangeEdited";
		DataChangeType["rowInserted"]="RowInserted";
		DataChangeType["rowDeleted"]="RowDeleted";
		DataChangeType["columnInserted"]="ColumnInserted";
		DataChangeType["columnDeleted"]="ColumnDeleted";
		DataChangeType["cellInserted"]="CellInserted";
		DataChangeType["cellDeleted"]="CellDeleted";
	})(DataChangeType=ExcelOp.DataChangeType || (ExcelOp.DataChangeType={}));
	var EventType;
	(function (EventType) {
		EventType["worksheetChanged"]="WorksheetChanged";
		EventType["worksheetSelectionChanged"]="WorksheetSelectionChanged";
		EventType["worksheetAdded"]="WorksheetAdded";
		EventType["worksheetActivated"]="WorksheetActivated";
		EventType["worksheetDeactivated"]="WorksheetDeactivated";
		EventType["tableChanged"]="TableChanged";
		EventType["tableSelectionChanged"]="TableSelectionChanged";
		EventType["worksheetDeleted"]="WorksheetDeleted";
		EventType["chartAdded"]="ChartAdded";
		EventType["chartActivated"]="ChartActivated";
		EventType["chartDeactivated"]="ChartDeactivated";
		EventType["chartDeleted"]="ChartDeleted";
		EventType["worksheetCalculated"]="WorksheetCalculated";
		EventType["visualSelectionChanged"]="VisualSelectionChanged";
		EventType["agaveVisualUpdate"]="AgaveVisualUpdate";
		EventType["tableAdded"]="TableAdded";
		EventType["tableDeleted"]="TableDeleted";
		EventType["tableFiltered"]="TableFiltered";
		EventType["worksheetFiltered"]="WorksheetFiltered";
		EventType["shapeActivated"]="ShapeActivated";
		EventType["shapeDeactivated"]="ShapeDeactivated";
		EventType["visualChange"]="VisualChange";
		EventType["workbookAutoSaveSettingChanged"]="WorkbookAutoSaveSettingChanged";
		EventType["worksheetFormatChanged"]="WorksheetFormatChanged";
		EventType["wacoperationEvent"]="WACOperationEvent";
		EventType["ribbonCommandExecuted"]="RibbonCommandExecuted";
	})(EventType=ExcelOp.EventType || (ExcelOp.EventType={}));
	var DocumentPropertyItem;
	(function (DocumentPropertyItem) {
		DocumentPropertyItem["title"]="Title";
		DocumentPropertyItem["subject"]="Subject";
		DocumentPropertyItem["author"]="Author";
		DocumentPropertyItem["keywords"]="Keywords";
		DocumentPropertyItem["comments"]="Comments";
		DocumentPropertyItem["template"]="Template";
		DocumentPropertyItem["lastAuth"]="LastAuth";
		DocumentPropertyItem["revision"]="Revision";
		DocumentPropertyItem["appName"]="AppName";
		DocumentPropertyItem["lastPrint"]="LastPrint";
		DocumentPropertyItem["creation"]="Creation";
		DocumentPropertyItem["lastSave"]="LastSave";
		DocumentPropertyItem["category"]="Category";
		DocumentPropertyItem["format"]="Format";
		DocumentPropertyItem["manager"]="Manager";
		DocumentPropertyItem["company"]="Company";
	})(DocumentPropertyItem=ExcelOp.DocumentPropertyItem || (ExcelOp.DocumentPropertyItem={}));
	var SubtotalLocationType;
	(function (SubtotalLocationType) {
		SubtotalLocationType["atTop"]="AtTop";
		SubtotalLocationType["atBottom"]="AtBottom";
		SubtotalLocationType["off"]="Off";
	})(SubtotalLocationType=ExcelOp.SubtotalLocationType || (ExcelOp.SubtotalLocationType={}));
	var PivotLayoutType;
	(function (PivotLayoutType) {
		PivotLayoutType["compact"]="Compact";
		PivotLayoutType["tabular"]="Tabular";
		PivotLayoutType["outline"]="Outline";
	})(PivotLayoutType=ExcelOp.PivotLayoutType || (ExcelOp.PivotLayoutType={}));
	var ProtectionSelectionMode;
	(function (ProtectionSelectionMode) {
		ProtectionSelectionMode["normal"]="Normal";
		ProtectionSelectionMode["unlocked"]="Unlocked";
		ProtectionSelectionMode["none"]="None";
	})(ProtectionSelectionMode=ExcelOp.ProtectionSelectionMode || (ExcelOp.ProtectionSelectionMode={}));
	var PageOrientation;
	(function (PageOrientation) {
		PageOrientation["portrait"]="Portrait";
		PageOrientation["landscape"]="Landscape";
	})(PageOrientation=ExcelOp.PageOrientation || (ExcelOp.PageOrientation={}));
	var PaperType;
	(function (PaperType) {
		PaperType["letter"]="Letter";
		PaperType["letterSmall"]="LetterSmall";
		PaperType["tabloid"]="Tabloid";
		PaperType["ledger"]="Ledger";
		PaperType["legal"]="Legal";
		PaperType["statement"]="Statement";
		PaperType["executive"]="Executive";
		PaperType["a3"]="A3";
		PaperType["a4"]="A4";
		PaperType["a4Small"]="A4Small";
		PaperType["a5"]="A5";
		PaperType["b4"]="B4";
		PaperType["b5"]="B5";
		PaperType["folio"]="Folio";
		PaperType["quatro"]="Quatro";
		PaperType["paper10x14"]="Paper10x14";
		PaperType["paper11x17"]="Paper11x17";
		PaperType["note"]="Note";
		PaperType["envelope9"]="Envelope9";
		PaperType["envelope10"]="Envelope10";
		PaperType["envelope11"]="Envelope11";
		PaperType["envelope12"]="Envelope12";
		PaperType["envelope14"]="Envelope14";
		PaperType["csheet"]="Csheet";
		PaperType["dsheet"]="Dsheet";
		PaperType["esheet"]="Esheet";
		PaperType["envelopeDL"]="EnvelopeDL";
		PaperType["envelopeC5"]="EnvelopeC5";
		PaperType["envelopeC3"]="EnvelopeC3";
		PaperType["envelopeC4"]="EnvelopeC4";
		PaperType["envelopeC6"]="EnvelopeC6";
		PaperType["envelopeC65"]="EnvelopeC65";
		PaperType["envelopeB4"]="EnvelopeB4";
		PaperType["envelopeB5"]="EnvelopeB5";
		PaperType["envelopeB6"]="EnvelopeB6";
		PaperType["envelopeItaly"]="EnvelopeItaly";
		PaperType["envelopeMonarch"]="EnvelopeMonarch";
		PaperType["envelopePersonal"]="EnvelopePersonal";
		PaperType["fanfoldUS"]="FanfoldUS";
		PaperType["fanfoldStdGerman"]="FanfoldStdGerman";
		PaperType["fanfoldLegalGerman"]="FanfoldLegalGerman";
	})(PaperType=ExcelOp.PaperType || (ExcelOp.PaperType={}));
	var ReadingOrder;
	(function (ReadingOrder) {
		ReadingOrder["context"]="Context";
		ReadingOrder["leftToRight"]="LeftToRight";
		ReadingOrder["rightToLeft"]="RightToLeft";
	})(ReadingOrder=ExcelOp.ReadingOrder || (ExcelOp.ReadingOrder={}));
	var BuiltInStyle;
	(function (BuiltInStyle) {
		BuiltInStyle["normal"]="Normal";
		BuiltInStyle["comma"]="Comma";
		BuiltInStyle["currency"]="Currency";
		BuiltInStyle["percent"]="Percent";
		BuiltInStyle["wholeComma"]="WholeComma";
		BuiltInStyle["wholeDollar"]="WholeDollar";
		BuiltInStyle["hlink"]="Hlink";
		BuiltInStyle["hlinkTrav"]="HlinkTrav";
		BuiltInStyle["note"]="Note";
		BuiltInStyle["warningText"]="WarningText";
		BuiltInStyle["emphasis1"]="Emphasis1";
		BuiltInStyle["emphasis2"]="Emphasis2";
		BuiltInStyle["emphasis3"]="Emphasis3";
		BuiltInStyle["sheetTitle"]="SheetTitle";
		BuiltInStyle["heading1"]="Heading1";
		BuiltInStyle["heading2"]="Heading2";
		BuiltInStyle["heading3"]="Heading3";
		BuiltInStyle["heading4"]="Heading4";
		BuiltInStyle["input"]="Input";
		BuiltInStyle["output"]="Output";
		BuiltInStyle["calculation"]="Calculation";
		BuiltInStyle["checkCell"]="CheckCell";
		BuiltInStyle["linkedCell"]="LinkedCell";
		BuiltInStyle["total"]="Total";
		BuiltInStyle["good"]="Good";
		BuiltInStyle["bad"]="Bad";
		BuiltInStyle["neutral"]="Neutral";
		BuiltInStyle["accent1"]="Accent1";
		BuiltInStyle["accent1_20"]="Accent1_20";
		BuiltInStyle["accent1_40"]="Accent1_40";
		BuiltInStyle["accent1_60"]="Accent1_60";
		BuiltInStyle["accent2"]="Accent2";
		BuiltInStyle["accent2_20"]="Accent2_20";
		BuiltInStyle["accent2_40"]="Accent2_40";
		BuiltInStyle["accent2_60"]="Accent2_60";
		BuiltInStyle["accent3"]="Accent3";
		BuiltInStyle["accent3_20"]="Accent3_20";
		BuiltInStyle["accent3_40"]="Accent3_40";
		BuiltInStyle["accent3_60"]="Accent3_60";
		BuiltInStyle["accent4"]="Accent4";
		BuiltInStyle["accent4_20"]="Accent4_20";
		BuiltInStyle["accent4_40"]="Accent4_40";
		BuiltInStyle["accent4_60"]="Accent4_60";
		BuiltInStyle["accent5"]="Accent5";
		BuiltInStyle["accent5_20"]="Accent5_20";
		BuiltInStyle["accent5_40"]="Accent5_40";
		BuiltInStyle["accent5_60"]="Accent5_60";
		BuiltInStyle["accent6"]="Accent6";
		BuiltInStyle["accent6_20"]="Accent6_20";
		BuiltInStyle["accent6_40"]="Accent6_40";
		BuiltInStyle["accent6_60"]="Accent6_60";
		BuiltInStyle["explanatoryText"]="ExplanatoryText";
	})(BuiltInStyle=ExcelOp.BuiltInStyle || (ExcelOp.BuiltInStyle={}));
	var PrintErrorType;
	(function (PrintErrorType) {
		PrintErrorType["asDisplayed"]="AsDisplayed";
		PrintErrorType["blank"]="Blank";
		PrintErrorType["dash"]="Dash";
		PrintErrorType["notAvailable"]="NotAvailable";
	})(PrintErrorType=ExcelOp.PrintErrorType || (ExcelOp.PrintErrorType={}));
	var WorksheetPositionType;
	(function (WorksheetPositionType) {
		WorksheetPositionType["none"]="None";
		WorksheetPositionType["before"]="Before";
		WorksheetPositionType["after"]="After";
		WorksheetPositionType["beginning"]="Beginning";
		WorksheetPositionType["end"]="End";
	})(WorksheetPositionType=ExcelOp.WorksheetPositionType || (ExcelOp.WorksheetPositionType={}));
	var PrintComments;
	(function (PrintComments) {
		PrintComments["noComments"]="NoComments";
		PrintComments["endSheet"]="EndSheet";
		PrintComments["inPlace"]="InPlace";
	})(PrintComments=ExcelOp.PrintComments || (ExcelOp.PrintComments={}));
	var PrintOrder;
	(function (PrintOrder) {
		PrintOrder["downThenOver"]="DownThenOver";
		PrintOrder["overThenDown"]="OverThenDown";
	})(PrintOrder=ExcelOp.PrintOrder || (ExcelOp.PrintOrder={}));
	var PrintMarginUnit;
	(function (PrintMarginUnit) {
		PrintMarginUnit["points"]="Points";
		PrintMarginUnit["inches"]="Inches";
		PrintMarginUnit["centimeters"]="Centimeters";
	})(PrintMarginUnit=ExcelOp.PrintMarginUnit || (ExcelOp.PrintMarginUnit={}));
	var HeaderFooterState;
	(function (HeaderFooterState) {
		HeaderFooterState["default"]="Default";
		HeaderFooterState["firstAndDefault"]="FirstAndDefault";
		HeaderFooterState["oddAndEven"]="OddAndEven";
		HeaderFooterState["firstOddAndEven"]="FirstOddAndEven";
	})(HeaderFooterState=ExcelOp.HeaderFooterState || (ExcelOp.HeaderFooterState={}));
	var AutoFillType;
	(function (AutoFillType) {
		AutoFillType["fillDefault"]="FillDefault";
		AutoFillType["fillCopy"]="FillCopy";
		AutoFillType["fillSeries"]="FillSeries";
		AutoFillType["fillFormats"]="FillFormats";
		AutoFillType["fillValues"]="FillValues";
		AutoFillType["fillDays"]="FillDays";
		AutoFillType["fillWeekdays"]="FillWeekdays";
		AutoFillType["fillMonths"]="FillMonths";
		AutoFillType["fillYears"]="FillYears";
		AutoFillType["linearTrend"]="LinearTrend";
		AutoFillType["growthTrend"]="GrowthTrend";
		AutoFillType["flashFill"]="FlashFill";
	})(AutoFillType=ExcelOp.AutoFillType || (ExcelOp.AutoFillType={}));
	var RangeCopyType;
	(function (RangeCopyType) {
		RangeCopyType["all"]="All";
		RangeCopyType["formulas"]="Formulas";
		RangeCopyType["values"]="Values";
		RangeCopyType["formats"]="Formats";
	})(RangeCopyType=ExcelOp.RangeCopyType || (ExcelOp.RangeCopyType={}));
	var LinkedDataTypeState;
	(function (LinkedDataTypeState) {
		LinkedDataTypeState["none"]="None";
		LinkedDataTypeState["validLinkedData"]="ValidLinkedData";
		LinkedDataTypeState["disambiguationNeeded"]="DisambiguationNeeded";
		LinkedDataTypeState["brokenLinkedData"]="BrokenLinkedData";
		LinkedDataTypeState["fetchingData"]="FetchingData";
	})(LinkedDataTypeState=ExcelOp.LinkedDataTypeState || (ExcelOp.LinkedDataTypeState={}));
	var GeometricShapeType;
	(function (GeometricShapeType) {
		GeometricShapeType["lineInverse"]="LineInverse";
		GeometricShapeType["triangle"]="Triangle";
		GeometricShapeType["rightTriangle"]="RightTriangle";
		GeometricShapeType["rectangle"]="Rectangle";
		GeometricShapeType["diamond"]="Diamond";
		GeometricShapeType["parallelogram"]="Parallelogram";
		GeometricShapeType["trapezoid"]="Trapezoid";
		GeometricShapeType["nonIsoscelesTrapezoid"]="NonIsoscelesTrapezoid";
		GeometricShapeType["pentagon"]="Pentagon";
		GeometricShapeType["hexagon"]="Hexagon";
		GeometricShapeType["heptagon"]="Heptagon";
		GeometricShapeType["octagon"]="Octagon";
		GeometricShapeType["decagon"]="Decagon";
		GeometricShapeType["dodecagon"]="Dodecagon";
		GeometricShapeType["star4"]="Star4";
		GeometricShapeType["star5"]="Star5";
		GeometricShapeType["star6"]="Star6";
		GeometricShapeType["star7"]="Star7";
		GeometricShapeType["star8"]="Star8";
		GeometricShapeType["star10"]="Star10";
		GeometricShapeType["star12"]="Star12";
		GeometricShapeType["star16"]="Star16";
		GeometricShapeType["star24"]="Star24";
		GeometricShapeType["star32"]="Star32";
		GeometricShapeType["roundRectangle"]="RoundRectangle";
		GeometricShapeType["round1Rectangle"]="Round1Rectangle";
		GeometricShapeType["round2SameRectangle"]="Round2SameRectangle";
		GeometricShapeType["round2DiagonalRectangle"]="Round2DiagonalRectangle";
		GeometricShapeType["snipRoundRectangle"]="SnipRoundRectangle";
		GeometricShapeType["snip1Rectangle"]="Snip1Rectangle";
		GeometricShapeType["snip2SameRectangle"]="Snip2SameRectangle";
		GeometricShapeType["snip2DiagonalRectangle"]="Snip2DiagonalRectangle";
		GeometricShapeType["plaque"]="Plaque";
		GeometricShapeType["ellipse"]="Ellipse";
		GeometricShapeType["teardrop"]="Teardrop";
		GeometricShapeType["homePlate"]="HomePlate";
		GeometricShapeType["chevron"]="Chevron";
		GeometricShapeType["pieWedge"]="PieWedge";
		GeometricShapeType["pie"]="Pie";
		GeometricShapeType["blockArc"]="BlockArc";
		GeometricShapeType["donut"]="Donut";
		GeometricShapeType["noSmoking"]="NoSmoking";
		GeometricShapeType["rightArrow"]="RightArrow";
		GeometricShapeType["leftArrow"]="LeftArrow";
		GeometricShapeType["upArrow"]="UpArrow";
		GeometricShapeType["downArrow"]="DownArrow";
		GeometricShapeType["stripedRightArrow"]="StripedRightArrow";
		GeometricShapeType["notchedRightArrow"]="NotchedRightArrow";
		GeometricShapeType["bentUpArrow"]="BentUpArrow";
		GeometricShapeType["leftRightArrow"]="LeftRightArrow";
		GeometricShapeType["upDownArrow"]="UpDownArrow";
		GeometricShapeType["leftUpArrow"]="LeftUpArrow";
		GeometricShapeType["leftRightUpArrow"]="LeftRightUpArrow";
		GeometricShapeType["quadArrow"]="QuadArrow";
		GeometricShapeType["leftArrowCallout"]="LeftArrowCallout";
		GeometricShapeType["rightArrowCallout"]="RightArrowCallout";
		GeometricShapeType["upArrowCallout"]="UpArrowCallout";
		GeometricShapeType["downArrowCallout"]="DownArrowCallout";
		GeometricShapeType["leftRightArrowCallout"]="LeftRightArrowCallout";
		GeometricShapeType["upDownArrowCallout"]="UpDownArrowCallout";
		GeometricShapeType["quadArrowCallout"]="QuadArrowCallout";
		GeometricShapeType["bentArrow"]="BentArrow";
		GeometricShapeType["uturnArrow"]="UturnArrow";
		GeometricShapeType["circularArrow"]="CircularArrow";
		GeometricShapeType["leftCircularArrow"]="LeftCircularArrow";
		GeometricShapeType["leftRightCircularArrow"]="LeftRightCircularArrow";
		GeometricShapeType["curvedRightArrow"]="CurvedRightArrow";
		GeometricShapeType["curvedLeftArrow"]="CurvedLeftArrow";
		GeometricShapeType["curvedUpArrow"]="CurvedUpArrow";
		GeometricShapeType["curvedDownArrow"]="CurvedDownArrow";
		GeometricShapeType["swooshArrow"]="SwooshArrow";
		GeometricShapeType["cube"]="Cube";
		GeometricShapeType["can"]="Can";
		GeometricShapeType["lightningBolt"]="LightningBolt";
		GeometricShapeType["heart"]="Heart";
		GeometricShapeType["sun"]="Sun";
		GeometricShapeType["moon"]="Moon";
		GeometricShapeType["smileyFace"]="SmileyFace";
		GeometricShapeType["irregularSeal1"]="IrregularSeal1";
		GeometricShapeType["irregularSeal2"]="IrregularSeal2";
		GeometricShapeType["foldedCorner"]="FoldedCorner";
		GeometricShapeType["bevel"]="Bevel";
		GeometricShapeType["frame"]="Frame";
		GeometricShapeType["halfFrame"]="HalfFrame";
		GeometricShapeType["corner"]="Corner";
		GeometricShapeType["diagonalStripe"]="DiagonalStripe";
		GeometricShapeType["chord"]="Chord";
		GeometricShapeType["arc"]="Arc";
		GeometricShapeType["leftBracket"]="LeftBracket";
		GeometricShapeType["rightBracket"]="RightBracket";
		GeometricShapeType["leftBrace"]="LeftBrace";
		GeometricShapeType["rightBrace"]="RightBrace";
		GeometricShapeType["bracketPair"]="BracketPair";
		GeometricShapeType["bracePair"]="BracePair";
		GeometricShapeType["callout1"]="Callout1";
		GeometricShapeType["callout2"]="Callout2";
		GeometricShapeType["callout3"]="Callout3";
		GeometricShapeType["accentCallout1"]="AccentCallout1";
		GeometricShapeType["accentCallout2"]="AccentCallout2";
		GeometricShapeType["accentCallout3"]="AccentCallout3";
		GeometricShapeType["borderCallout1"]="BorderCallout1";
		GeometricShapeType["borderCallout2"]="BorderCallout2";
		GeometricShapeType["borderCallout3"]="BorderCallout3";
		GeometricShapeType["accentBorderCallout1"]="AccentBorderCallout1";
		GeometricShapeType["accentBorderCallout2"]="AccentBorderCallout2";
		GeometricShapeType["accentBorderCallout3"]="AccentBorderCallout3";
		GeometricShapeType["wedgeRectCallout"]="WedgeRectCallout";
		GeometricShapeType["wedgeRRectCallout"]="WedgeRRectCallout";
		GeometricShapeType["wedgeEllipseCallout"]="WedgeEllipseCallout";
		GeometricShapeType["cloudCallout"]="CloudCallout";
		GeometricShapeType["cloud"]="Cloud";
		GeometricShapeType["ribbon"]="Ribbon";
		GeometricShapeType["ribbon2"]="Ribbon2";
		GeometricShapeType["ellipseRibbon"]="EllipseRibbon";
		GeometricShapeType["ellipseRibbon2"]="EllipseRibbon2";
		GeometricShapeType["leftRightRibbon"]="LeftRightRibbon";
		GeometricShapeType["verticalScroll"]="VerticalScroll";
		GeometricShapeType["horizontalScroll"]="HorizontalScroll";
		GeometricShapeType["wave"]="Wave";
		GeometricShapeType["doubleWave"]="DoubleWave";
		GeometricShapeType["plus"]="Plus";
		GeometricShapeType["flowChartProcess"]="FlowChartProcess";
		GeometricShapeType["flowChartDecision"]="FlowChartDecision";
		GeometricShapeType["flowChartInputOutput"]="FlowChartInputOutput";
		GeometricShapeType["flowChartPredefinedProcess"]="FlowChartPredefinedProcess";
		GeometricShapeType["flowChartInternalStorage"]="FlowChartInternalStorage";
		GeometricShapeType["flowChartDocument"]="FlowChartDocument";
		GeometricShapeType["flowChartMultidocument"]="FlowChartMultidocument";
		GeometricShapeType["flowChartTerminator"]="FlowChartTerminator";
		GeometricShapeType["flowChartPreparation"]="FlowChartPreparation";
		GeometricShapeType["flowChartManualInput"]="FlowChartManualInput";
		GeometricShapeType["flowChartManualOperation"]="FlowChartManualOperation";
		GeometricShapeType["flowChartConnector"]="FlowChartConnector";
		GeometricShapeType["flowChartPunchedCard"]="FlowChartPunchedCard";
		GeometricShapeType["flowChartPunchedTape"]="FlowChartPunchedTape";
		GeometricShapeType["flowChartSummingJunction"]="FlowChartSummingJunction";
		GeometricShapeType["flowChartOr"]="FlowChartOr";
		GeometricShapeType["flowChartCollate"]="FlowChartCollate";
		GeometricShapeType["flowChartSort"]="FlowChartSort";
		GeometricShapeType["flowChartExtract"]="FlowChartExtract";
		GeometricShapeType["flowChartMerge"]="FlowChartMerge";
		GeometricShapeType["flowChartOfflineStorage"]="FlowChartOfflineStorage";
		GeometricShapeType["flowChartOnlineStorage"]="FlowChartOnlineStorage";
		GeometricShapeType["flowChartMagneticTape"]="FlowChartMagneticTape";
		GeometricShapeType["flowChartMagneticDisk"]="FlowChartMagneticDisk";
		GeometricShapeType["flowChartMagneticDrum"]="FlowChartMagneticDrum";
		GeometricShapeType["flowChartDisplay"]="FlowChartDisplay";
		GeometricShapeType["flowChartDelay"]="FlowChartDelay";
		GeometricShapeType["flowChartAlternateProcess"]="FlowChartAlternateProcess";
		GeometricShapeType["flowChartOffpageConnector"]="FlowChartOffpageConnector";
		GeometricShapeType["actionButtonBlank"]="ActionButtonBlank";
		GeometricShapeType["actionButtonHome"]="ActionButtonHome";
		GeometricShapeType["actionButtonHelp"]="ActionButtonHelp";
		GeometricShapeType["actionButtonInformation"]="ActionButtonInformation";
		GeometricShapeType["actionButtonForwardNext"]="ActionButtonForwardNext";
		GeometricShapeType["actionButtonBackPrevious"]="ActionButtonBackPrevious";
		GeometricShapeType["actionButtonEnd"]="ActionButtonEnd";
		GeometricShapeType["actionButtonBeginning"]="ActionButtonBeginning";
		GeometricShapeType["actionButtonReturn"]="ActionButtonReturn";
		GeometricShapeType["actionButtonDocument"]="ActionButtonDocument";
		GeometricShapeType["actionButtonSound"]="ActionButtonSound";
		GeometricShapeType["actionButtonMovie"]="ActionButtonMovie";
		GeometricShapeType["gear6"]="Gear6";
		GeometricShapeType["gear9"]="Gear9";
		GeometricShapeType["funnel"]="Funnel";
		GeometricShapeType["mathPlus"]="MathPlus";
		GeometricShapeType["mathMinus"]="MathMinus";
		GeometricShapeType["mathMultiply"]="MathMultiply";
		GeometricShapeType["mathDivide"]="MathDivide";
		GeometricShapeType["mathEqual"]="MathEqual";
		GeometricShapeType["mathNotEqual"]="MathNotEqual";
		GeometricShapeType["cornerTabs"]="CornerTabs";
		GeometricShapeType["squareTabs"]="SquareTabs";
		GeometricShapeType["plaqueTabs"]="PlaqueTabs";
		GeometricShapeType["chartX"]="ChartX";
		GeometricShapeType["chartStar"]="ChartStar";
		GeometricShapeType["chartPlus"]="ChartPlus";
	})(GeometricShapeType=ExcelOp.GeometricShapeType || (ExcelOp.GeometricShapeType={}));
	var ConnectorType;
	(function (ConnectorType) {
		ConnectorType["straight"]="Straight";
		ConnectorType["elbow"]="Elbow";
		ConnectorType["curve"]="Curve";
	})(ConnectorType=ExcelOp.ConnectorType || (ExcelOp.ConnectorType={}));
	var ContentType;
	(function (ContentType) {
		ContentType["plain"]="Plain";
	})(ContentType=ExcelOp.ContentType || (ExcelOp.ContentType={}));
	var SpecialCellType;
	(function (SpecialCellType) {
		SpecialCellType["conditionalFormats"]="ConditionalFormats";
		SpecialCellType["dataValidations"]="DataValidations";
		SpecialCellType["blanks"]="Blanks";
		SpecialCellType["constants"]="Constants";
		SpecialCellType["formulas"]="Formulas";
		SpecialCellType["sameConditionalFormat"]="SameConditionalFormat";
		SpecialCellType["sameDataValidation"]="SameDataValidation";
		SpecialCellType["visible"]="Visible";
	})(SpecialCellType=ExcelOp.SpecialCellType || (ExcelOp.SpecialCellType={}));
	var SpecialCellValueType;
	(function (SpecialCellValueType) {
		SpecialCellValueType["all"]="All";
		SpecialCellValueType["errors"]="Errors";
		SpecialCellValueType["errorsLogical"]="ErrorsLogical";
		SpecialCellValueType["errorsNumbers"]="ErrorsNumbers";
		SpecialCellValueType["errorsText"]="ErrorsText";
		SpecialCellValueType["errorsLogicalNumber"]="ErrorsLogicalNumber";
		SpecialCellValueType["errorsLogicalText"]="ErrorsLogicalText";
		SpecialCellValueType["errorsNumberText"]="ErrorsNumberText";
		SpecialCellValueType["logical"]="Logical";
		SpecialCellValueType["logicalNumbers"]="LogicalNumbers";
		SpecialCellValueType["logicalText"]="LogicalText";
		SpecialCellValueType["logicalNumbersText"]="LogicalNumbersText";
		SpecialCellValueType["numbers"]="Numbers";
		SpecialCellValueType["numbersText"]="NumbersText";
		SpecialCellValueType["text"]="Text";
	})(SpecialCellValueType=ExcelOp.SpecialCellValueType || (ExcelOp.SpecialCellValueType={}));
	var Placement;
	(function (Placement) {
		Placement["twoCell"]="TwoCell";
		Placement["oneCell"]="OneCell";
		Placement["absolute"]="Absolute";
	})(Placement=ExcelOp.Placement || (ExcelOp.Placement={}));
	var FillPattern;
	(function (FillPattern) {
		FillPattern["none"]="None";
		FillPattern["solid"]="Solid";
		FillPattern["gray50"]="Gray50";
		FillPattern["gray75"]="Gray75";
		FillPattern["gray25"]="Gray25";
		FillPattern["horizontal"]="Horizontal";
		FillPattern["vertical"]="Vertical";
		FillPattern["down"]="Down";
		FillPattern["up"]="Up";
		FillPattern["checker"]="Checker";
		FillPattern["semiGray75"]="SemiGray75";
		FillPattern["lightHorizontal"]="LightHorizontal";
		FillPattern["lightVertical"]="LightVertical";
		FillPattern["lightDown"]="LightDown";
		FillPattern["lightUp"]="LightUp";
		FillPattern["grid"]="Grid";
		FillPattern["crissCross"]="CrissCross";
		FillPattern["gray16"]="Gray16";
		FillPattern["gray8"]="Gray8";
		FillPattern["linearGradient"]="LinearGradient";
		FillPattern["rectangularGradient"]="RectangularGradient";
	})(FillPattern=ExcelOp.FillPattern || (ExcelOp.FillPattern={}));
	var ShapeTextHorizontalAlignType;
	(function (ShapeTextHorizontalAlignType) {
		ShapeTextHorizontalAlignType["left"]="Left";
		ShapeTextHorizontalAlignType["center"]="Center";
		ShapeTextHorizontalAlignType["right"]="Right";
		ShapeTextHorizontalAlignType["justify"]="Justify";
		ShapeTextHorizontalAlignType["justifyLow"]="JustifyLow";
		ShapeTextHorizontalAlignType["distributed"]="Distributed";
		ShapeTextHorizontalAlignType["thaiDistributed"]="ThaiDistributed";
		ShapeTextHorizontalAlignType["shapeTextHorizontalAlignType_MaxEnumIDs"]="ShapeTextHorizontalAlignType_MaxEnumIDs";
	})(ShapeTextHorizontalAlignType=ExcelOp.ShapeTextHorizontalAlignType || (ExcelOp.ShapeTextHorizontalAlignType={}));
	var ShapeTextVerticalAlignType;
	(function (ShapeTextVerticalAlignType) {
		ShapeTextVerticalAlignType["top"]="Top";
		ShapeTextVerticalAlignType["middle"]="Middle";
		ShapeTextVerticalAlignType["bottom"]="Bottom";
		ShapeTextVerticalAlignType["justified"]="Justified";
		ShapeTextVerticalAlignType["distributed"]="Distributed";
		ShapeTextVerticalAlignType["shapeTextVerticalAlignType_MaxEnumIDs"]="ShapeTextVerticalAlignType_MaxEnumIDs";
	})(ShapeTextVerticalAlignType=ExcelOp.ShapeTextVerticalAlignType || (ExcelOp.ShapeTextVerticalAlignType={}));
	var ShapeTextVertOverflowType;
	(function (ShapeTextVertOverflowType) {
		ShapeTextVertOverflowType["overflow"]="Overflow";
		ShapeTextVertOverflowType["ellipsis"]="Ellipsis";
		ShapeTextVertOverflowType["clip"]="Clip";
		ShapeTextVertOverflowType["shapeTextVertOverflowType_MaxEnumIDs"]="ShapeTextVertOverflowType_MaxEnumIDs";
	})(ShapeTextVertOverflowType=ExcelOp.ShapeTextVertOverflowType || (ExcelOp.ShapeTextVertOverflowType={}));
	var ShapeTextHorzOverflowType;
	(function (ShapeTextHorzOverflowType) {
		ShapeTextHorzOverflowType["overflow"]="Overflow";
		ShapeTextHorzOverflowType["clip"]="Clip";
		ShapeTextHorzOverflowType["shapeTextHorzOverflowType_MaxEnumIDs"]="ShapeTextHorzOverflowType_MaxEnumIDs";
	})(ShapeTextHorzOverflowType=ExcelOp.ShapeTextHorzOverflowType || (ExcelOp.ShapeTextHorzOverflowType={}));
	var ShapeTextReadingOrder;
	(function (ShapeTextReadingOrder) {
		ShapeTextReadingOrder["ltr"]="LTR";
		ShapeTextReadingOrder["rtl"]="RTL";
	})(ShapeTextReadingOrder=ExcelOp.ShapeTextReadingOrder || (ExcelOp.ShapeTextReadingOrder={}));
	var ShapeTextOrientationType;
	(function (ShapeTextOrientationType) {
		ShapeTextOrientationType["horizontal"]="Horizontal";
		ShapeTextOrientationType["vertical"]="Vertical";
		ShapeTextOrientationType["vertical270"]="Vertical270";
		ShapeTextOrientationType["wordArtVertical"]="WordArtVertical";
		ShapeTextOrientationType["eastAsianVertical"]="EastAsianVertical";
		ShapeTextOrientationType["mongolianVertical"]="MongolianVertical";
		ShapeTextOrientationType["wordArtVerticalRTL"]="WordArtVerticalRTL";
		ShapeTextOrientationType["shapeTextOrientationType_MaxEnumIDs"]="ShapeTextOrientationType_MaxEnumIDs";
	})(ShapeTextOrientationType=ExcelOp.ShapeTextOrientationType || (ExcelOp.ShapeTextOrientationType={}));
	var ShapeAutoSize;
	(function (ShapeAutoSize) {
		ShapeAutoSize["autoSizeNone"]="AutoSizeNone";
		ShapeAutoSize["autoSizeTextToFitShape"]="AutoSizeTextToFitShape";
		ShapeAutoSize["autoSizeShapeToFitText"]="AutoSizeShapeToFitText";
		ShapeAutoSize["autoSizeMixed"]="AutoSizeMixed";
	})(ShapeAutoSize=ExcelOp.ShapeAutoSize || (ExcelOp.ShapeAutoSize={}));
	var CloseBehavior;
	(function (CloseBehavior) {
		CloseBehavior["save"]="Save";
		CloseBehavior["skipSave"]="SkipSave";
	})(CloseBehavior=ExcelOp.CloseBehavior || (ExcelOp.CloseBehavior={}));
	var SaveBehavior;
	(function (SaveBehavior) {
		SaveBehavior["save"]="Save";
		SaveBehavior["prompt"]="Prompt";
	})(SaveBehavior=ExcelOp.SaveBehavior || (ExcelOp.SaveBehavior={}));
	var SlicerSortType;
	(function (SlicerSortType) {
		SlicerSortType["dataSourceOrder"]="DataSourceOrder";
		SlicerSortType["ascending"]="Ascending";
		SlicerSortType["descending"]="Descending";
	})(SlicerSortType=ExcelOp.SlicerSortType || (ExcelOp.SlicerSortType={}));
	var RibbonTab;
	(function (RibbonTab) {
		RibbonTab["others"]="Others";
		RibbonTab["home"]="Home";
		RibbonTab["insert"]="Insert";
		RibbonTab["draw"]="Draw";
		RibbonTab["pageLayout"]="PageLayout";
		RibbonTab["formulas"]="Formulas";
		RibbonTab["data"]="Data";
		RibbonTab["review"]="Review";
		RibbonTab["view"]="View";
		RibbonTab["developer"]="Developer";
		RibbonTab["addIns"]="AddIns";
		RibbonTab["help"]="Help";
	})(RibbonTab=ExcelOp.RibbonTab || (ExcelOp.RibbonTab={}));
	var FunctionResult=(function (_super) {
		__extends(FunctionResult, _super);
		function FunctionResult() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(FunctionResult.prototype, "_className", {
			get: function () {
				return "FunctionResult<T>";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(FunctionResult.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["error", "value"];
			},
			enumerable: true,
			configurable: true
		});
		FunctionResult.prototype.retrieve=function () {
			var select=[];
			for (var _i=0; _i < arguments.length; _i++) {
				select[_i]=arguments[_i];
			}
			return _invokeRetrieve(this, select);
		};
		FunctionResult.prototype.toJSON=function () {
			return {};
		};
		return FunctionResult;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.FunctionResult=FunctionResult;
	var Functions=(function (_super) {
		__extends(Functions, _super);
		function Functions() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(Functions.prototype, "_className", {
			get: function () {
				return "Functions";
			},
			enumerable: true,
			configurable: true
		});
		Functions.prototype.toJSON=function () {
			return {};
		};
		return Functions;
	}(OfficeExtension.ClientObjectBase));
	ExcelOp.Functions=Functions;
})(ExcelOp || (ExcelOp={}));

