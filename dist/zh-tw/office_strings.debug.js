/* Version: 16.0.11030.10000 */

Type.registerNamespace("Strings");
Strings.OfficeOM = function()
{
};
Strings.OfficeOM.registerClass("Strings.OfficeOM");
Strings.OfficeOM.L_APICallFailed = "API 呼叫失敗";
Strings.OfficeOM.L_APINotSupported = "不支援 API";
Strings.OfficeOM.L_ActivityLimitReached = "已達活動限制。";
Strings.OfficeOM.L_AddBindingFromPromptDefaultText = "請選擇一個選項。";
Strings.OfficeOM.L_AddinIsAlreadyRequestingToken = "增益集已經要求存取權杖。";
Strings.OfficeOM.L_AddinIsAlreadyRequestingTokenMessage = "作業失敗，因為此增益集已經要求存取權杖。";
Strings.OfficeOM.L_ApiNotFoundDetails = "方法或屬性 {0} 是 {1} 需求集合的一部分，您的 {2} 版本無法使用它。";
Strings.OfficeOM.L_AppNameNotExist = "{0} 的增益集名稱不存在。";
Strings.OfficeOM.L_AppNotExistInitializeNotCalled = "應用程式 {0} 不存在。沒有呼叫 Microsoft.Office.WebExtension.initialize(reason)。";
Strings.OfficeOM.L_AttemptingToSetReadOnlyProperty = "正在嘗試設定唯讀屬性 '{0}'。";
Strings.OfficeOM.L_BadSelectorString = "傳入選取器之字串的格式不適當或不受支援。";
Strings.OfficeOM.L_BindingCreationError = "繫結建立錯誤";
Strings.OfficeOM.L_BindingNotExist = "指定的繫結不存在。";
Strings.OfficeOM.L_BindingToMultipleSelection = "不支援選取多個非連續項目。";
Strings.OfficeOM.L_BrowserAPINotSupported = "此瀏覽器不支援要求的 API。";
Strings.OfficeOM.L_CallbackNotAFunction = "回撥必須是類型函數、曾是類型 {0}。"
Strings.OfficeOM.L_CannotApplyPropertyThroughSetMethod = "無法透過 \"object.set\" 方法套用對屬性 '{0}' 所做的變更。";
Strings.OfficeOM.L_CannotNavigateTo = "物件位於不支援導覽的位置。";
Strings.OfficeOM.L_CannotRegisterEvent = "無法登錄事件處理常式。";
Strings.OfficeOM.L_CannotWriteToSelection = "無法寫入至目前的選取範圍。";
Strings.OfficeOM.L_CellDataAmountBeyondLimits = "附註: 我們建議表格中的儲存格數量低於 20,000。";
Strings.OfficeOM.L_CellFormatAmountBeyondLimits = "附註: 我們建議由格式 API 呼叫設定的格式組低於 100。";
Strings.OfficeOM.L_CloseFileBeforeRetrieve = "擷取另一個之前，呼叫目前的檔案中的 closeAsync。";
Strings.OfficeOM.L_CoercionTypeNotMatchBinding = "指定的強制型轉類型與此繫結類型不相容。";
Strings.OfficeOM.L_CoercionTypeNotSupported = "不支援指定的強制型轉類型。";
Strings.OfficeOM.L_ColIndexOutOfRange = "欄索引值不在允許的範圍內。請用小於欄數的值 (0 或更大)。";
Strings.OfficeOM.L_ConnectionFailureWithDetails = "要求失敗，狀態碼為 {0}、錯誤碼為 {1}，且錯誤訊息如下: {2}";
Strings.OfficeOM.L_ConnectionFailureWithStatus = "要求失敗，狀態碼為 {0}。";
Strings.OfficeOM.L_CustomFunctionDefinitionMissing = "具有此名稱且代表函數定義的屬性必須存在於 Excel.Script.CustomFunctions。";
Strings.OfficeOM.L_CustomFunctionImplementationMissing = "Excel.Script.CustomFunctions 上具有此名稱且代表函數定義的屬性必須包含實作該函數的 'call' 屬性。";
Strings.OfficeOM.L_CustomFunctionNameCannotSplit = "函數名稱必須包含非空白命名空間與非空白簡短名稱。";
Strings.OfficeOM.L_CustomFunctionNameContainsBadChars = "函數名稱只能包含字母、數字、底線與句號。";
Strings.OfficeOM.L_CustomXmlError = "自訂 XML 錯誤。";
Strings.OfficeOM.L_CustomXmlExceedQuotaMessage = "XPath 的選取項目上限為 1024 個項目。";
Strings.OfficeOM.L_CustomXmlExceedQuotaName = "已達選取項目上限";
Strings.OfficeOM.L_CustomXmlNodeNotFound = "找不到指定的節點。";
Strings.OfficeOM.L_CustomXmlOutOfDateMessage = "資料已過期。請重新擷取物件。";
Strings.OfficeOM.L_CustomXmlOutOfDateName = "資料不是最新的";
Strings.OfficeOM.L_DataNotMatchBindingSize = "提供的資料物件與目前選擇的大小不符。";
Strings.OfficeOM.L_DataNotMatchBindingType = "指定的資料物件與繫結類型不相容。";
Strings.OfficeOM.L_DataNotMatchCoercionType = "指定的資料物件類型與目前的選擇不相容。";
Strings.OfficeOM.L_DataNotMatchSelection = "提供的資料物件與目前選取範圍的圖形或維度不相容。";
Strings.OfficeOM.L_DataReadError = "資料讀取錯誤";
Strings.OfficeOM.L_DataStale = "資料不是最新的";
Strings.OfficeOM.L_DataWriteError = "資料寫入錯誤";
Strings.OfficeOM.L_DataWriteReminder = "資料寫入提醒";
Strings.OfficeOM.L_DialogAddressNotTrusted = "URL 的網域沒有包含在資訊清單的 AppDomains 元素中。";
Strings.OfficeOM.L_DialogAlreadyOpened = "作業失敗，因為這個增益集已有使用中的對話方塊。";
Strings.OfficeOM.L_DialogInvalidScheme = "不支援 URL 配置。請改為使用 HTTPS。";
Strings.OfficeOM.L_DialogNavigateError = "對話方塊導覽錯誤";
Strings.OfficeOM.L_DialogOK = "確定";
Strings.OfficeOM.L_DialogRequireHTTPS = "不支援 HTTP 通訊協定。請改用 HTTPS";
Strings.OfficeOM.L_DisplayDialogError = "顯示對話方塊錯誤";
Strings.OfficeOM.L_DocumentReadOnly = "目前的文件模式不允許執行要求的作業。";
Strings.OfficeOM.L_ElementMissing = "無法設定表格儲存格格式，因為缺少某些參數值。請再次檢查參數，然後重試一次。";
Strings.OfficeOM.L_EventHandlerAdditionFailed = "無法新增事件處理常式。";
Strings.OfficeOM.L_EventHandlerNotExist = "此繫結找不到指定的事件處理常式。";
Strings.OfficeOM.L_EventHandlerRemovalFailed = "無法移除事件處理常式。";
Strings.OfficeOM.L_EventRegistrationError = "事件註冊錯誤";
Strings.OfficeOM.L_FileTypeNotSupported = "指定的檔案類型不受支援。";
Strings.OfficeOM.L_FormatValueOutOfRange = "值超出允許的範圍。";
Strings.OfficeOM.L_FormattingReminder = "格式設定提醒";
Strings.OfficeOM.L_FunctionCallFailed = "函數 {0} 呼叫失敗，錯誤碼: {1}。";
Strings.OfficeOM.L_GetDataIsTooLarge = "要求的資料集太大。";
Strings.OfficeOM.L_GetDataParametersConflict = "指定的參數相衝突。";
Strings.OfficeOM.L_GetSelectionNotSupported = "不支援目前的選取範圍。";
Strings.OfficeOM.L_HostError = "主機錯誤";
Strings.OfficeOM.L_InValidOptionalArgument = "無效的選擇性引數";
Strings.OfficeOM.L_IndexOutOfRange = "索引超出範圍。";
Strings.OfficeOM.L_InitializeNotReady = "Office.js 尚未完全載入。請稍後再試，或將您的初始化程式碼新增到 Office.initialize 函數。";
Strings.OfficeOM.L_InternalError = "內部錯誤";
Strings.OfficeOM.L_InternalErrorDescription = "發生內部錯誤。";
Strings.OfficeOM.L_InvalidAPICall = "無效的 API 呼叫";
Strings.OfficeOM.L_InvalidApiArgumentsMessage = "輸入引數無效。";
Strings.OfficeOM.L_InvalidApiCallInContext = "目前的內容中發生無效的 API 呼叫。";
Strings.OfficeOM.L_InvalidArgument = "引數 '{0}' 不適用於這種情況、遺失，或格式不正確。";
Strings.OfficeOM.L_InvalidArgumentGeneric = "傳遞到函數的引數在此情況下不適用、已遺失或格式不正確。";
Strings.OfficeOM.L_InvalidBinding = "無效的繫結";
Strings.OfficeOM.L_InvalidBindingError = "無效的繫結錯誤";
Strings.OfficeOM.L_InvalidBindingOperation = "無效的繫結作業";
Strings.OfficeOM.L_InvalidCellsValue = "一或多個儲存格參數有不允許使用的值。請再次檢查值，然後重試一次。";
Strings.OfficeOM.L_InvalidCoercion = "無效的強制型轉類型";
Strings.OfficeOM.L_InvalidColumnsForBinding = "指定的欄無效。";
Strings.OfficeOM.L_InvalidDataFormat = "指定的資料物件格式無效。";
Strings.OfficeOM.L_InvalidDataObject = "無效的資料物件";
Strings.OfficeOM.L_InvalidFormat = "無效的格式錯誤";
Strings.OfficeOM.L_InvalidFormatValue = "一或多個格式參數有不允許使用的值。請再次檢查值，然後重試一次。";
Strings.OfficeOM.L_InvalidGetColumns = "指定的欄無效。";
Strings.OfficeOM.L_InvalidGetRowColumnCounts = "指定的 rowCount 或 columnCount 值無效。";
Strings.OfficeOM.L_InvalidGetRows = "指定的列無效。";
Strings.OfficeOM.L_InvalidGetStartRowColumn = "指定的 startRow 或 startColumn 值無效。";
Strings.OfficeOM.L_InvalidGrant = "缺少預先授權。";
Strings.OfficeOM.L_InvalidGrantMessage = "缺少此增益集的授權。";
Strings.OfficeOM.L_InvalidNamedItemForBindingType = "指定的繫結類型與提供的命名項目不相容。";
Strings.OfficeOM.L_InvalidNode = "無效的節點";
Strings.OfficeOM.L_InvalidObjectPath = '物件路徑 \'{0}\' 不適用於您嘗試處理的狀況。如果您正在跨多個 "context.sync" 呼叫來使用物件，且處於 ".run" 批次的循序執行之外，則請使用 "context.trackedObjects.add()" 和 "context.trackedObjects.remove()" 方法來管理物件的存留期。';
Strings.OfficeOM.L_InvalidOperationInCellEditMode = "Excel 處於儲存格編輯模式。請按 ENTER 或 TAB 或是選取另一個儲存格以結束編輯模式，然後再試一次。";
Strings.OfficeOM.L_InvalidOrTimedOutSession = "無效或逾時的工作階段";
Strings.OfficeOM.L_InvalidOrTimedOutSessionMessage = "您的 Office Online 工作階段已逾時或無效。若要繼續，請重新整理頁面。";
Strings.OfficeOM.L_InvalidParameters = "函數 {0} 有無效的參數。";
Strings.OfficeOM.L_InvalidReadForBlankRow = "指定的列是空白的。";
Strings.OfficeOM.L_InvalidRequestContext = "無法使用跨不同要求內容的物件。";
Strings.OfficeOM.L_InvalidResourceUrl = "提供的應用程式資源 URL 無效。";
Strings.OfficeOM.L_InvalidResourceUrlMessage = "資訊清單中指定的資源 URL 無效。";
Strings.OfficeOM.L_InvalidSSOAddinMessage = "此增益集不支援身分識別 API。";
Strings.OfficeOM.L_InvalidSelectionForBindingType = "無法以目前的選取範圍與指定的繫結類型建立繫結。";
Strings.OfficeOM.L_InvalidSetColumns = "指定的欄無效。";
Strings.OfficeOM.L_InvalidSetRows = "指定的列無效。";
Strings.OfficeOM.L_InvalidSetStartRowColumn = "指定的 startRow 或 startColumn 值無效。";
Strings.OfficeOM.L_InvalidTableOptionValue = "一或多個 tableOptions 參數有不允許使用的值，請再次檢查值，然後重試一次。";
Strings.OfficeOM.L_InvalidValue = "無效值";
Strings.OfficeOM.L_MemoryLimit = "已超過記憶體限制";
Strings.OfficeOM.L_MissingParameter = "缺少參數。";
Strings.OfficeOM.L_MissingRequiredArguments = "遺失部分必要引數";
Strings.OfficeOM.L_MultipleNamedItemFound = "找到多個同名稱的物件。";
Strings.OfficeOM.L_NamedItemNotFound = "命名項目不存在。";
Strings.OfficeOM.L_NavOutOfBound = "作業失敗，因為索引超出範圍。";
Strings.OfficeOM.L_NetworkProblem = "網路問題";
Strings.OfficeOM.L_NetworkProblemRetrieveFile = "網路發生問題，無法擷取檔案。";
Strings.OfficeOM.L_NewWindowCrossZone = "您瀏覽器的安全性設定導致我們無法建立對話方塊。請嘗試其他瀏覽器，或{0}以讓 '{1}' 和網址列顯示的網域位於相同的安全性區域。";
Strings.OfficeOM.L_NewWindowCrossZoneConfigureBrowserLink = "設定您的瀏覽器";
Strings.OfficeOM.L_NewWindowCrossZoneErrorString = "瀏覽器限制導致我們無法建立對話方塊。對話方塊的網域和增益集主機的網域不是位於相同的安全性區域。";
Strings.OfficeOM.L_NoCapability = "您沒有足夠的權限可執行此動作。";
Strings.OfficeOM.L_NonUniformPartialGetNotSupported = "表格含有合併儲存格時，座標參數不能與強制型轉類型表格共同使用。";
Strings.OfficeOM.L_NonUniformPartialSetNotSupported = "表格含有合併儲存格時，座標參數不能與強制型轉類型表格共同使用。";
Strings.OfficeOM.L_NotImplemented = "函數 {0} 未實作。";
Strings.OfficeOM.L_NotSupported = "不支援函數 {0}。";
Strings.OfficeOM.L_NotSupportedBindingType = "不支援指定的繫結類型 {0}。";
Strings.OfficeOM.L_NotSupportedEventType = "不支援指定的事件類型 {0}。";
Strings.OfficeOM.L_OperationCancelledError = "作業已取消";
Strings.OfficeOM.L_OperationCancelledErrorMessage = "使用者已取消作業。";
Strings.OfficeOM.L_OperationNotSupported = "不支援此作業。";
Strings.OfficeOM.L_OperationNotSupportedOnMatrixData = "選取的內容需要使用表格格式。請將資料格式化為表格，然後再試一次。";
Strings.OfficeOM.L_OperationNotSupportedOnThisBindingType = "不支援此繫結類型的作業。";
Strings.OfficeOM.L_OsfControlTypeNotSupported = "OsfControl 類型不受支援。";
Strings.OfficeOM.L_OutOfRange = "超出範圍";
Strings.OfficeOM.L_OverwriteWorksheetData = "設定作業失敗，因為提供的資料物件會覆寫資料或將資料移位。";
Strings.OfficeOM.L_PermissionDenied = "權限遭拒";
Strings.OfficeOM.L_PropertyDoesNotExist = "物件中沒有屬性 '{0}'。";
Strings.OfficeOM.L_PropertyNotLoaded = "無法使用屬性 '{0}'。在讀取屬性值之前，呼叫包含物件上的 load 方法，並在相關的要求內容上呼叫 \"context.sync()\"。";
Strings.OfficeOM.L_ReadSettingsError = "讀取設定錯誤";
Strings.OfficeOM.L_RedundantCallbackSpecification = "回呼無法在引數清單和選用物件中指定。";
Strings.OfficeOM.L_RequestTimeout = "呼叫執行耗時過長。";
Strings.OfficeOM.L_RequestTokenUnavailable = "已節流此 API 減緩呼叫頻率。";
Strings.OfficeOM.L_RowIndexOutOfRange = "列索引值不在允許的範圍內。請用小於列數的值 (0 或更大)。";
Strings.OfficeOM.L_RunMustReturnPromise = '傳遞到 ".run" 方法的批次函式沒有傳回承諾。函式必須傳回承諾，以便在批次作業完成時可釋放所有自動追蹤的物件。一般而言，您以透過來自 "context.sync()" 的回應來傳回承諾。';
Strings.OfficeOM.L_SSOClientError = "來自 Office 的驗證要求發生錯誤。";
Strings.OfficeOM.L_SSOClientErrorMessage = "用戶端發生未預期的錯誤。";
Strings.OfficeOM.L_SSOConnectionLostError = "連線在登入程序期間已經中斷。";
Strings.OfficeOM.L_SSOConnectionLostErrorMessage = "連線在登入程序期間已經中斷，因此使用者可能無法登入。這可能是使用者的瀏覽器設定 (例如安全性區域) 所致。";
Strings.OfficeOM.L_SSOServerError = "驗證提供者時發生錯誤。";
Strings.OfficeOM.L_SSOServerErrorMessage = "伺服器發生未預期的錯誤。";
Strings.OfficeOM.L_SSOUnsupportedPlatform = "此平台不支援 API。";
Strings.OfficeOM.L_SSOUserConsentNotSupportedByCurrentAddinCategory = "此增益集不支援使用者同意。";
Strings.OfficeOM.L_SSOUserConsentNotSupportedByCurrentAddinCategoryMessage = "作業失敗，因為此增益集在此類別中不支援使用者同意";
Strings.OfficeOM.L_SaveSettingsError = "儲存設定錯誤";
Strings.OfficeOM.L_SelectionCannotBound = "無法繫結到目前的選取範圍。";
Strings.OfficeOM.L_SelectionNotSupportCoercionType = "目前的選取範圍與指定的強制型轉類型不相容。";
Strings.OfficeOM.L_SetDataIsTooLarge = "指定的資料物件太大。";
Strings.OfficeOM.L_SetDataParametersConflict = "指定的參數相衝突。";
Strings.OfficeOM.L_SettingNameNotExist = "指定的設定名稱不存在。";
Strings.OfficeOM.L_SettingsAreStale = "設定無法儲存，因為不是最新的。";
Strings.OfficeOM.L_SettingsCannotSave = "無法儲存設定。";
Strings.OfficeOM.L_SettingsStaleError = "設定過時錯誤";
Strings.OfficeOM.L_ShowWindowDialogNotification = "「{0}」想要顯示新的視窗。";
Strings.OfficeOM.L_ShowWindowDialogNotificationAllow = "允許";
Strings.OfficeOM.L_ShowWindowDialogNotificationIgnore = "忽略";
Strings.OfficeOM.L_ShuttingDown = "作業失敗，因為伺服器上的資料不是最新的。";
Strings.OfficeOM.L_SliceSizeNotSupported = "不支援指定的圖塊大小。";
Strings.OfficeOM.L_SpecifiedIdNotExist = "指定的識別碼不存在。";
Strings.OfficeOM.L_Timeout = "作業逾時。";
Strings.OfficeOM.L_TooManyArguments = "引數過多";
Strings.OfficeOM.L_TooManyIncompleteRequests = "等候前一個呼叫完成。";
Strings.OfficeOM.L_TooManyOptionalFunction = "參數清單中的多個選擇性函數";
Strings.OfficeOM.L_TooManyOptionalObjects = "參數清單中的多個選擇性物件";
Strings.OfficeOM.L_UnknownBindingType = "不支援此繫結類型。";
Strings.OfficeOM.L_UnsupportedDataObject = "提供的資料物件類型不受支援。";
Strings.OfficeOM.L_UnsupportedEnumeration = "列舉不受支援";
Strings.OfficeOM.L_UnsupportedEnumerationMessage = "目前的主機應用程式中不支援列舉。";
Strings.OfficeOM.L_UnsupportedUserIdentity = "不支援使用者身分識別類型。";
Strings.OfficeOM.L_UnsupportedUserIdentityMessage = "不支援使用者的身分識別類型。";
Strings.OfficeOM.L_UserAborted = "使用者已中止同意要求。";
Strings.OfficeOM.L_UserAbortedMessage = "使用者未同意增益集的使用權限。";
Strings.OfficeOM.L_UserClickIgnore = "使用者選擇略過對話方塊。";
Strings.OfficeOM.L_UserNotSignedIn = "沒有使用者登入 Office。";
Strings.OfficeOM.L_ValueNotLoaded = '尚未載入結果物件的值。在讀取值屬性之前，請呼叫相關聯之要求內容中的 "context.sync()"。';