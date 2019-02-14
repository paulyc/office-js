
var Word;
(function (Word) {
	/// <summary> [Api set: WordApi] </summary>
	var Alignment = {
		__proto__: null,
		"mixed": "mixed",
		"unknown": "unknown",
		"left": "left",
		"centered": "centered",
		"right": "right",
		"justified": "justified",
	}
	Word.Alignment = Alignment;
})(Word || (Word = {__proto__: null}));

var Word;
(function (Word) {
	var Application = (function(_super) {
		__extends(Application, _super);
		function Application() {
			/// <summary> Represents the application object. [Api set: WordApi 1.3] </summary>
			/// <field name="context" type="Word.RequestContext">The request context associated with this object.</field>
			/// <field name="isNull" type="Boolean">Returns a boolean value for whether the corresponding object is null. You must call "context.sync()" before reading the isNull property.</field>
		}

		Application.prototype.load = function(option) {
			/// <summary>
			/// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
			/// </summary>
			/// <param name="option" type="string | string[] | OfficeExtension.LoadOption"/>
			/// <returns type="Word.Application"/>
		}
		Application.prototype.createDocument = function(base64File) {
			/// <summary>
			/// Creates a new document by using an optional base64 encoded .docx file. [Api set: WordApi 1.3]
			/// </summary>
			/// <param name="base64File" type="String" optional="true">Optional. The base64 encoded .docx file. The default value is null.</param>
			/// <returns type="Word.DocumentCreated"></returns>
		}

		return Application;
	})(OfficeExtension.ClientObject);
	Word.Application = Application;
})(Word || (Word = {__proto__: null}));

var Word;
(function (Word) {
	var Body = (function(_super) {
		__extends(Body, _super);
		function Body() {
			/// <summary> Represents the body of a document or a section. [Api set: WordApi 1.1] </summary>
			/// <field name="context" type="Word.RequestContext">The request context associated with this object.</field>
			/// <field name="isNull" type="Boolean">Returns a boolean value for whether the corresponding object is null. You must call "context.sync()" before reading the isNull property.</field>
			/// <field name="contentControls" type="Word.ContentControlCollection">Gets the collection of rich text content control objects in the body. Read-only. [Api set: WordApi 1.1]</field>
			/// <field name="font" type="Word.Font">Gets the text format of the body. Use this to get and set font name, size, color and other properties. Read-only. [Api set: WordApi 1.1]</field>
			/// <field name="inlinePictures" type="Word.InlinePictureCollection">Gets the collection of InlinePicture objects in the body. The collection does not include floating images. Read-only. [Api set: WordApi 1.1]</field>
			/// <field name="lists" type="Word.ListCollection">Gets the collection of list objects in the body. Read-only. [Api set: WordApi 1.3]</field>
			/// <field name="paragraphs" type="Word.ParagraphCollection">Gets the collection of paragraph objects in the body. Read-only. [Api set: WordApi 1.1]</field>
			/// <field name="parentBody" type="Word.Body">Gets the parent body of the body. For example, a table cell body&apos;s parent body could be a header. Throws if there isn&apos;t a parent body. Read-only. [Api set: WordApi 1.3]</field>
			/// <field name="parentBodyOrNullObject" type="Word.Body">Gets the parent body of the body. For example, a table cell body&apos;s parent body could be a header. Returns a null object if there isn&apos;t a parent body. Read-only. [Api set: WordApi 1.3]</field>
			/// <field name="parentContentControl" type="Word.ContentControl">Gets the content control that contains the body. Throws if there isn&apos;t a parent content control. Read-only. [Api set: WordApi 1.1]</field>
			/// <field name="parentContentControlOrNullObject" type="Word.ContentControl">Gets the content control that contains the body. Returns a null object if there isn&apos;t a parent content control. Read-only. [Api set: WordApi 1.3]</field>
			/// <field name="parentSection" type="Word.Section">Gets the parent section of the body. Throws if there isn&apos;t a parent section. Read-only. [Api set: WordApi 1.3]</field>
			/// <field name="parentSectionOrNullObject" type="Word.Section">Gets the parent section of the body. Returns a null object if there isn&apos;t a parent section. Read-only. [Api set: WordApi 1.3]</field>
			/// <field name="style" type="String">Gets or sets the style name for the body. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the &quot;styleBuiltIn&quot; property. [Api set: WordApi 1.1]</field>
			/// <field name="styleBuiltIn" type="String">Gets or sets the built-in style name for the body. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the &quot;style&quot; property. [Api set: WordApi 1.3]</field>
			/// <field name="tables" type="Word.TableCollection">Gets the collection of table objects in the body. Read-only. [Api set: WordApi 1.3]</field>
			/// <field name="text" type="String">Gets the text of the body. Use the insertText method to insert text. Read-only. [Api set: WordApi 1.1]</field>
			/// <field name="type" type="String">Gets the type of the body. The type can be &apos;MainDoc&apos;, &apos;Section&apos;, &apos;Header&apos;, &apos;Footer&apos;, or &apos;TableCell&apos;. Read-only. [Api set: WordApi 1.3]</field>
		}

		Body.prototype.load = function(option) {
			/// <summary>
			/// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
			/// </summary>
			/// <param name="option" type="string | string[] | OfficeExtension.LoadOption"/>
			/// <returns type="Word.Body"/>
		}

		Body.prototype.set = function() {
			/// <signature>
			/// <summary>Sets multiple properties on the object at the same time, based on JSON input.</summary>
			/// <param name="properties" type="Word.Interfaces.BodyUpdateData">Properties described by the Word.Interfaces.BodyUpdateData interface.</param>
			/// <param name="options" type="string">Options of the form { throwOnReadOnly?: boolean }
			/// <br />
			/// * throwOnReadOnly: Throw an error if the passed-in property list includes read-only properties (default = true).
			/// </param>
			/// </signature>
			/// <signature>
			/// <summary>Sets multiple properties on the object at the same time, based on an existing loaded object.</summary>
			/// <param name="properties" type="Body">An existing Body object, with properties that have already been loaded and synced.</param>
			/// </signature>
		}
		Body.prototype.clear = function() {
			/// <summary>
			/// Clears the contents of the body object. The user can perform the undo operation on the cleared content. [Api set: WordApi 1.1]
			/// </summary>
			/// <returns ></returns>
		}
		Body.prototype.getHtml = function() {
			/// <summary>
			/// Gets an HTML representation of the body object. When rendered in a web page or HTML viewer, the formatting will be a close, but not exact, match for of the formatting of the document. This method does not return the exact same HTML for the same document on different platforms (Windows, Mac, Word Online, etc.). If you need exact fidelity, or consistency across platforms, use `Body.getOoxml()` and convert the returned XML to HTML. [Api set: WordApi 1.1]
			/// </summary>
			/// <returns type="OfficeExtension.ClientResult&lt;string&gt;"></returns>
			var result = new OfficeExtension.ClientResult();
			result.__proto__ = null;
			result.value = '';
			return result;
		}
		Body.prototype.getOoxml = function() {
			/// <summary>
			/// Gets the OOXML (Office Open XML) representation of the body object. [Api set: WordApi 1.1]
			/// </summary>
			/// <returns type="OfficeExtension.ClientResult&lt;string&gt;"></returns>
			var result = new OfficeExtension.ClientResult();
			result.__proto__ = null;
			result.value = '';
			return result;
		}
		Body.prototype.getRange = function(rangeLocation) {
			/// <summary>
			/// Gets the whole body, or the starting or ending point of the body, as a range. [Api set: WordApi 1.3]
			/// </summary>
			/// <param name="rangeLocation" type="String" optional="true">Optional. The range location can be &apos;Whole&apos;, &apos;Start&apos;, &apos;End&apos;, &apos;After&apos;, or &apos;Content&apos;.</param>
			/// <returns type="Word.Range"></returns>
		}
		Body.prototype.insertBreak = function(breakType, insertLocation) {
			/// <summary>
			/// Inserts a break at the specified location in the main document. The insertLocation value can be &apos;Start&apos; or &apos;End&apos;. [Api set: WordApi 1.1]
			/// </summary>
			/// <param name="breakType" type="String">Required. The break type to add to the body.</param>
			/// <param name="insertLocation" type="String">Required. The value can be &apos;Start&apos; or &apos;End&apos;.</param>
			/// <returns ></returns>
		}
		Body.prototype.insertContentControl = function() {
			/// <summary>
			/// Wraps the body object with a Rich Text content control. [Api set: WordApi 1.1]
			/// </summary>
			/// <returns type="Word.ContentControl"></returns>
		}
		Body.prototype.insertFileFromBase64 = function(base64File, insertLocation) {
			/// <summary>
			/// Inserts a document into the body at the specified location. The insertLocation value can be &apos;Replace&apos;, &apos;Start&apos;, or &apos;End&apos;. [Api set: WordApi 1.1]
			/// </summary>
			/// <param name="base64File" type="String">Required. The base64 encoded content of a .docx file.</param>
			/// <param name="insertLocation" type="String">Required. The value can be &apos;Replace&apos;, &apos;Start&apos;, or &apos;End&apos;.</param>
			/// <returns type="Word.Range"></returns>
		}
		Body.prototype.insertHtml = function(html, insertLocation) {
			/// <summary>
			/// Inserts HTML at the specified location. The insertLocation value can be &apos;Replace&apos;, &apos;Start&apos;, or &apos;End&apos;. [Api set: WordApi 1.1]
			/// </summary>
			/// <param name="html" type="String">Required. The HTML to be inserted in the document.</param>
			/// <param name="insertLocation" type="String">Required. The value can be &apos;Replace&apos;, &apos;Start&apos;, or &apos;End&apos;.</param>
			/// <returns type="Word.Range"></returns>
		}
		Body.prototype.insertInlinePictureFromBase64 = function(base64EncodedImage, insertLocation) {
			/// <summary>
			/// Inserts a picture into the body at the specified location. The insertLocation value can be &apos;Start&apos; or &apos;End&apos;. [Api set: WordApi 1.2]
			/// </summary>
			/// <param name="base64EncodedImage" type="String">Required. The base64 encoded image to be inserted in the body.</param>
			/// <param name="insertLocation" type="String">Required. The value can be &apos;Start&apos; or &apos;End&apos;.</param>
			/// <returns type="Word.InlinePicture"></returns>
		}
		Body.prototype.insertOoxml = function(ooxml, insertLocation) {
			/// <summary>
			/// Inserts OOXML at the specified location.  The insertLocation value can be &apos;Replace&apos;, &apos;Start&apos;, or &apos;End&apos;. [Api set: WordApi 1.1]
			/// </summary>
			/// <param name="ooxml" type="String">Required. The OOXML to be inserted.</param>
			/// <param name="insertLocation" type="String">Required. The value can be &apos;Replace&apos;, &apos;Start&apos;, or &apos;End&apos;.</param>
			/// <returns type="Word.Range"></returns>
		}
		Body.prototype.insertParagraph = function(paragraphText, insertLocation) {
			/// <summary>
			/// Inserts a paragraph at the specified location. The insertLocation value can be &apos;Start&apos; or &apos;End&apos;. [Api set: WordApi 1.1]
			/// </summary>
			/// <param name="paragraphText" type="String">Required. The paragraph text to be inserted.</param>
			/// <param name="insertLocation" type="String">Required. The value can be &apos;Start&apos; or &apos;End&apos;.</param>
			/// <returns type="Word.Paragraph"></returns>
		}
		Body.prototype.insertTable = function(rowCount, columnCount, insertLocation, values) {
			/// <summary>
			/// Inserts a table with the specified number of rows and columns. The insertLocation value can be &apos;Start&apos; or &apos;End&apos;. [Api set: WordApi 1.3]
			/// </summary>
			/// <param name="rowCount" type="Number">Required. The number of rows in the table.</param>
			/// <param name="columnCount" type="Number">Required. The number of columns in the table.</param>
			/// <param name="insertLocation" type="String">Required. The value can be &apos;Start&apos; or &apos;End&apos;.</param>
			/// <param name="values" type="Array" elementType="Array" optional="true">Optional 2D array. Cells are filled if the corresponding strings are specified in the array.</param>
			/// <returns type="Word.Table"></returns>
		}
		Body.prototype.insertText = function(text, insertLocation) {
			/// <summary>
			/// Inserts text into the body at the specified location. The insertLocation value can be &apos;Replace&apos;, &apos;Start&apos;, or &apos;End&apos;. [Api set: WordApi 1.1]
			/// </summary>
			/// <param name="text" type="String">Required. Text to be inserted.</param>
			/// <param name="insertLocation" type="String">Required. The value can be &apos;Replace&apos;, &apos;Start&apos;, or &apos;End&apos;.</param>
			/// <returns type="Word.Range"></returns>
		}
		Body.prototype.search = function(searchText, searchOptions) {
			/// <summary>
			/// Performs a search with the specified SearchOptions on the scope of the body object. The search results are a collection of range objects. [Api set: WordApi 1.1]
			/// </summary>
			/// <param name="searchText" type="String">Required. The search text. Can be a maximum of 255 characters.</param>
			/// <param name="searchOptions" type="Word.SearchOptions" optional="true">Optional. Options for the search.</param>
			/// <returns type="Word.RangeCollection"></returns>
		}
		Body.prototype.select = function(selectionMode) {
			/// <summary>
			/// Selects the body and navigates the Word UI to it. [Api set: WordApi 1.1]
			/// </summary>
			/// <param name="selectionMode" type="String" optional="true">Optional. The selection mode can be &apos;Select&apos;, &apos;Start&apos;, or &apos;End&apos;. &apos;Select&apos; is the default.</param>
			/// <returns ></returns>
		}

		Body.prototype.track = function() {
			/// <summary>
			/// Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
			/// </summary>
			/// <returns type="Word.Body"/>
		}

		Body.prototype.untrack = function() {
			/// <summary>
			/// Release the memory associated with this object, if has previous been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
			/// </summary>
			/// <returns type="Word.Body"/>
		}

		return Body;
	})(OfficeExtension.ClientObject);
	Word.Body = Body;
})(Word || (Word = {__proto__: null}));

var Word;
(function (Word) {
	/// <summary> [Api set: WordApi] </summary>
	var BodyType = {
		__proto__: null,
		"unknown": "unknown",
		"mainDoc": "mainDoc",
		"section": "section",
		"header": "header",
		"footer": "footer",
		"tableCell": "tableCell",
	}
	Word.BodyType = BodyType;
})(Word || (Word = {__proto__: null}));

var Word;
(function (Word) {
	/// <summary> [Api set: WordApi] </summary>
	var BorderLocation = {
		__proto__: null,
		"top": "top",
		"left": "left",
		"bottom": "bottom",
		"right": "right",
		"insideHorizontal": "insideHorizontal",
		"insideVertical": "insideVertical",
		"inside": "inside",
		"outside": "outside",
		"all": "all",
	}
	Word.BorderLocation = BorderLocation;
})(Word || (Word = {__proto__: null}));

var Word;
(function (Word) {
	/// <summary> [Api set: WordApi] </summary>
	var BorderType = {
		__proto__: null,
		"mixed": "mixed",
		"none": "none",
		"single": "single",
		"double": "double",
		"dotted": "dotted",
		"dashed": "dashed",
		"dotDashed": "dotDashed",
		"dot2Dashed": "dot2Dashed",
		"triple": "triple",
		"thinThickSmall": "thinThickSmall",
		"thickThinSmall": "thickThinSmall",
		"thinThickThinSmall": "thinThickThinSmall",
		"thinThickMed": "thinThickMed",
		"thickThinMed": "thickThinMed",
		"thinThickThinMed": "thinThickThinMed",
		"thinThickLarge": "thinThickLarge",
		"thickThinLarge": "thickThinLarge",
		"thinThickThinLarge": "thinThickThinLarge",
		"wave": "wave",
		"doubleWave": "doubleWave",
		"dashedSmall": "dashedSmall",
		"dashDotStroked": "dashDotStroked",
		"threeDEmboss": "threeDEmboss",
		"threeDEngrave": "threeDEngrave",
	}
	Word.BorderType = BorderType;
})(Word || (Word = {__proto__: null}));

var Word;
(function (Word) {
	/// <summary> Specifies the form of a break. [Api set: WordApi] </summary>
	var BreakType = {
		__proto__: null,
		"page": "page",
		"sectionNext": "sectionNext",
		"sectionContinuous": "sectionContinuous",
		"sectionEven": "sectionEven",
		"sectionOdd": "sectionOdd",
		"line": "line",
	}
	Word.BreakType = BreakType;
})(Word || (Word = {__proto__: null}));

var Word;
(function (Word) {
	/// <summary> [Api set: WordApi] </summary>
	var CellPaddingLocation = {
		__proto__: null,
		"top": "top",
		"left": "left",
		"bottom": "bottom",
		"right": "right",
	}
	Word.CellPaddingLocation = CellPaddingLocation;
})(Word || (Word = {__proto__: null}));

var Word;
(function (Word) {
	var ContentControl = (function(_super) {
		__extends(ContentControl, _super);
		function ContentControl() {
			/// <summary> Represents a content control. Content controls are bounded and potentially labeled regions in a document that serve as containers for specific types of content. Individual content controls may contain contents such as images, tables, or paragraphs of formatted text. Currently, only rich text content controls are supported. [Api set: WordApi 1.1] </summary>
			/// <field name="context" type="Word.RequestContext">The request context associated with this object.</field>
			/// <field name="isNull" type="Boolean">Returns a boolean value for whether the corresponding object is null. You must call "context.sync()" before reading the isNull property.</field>
			/// <field name="appearance" type="String">Gets or sets the appearance of the content control. The value can be &apos;BoundingBox&apos;, &apos;Tags&apos;, or &apos;Hidden&apos;. [Api set: WordApi 1.1]</field>
			/// <field name="cannotDelete" type="Boolean">Gets or sets a value that indicates whether the user can delete the content control. Mutually exclusive with removeWhenEdited. [Api set: WordApi 1.1]</field>
			/// <field name="cannotEdit" type="Boolean">Gets or sets a value that indicates whether the user can edit the contents of the content control. [Api set: WordApi 1.1]</field>
			/// <field name="color" type="String">Gets or sets the color of the content control. Color is specified in &apos;#RRGGBB&apos; format or by using the color name. [Api set: WordApi 1.1]</field>
			/// <field name="contentControls" type="Word.ContentControlCollection">Gets the collection of content control objects in the content control. Read-only. [Api set: WordApi 1.1]</field>
			/// <field name="font" type="Word.Font">Gets the text format of the content control. Use this to get and set font name, size, color, and other properties. Read-only. [Api set: WordApi 1.1]</field>
			/// <field name="id" type="Number">Gets an integer that represents the content control identifier. Read-only. [Api set: WordApi 1.1]</field>
			/// <field name="inlinePictures" type="Word.InlinePictureCollection">Gets the collection of inlinePicture objects in the content control. The collection does not include floating images. Read-only. [Api set: WordApi 1.1]</field>
			/// <field name="lists" type="Word.ListCollection">Gets the collection of list objects in the content control. Read-only. [Api set: WordApi 1.3]</field>
			/// <field name="paragraphs" type="Word.ParagraphCollection">Get the collection of paragraph objects in the content control. Read-only. [Api set: WordApi 1.1]</field>
			/// <field name="parentBody" type="Word.Body">Gets the parent body of the content control. Read-only. [Api set: WordApi 1.3]</field>
			/// <field name="parentContentControl" type="Word.ContentControl">Gets the content control that contains the content control. Throws if there isn&apos;t a parent content control. Read-only. [Api set: WordApi 1.1]</field>
			/// <field name="parentContentControlOrNullObject" type="Word.ContentControl">Gets the content control that contains the content control. Returns a null object if there isn&apos;t a parent content control. Read-only. [Api set: WordApi 1.3]</field>
			/// <field name="parentTable" type="Word.Table">Gets the table that contains the content control. Throws if it is not contained in a table. Read-only. [Api set: WordApi 1.3]</field>
			/// <field name="parentTableCell" type="Word.TableCell">Gets the table cell that contains the content control. Throws if it is not contained in a table cell. Read-only. [Api set: WordApi 1.3]</field>
			/// <field name="parentTableCellOrNullObject" type="Word.TableCell">Gets the table cell that contains the content control. Returns a null object if it is not contained in a table cell. Read-only. [Api set: WordApi 1.3]</field>
			/// <field name="parentTableOrNullObject" type="Word.Table">Gets the table that contains the content control. Returns a null object if it is not contained in a table. Read-only. [Api set: WordApi 1.3]</field>
			/// <field name="placeholderText" type="String">Gets or sets the placeholder text of the content control. Dimmed text will be displayed when the content control is empty. [Api set: WordApi 1.1]</field>
			/// <field name="removeWhenEdited" type="Boolean">Gets or sets a value that indicates whether the content control is removed after it is edited. Mutually exclusive with cannotDelete. [Api set: WordApi 1.1]</field>
			/// <field name="style" type="String">Gets or sets the style name for the content control. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the &quot;styleBuiltIn&quot; property. [Api set: WordApi 1.1]</field>
			/// <field name="styleBuiltIn" type="String">Gets or sets the built-in style name for the content control. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the &quot;style&quot; property. [Api set: WordApi 1.3]</field>
			/// <field name="subtype" type="String">Gets the content control subtype. The subtype can be &apos;RichTextInline&apos;, &apos;RichTextParagraphs&apos;, &apos;RichTextTableCell&apos;, &apos;RichTextTableRow&apos; and &apos;RichTextTable&apos; for rich text content controls. Read-only. [Api set: WordApi 1.3]</field>
			/// <field name="tables" type="Word.TableCollection">Gets the collection of table objects in the content control. Read-only. [Api set: WordApi 1.3]</field>
			/// <field name="tag" type="String">Gets or sets a tag to identify a content control. [Api set: WordApi 1.1]</field>
			/// <field name="text" type="String">Gets the text of the content control. Read-only. [Api set: WordApi 1.1]</field>
			/// <field name="title" type="String">Gets or sets the title for a content control. [Api set: WordApi 1.1]</field>
			/// <field name="type" type="String">Gets the content control type. Only rich text content controls are supported currently. Read-only. [Api set: WordApi 1.1]</field>
			/// <field name="onDataChanged" type="OfficeExtension.EventHandlers">Occurs when data within the content control are changed. To get the new text, load this content control in the handler. To get the old text, do not load it. [Api set: WordApi 1.4]</field>
			/// <field name="onDeleted" type="OfficeExtension.EventHandlers">Occurs when the content control is deleted. Do not load this content control in the handler, otherwise you won&apos;t be able to get its original properties. [Api set: WordApi 1.4]</field>
			/// <field name="onSelectionChanged" type="OfficeExtension.EventHandlers">Occurs when selection within the content control is changed. [Api set: WordApi 1.4]</field>
		}

		ContentControl.prototype.load = function(option) {
			/// <summary>
			/// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
			/// </summary>
			/// <param name="option" type="string | string[] | OfficeExtension.LoadOption"/>
			/// <returns type="Word.ContentControl"/>
		}

		ContentControl.prototype.set = function() {
			/// <signature>
			/// <summary>Sets multiple properties on the object at the same time, based on JSON input.</summary>
			/// <param name="properties" type="Word.Interfaces.ContentControlUpdateData">Properties described by the Word.Interfaces.ContentControlUpdateData interface.</param>
			/// <param name="options" type="string">Options of the form { throwOnReadOnly?: boolean }
			/// <br />
			/// * throwOnReadOnly: Throw an error if the passed-in property list includes read-only properties (default = true).
			/// </param>
			/// </signature>
			/// <signature>
			/// <summary>Sets multiple properties on the object at the same time, based on an existing loaded object.</summary>
			/// <param name="properties" type="ContentControl">An existing ContentControl object, with properties that have already been loaded and synced.</param>
			/// </signature>
		}
		ContentControl.prototype.clear = function() {
			/// <summary>
			/// Clears the contents of the content control. The user can perform the undo operation on the cleared content. [Api set: WordApi 1.1]
			/// </summary>
			/// <returns ></returns>
		}
		ContentControl.prototype.delete = function(keepContent) {
			/// <summary>
			/// Deletes the content control and its content. If keepContent is set to true, the content is not deleted. [Api set: WordApi 1.1]
			/// </summary>
			/// <param name="keepContent" type="Boolean">Required. Indicates whether the content should be deleted with the content control. If keepContent is set to true, the content is not deleted.</param>
			/// <returns ></returns>
		}
		ContentControl.prototype.getHtml = function() {
			/// <summary>
			/// Gets an HTML representation of the content control object. When rendered in a web page or HTML viewer, the formatting will be a close, but not exact, match for of the formatting of the document. This method does not return the exact same HTML for the same document on different platforms (Windows, Mac, Word Online, etc.). If you need exact fidelity, or consistency across platforms, use `ContentControl.getOoxml()` and convert the returned XML to HTML. [Api set: WordApi 1.1]
			/// </summary>
			/// <returns type="OfficeExtension.ClientResult&lt;string&gt;"></returns>
			var result = new OfficeExtension.ClientResult();
			result.__proto__ = null;
			result.value = '';
			return result;
		}
		ContentControl.prototype.getOoxml = function() {
			/// <summary>
			/// Gets the Office Open XML (OOXML) representation of the content control object. [Api set: WordApi 1.1]
			/// </summary>
			/// <returns type="OfficeExtension.ClientResult&lt;string&gt;"></returns>
			var result = new OfficeExtension.ClientResult();
			result.__proto__ = null;
			result.value = '';
			return result;
		}
		ContentControl.prototype.getRange = function(rangeLocation) {
			/// <summary>
			/// Gets the whole content control, or the starting or ending point of the content control, as a range. [Api set: WordApi 1.3]
			/// </summary>
			/// <param name="rangeLocation" type="String" optional="true">Optional. The range location can be &apos;Whole&apos;, &apos;Before&apos;, &apos;Start&apos;, &apos;End&apos;, &apos;After&apos;, or &apos;Content&apos;.</param>
			/// <returns type="Word.Range"></returns>
		}
		ContentControl.prototype.getTextRanges = function(endingMarks, trimSpacing) {
			/// <summary>
			/// Gets the text ranges in the content control by using punctuation marks and/or other ending marks. [Api set: WordApi 1.3]
			/// </summary>
			/// <param name="endingMarks" type="Array" elementType="String">Required. The punctuation marks and/or other ending marks as an array of strings.</param>
			/// <param name="trimSpacing" type="Boolean" optional="true">Optional. Indicates whether to trim spacing characters (spaces, tabs, column breaks, and paragraph end marks) from the start and end of the ranges returned in the range collection. Default is false which indicates that spacing characters at the start and end of the ranges are included in the range collection.</param>
			/// <returns type="Word.RangeCollection"></returns>
		}
		ContentControl.prototype.insertBreak = function(breakType, insertLocation) {
			/// <summary>
			/// Inserts a break at the specified location in the main document. The insertLocation value can be &apos;Start&apos;, &apos;End&apos;, &apos;Before&apos;, or &apos;After&apos;. This method cannot be used with &apos;RichTextTable&apos;, &apos;RichTextTableRow&apos; and &apos;RichTextTableCell&apos; content controls. [Api set: WordApi 1.1]
			/// </summary>
			/// <param name="breakType" type="String">Required. Type of break.</param>
			/// <param name="insertLocation" type="String">Required. The value can be &apos;Start&apos;, &apos;End&apos;, &apos;Before&apos;, or &apos;After&apos;.</param>
			/// <returns ></returns>
		}
		ContentControl.prototype.insertFileFromBase64 = function(base64File, insertLocation) {
			/// <summary>
			/// Inserts a document into the content control at the specified location. The insertLocation value can be &apos;Replace&apos;, &apos;Start&apos;, or &apos;End&apos;. [Api set: WordApi 1.1]
			/// </summary>
			/// <param name="base64File" type="String">Required. The base64 encoded content of a .docx file.</param>
			/// <param name="insertLocation" type="String">Required. The value can be &apos;Replace&apos;, &apos;Start&apos;, or &apos;End&apos;. &apos;Replace&apos; cannot be used with &apos;RichTextTable&apos; and &apos;RichTextTableRow&apos; content controls.</param>
			/// <returns type="Word.Range"></returns>
		}
		ContentControl.prototype.insertHtml = function(html, insertLocation) {
			/// <summary>
			/// Inserts HTML into the content control at the specified location. The insertLocation value can be &apos;Replace&apos;, &apos;Start&apos;, or &apos;End&apos;. [Api set: WordApi 1.1]
			/// </summary>
			/// <param name="html" type="String">Required. The HTML to be inserted in to the content control.</param>
			/// <param name="insertLocation" type="String">Required. The value can be &apos;Replace&apos;, &apos;Start&apos;, or &apos;End&apos;. &apos;Replace&apos; cannot be used with &apos;RichTextTable&apos; and &apos;RichTextTableRow&apos; content controls.</param>
			/// <returns type="Word.Range"></returns>
		}
		ContentControl.prototype.insertInlinePictureFromBase64 = function(base64EncodedImage, insertLocation) {
			/// <summary>
			/// Inserts an inline picture into the content control at the specified location. The insertLocation value can be &apos;Replace&apos;, &apos;Start&apos;, or &apos;End&apos;. [Api set: WordApi 1.2]
			/// </summary>
			/// <param name="base64EncodedImage" type="String">Required. The base64 encoded image to be inserted in the content control.</param>
			/// <param name="insertLocation" type="String">Required. The value can be &apos;Replace&apos;, &apos;Start&apos;, or &apos;End&apos;. &apos;Replace&apos; cannot be used with &apos;RichTextTable&apos; and &apos;RichTextTableRow&apos; content controls.</param>
			/// <returns type="Word.InlinePicture"></returns>
		}
		ContentControl.prototype.insertOoxml = function(ooxml, insertLocation) {
			/// <summary>
			/// Inserts OOXML into the content control at the specified location.  The insertLocation value can be &apos;Replace&apos;, &apos;Start&apos;, or &apos;End&apos;. [Api set: WordApi 1.1]
			/// </summary>
			/// <param name="ooxml" type="String">Required. The OOXML to be inserted in to the content control.</param>
			/// <param name="insertLocation" type="String">Required. The value can be &apos;Replace&apos;, &apos;Start&apos;, or &apos;End&apos;. &apos;Replace&apos; cannot be used with &apos;RichTextTable&apos; and &apos;RichTextTableRow&apos; content controls.</param>
			/// <returns type="Word.Range"></returns>
		}
		ContentControl.prototype.insertParagraph = function(paragraphText, insertLocation) {
			/// <summary>
			/// Inserts a paragraph at the specified location. The insertLocation value can be &apos;Start&apos;, &apos;End&apos;, &apos;Before&apos;, or &apos;After&apos;. [Api set: WordApi 1.1]
			/// </summary>
			/// <param name="paragraphText" type="String">Required. The paragraph text to be inserted.</param>
			/// <param name="insertLocation" type="String">Required. The value can be &apos;Start&apos;, &apos;End&apos;, &apos;Before&apos;, or &apos;After&apos;. &apos;Before&apos; and &apos;After&apos; cannot be used with &apos;RichTextTable&apos;, &apos;RichTextTableRow&apos; and &apos;RichTextTableCell&apos; content controls.</param>
			/// <returns type="Word.Paragraph"></returns>
		}
		ContentControl.prototype.insertTable = function(rowCount, columnCount, insertLocation, values) {
			/// <summary>
			/// Inserts a table with the specified number of rows and columns into, or next to, a content control. The insertLocation value can be &apos;Start&apos;, &apos;End&apos;, &apos;Before&apos;, or &apos;After&apos;. [Api set: WordApi 1.3]
			/// </summary>
			/// <param name="rowCount" type="Number">Required. The number of rows in the table.</param>
			/// <param name="columnCount" type="Number">Required. The number of columns in the table.</param>
			/// <param name="insertLocation" type="String">Required. The value can be &apos;Start&apos;, &apos;End&apos;, &apos;Before&apos;, or &apos;After&apos;. &apos;Before&apos; and &apos;After&apos; cannot be used with &apos;RichTextTable&apos;, &apos;RichTextTableRow&apos; and &apos;RichTextTableCell&apos; content controls.</param>
			/// <param name="values" type="Array" elementType="Array" optional="true">Optional 2D array. Cells are filled if the corresponding strings are specified in the array.</param>
			/// <returns type="Word.Table"></returns>
		}
		ContentControl.prototype.insertText = function(text, insertLocation) {
			/// <summary>
			/// Inserts text into the content control at the specified location. The insertLocation value can be &apos;Replace&apos;, &apos;Start&apos;, or &apos;End&apos;. [Api set: WordApi 1.1]
			/// </summary>
			/// <param name="text" type="String">Required. The text to be inserted in to the content control.</param>
			/// <param name="insertLocation" type="String">Required. The value can be &apos;Replace&apos;, &apos;Start&apos;, or &apos;End&apos;. &apos;Replace&apos; cannot be used with &apos;RichTextTable&apos; and &apos;RichTextTableRow&apos; content controls.</param>
			/// <returns type="Word.Range"></returns>
		}
		ContentControl.prototype.search = function(searchText, searchOptions) {
			/// <summary>
			/// Performs a search with the specified SearchOptions on the scope of the content control object. The search results are a collection of range objects. [Api set: WordApi 1.1]
			/// </summary>
			/// <param name="searchText" type="String">Required. The search text.</param>
			/// <param name="searchOptions" type="Word.SearchOptions" optional="true">Optional. Options for the search.</param>
			/// <returns type="Word.RangeCollection"></returns>
		}
		ContentControl.prototype.select = function(selectionMode) {
			/// <summary>
			/// Selects the content control. This causes Word to scroll to the selection. [Api set: WordApi 1.1]
			/// </summary>
			/// <param name="selectionMode" type="String" optional="true">Optional. The selection mode can be &apos;Select&apos;, &apos;Start&apos;, or &apos;End&apos;. &apos;Select&apos; is the default.</param>
			/// <returns ></returns>
		}
		ContentControl.prototype.split = function(delimiters, multiParagraphs, trimDelimiters, trimSpacing) {
			/// <summary>
			/// Splits the content control into child ranges by using delimiters. [Api set: WordApi 1.3]
			/// </summary>
			/// <param name="delimiters" type="Array" elementType="String">Required. The delimiters as an array of strings.</param>
			/// <param name="multiParagraphs" type="Boolean" optional="true">Optional. Indicates whether a returned child range can cover multiple paragraphs. Default is false which indicates that the paragraph boundaries are also used as delimiters.</param>
			/// <param name="trimDelimiters" type="Boolean" optional="true">Optional. Indicates whether to trim delimiters from the ranges in the range collection. Default is false which indicates that the delimiters are included in the ranges returned in the range collection.</param>
			/// <param name="trimSpacing" type="Boolean" optional="true">Optional. Indicates whether to trim spacing characters (spaces, tabs, column breaks, and paragraph end marks) from the start and end of the ranges returned in the range collection. Default is false which indicates that spacing characters at the start and end of the ranges are included in the range collection.</param>
			/// <returns type="Word.RangeCollection"></returns>
		}
		ContentControl.prototype.onDataChanged = {
			__proto__: null,
			add: function (handler) {
				/// <param name="handler" type="function(eventArgs: Word.Interfaces.ContentControlEventArgs)">Handler for the event. EventArgs: Provides information about the content control that raised an event. </param>
				/// <returns type="OfficeExtension.EventHandlerResult"></returns>
				var eventInfo = new Word.Interfaces.ContentControlEventArgs();
				eventInfo.__proto__ = null;
				handler(eventInfo);
			},
			remove: function (handler) {
				/// <param name="handler" type="function(eventArgs: Word.Interfaces.ContentControlEventArgs)">Handler for the event.</param>
				return;
			}
		};
		ContentControl.prototype.onDeleted = {
			__proto__: null,
			add: function (handler) {
				/// <param name="handler" type="function(eventArgs: Word.Interfaces.ContentControlEventArgs)">Handler for the event. EventArgs: Provides information about the content control that raised an event. </param>
				/// <returns type="OfficeExtension.EventHandlerResult"></returns>
				var eventInfo = new Word.Interfaces.ContentControlEventArgs();
				eventInfo.__proto__ = null;
				handler(eventInfo);
			},
			remove: function (handler) {
				/// <param name="handler" type="function(eventArgs: Word.Interfaces.ContentControlEventArgs)">Handler for the event.</param>
				return;
			}
		};
		ContentControl.prototype.onSelectionChanged = {
			__proto__: null,
			add: function (handler) {
				/// <param name="handler" type="function(eventArgs: Word.Interfaces.ContentControlEventArgs)">Handler for the event. EventArgs: Provides information about the content control that raised an event. </param>
				/// <returns type="OfficeExtension.EventHandlerResult"></returns>
				var eventInfo = new Word.Interfaces.ContentControlEventArgs();
				eventInfo.__proto__ = null;
				handler(eventInfo);
			},
			remove: function (handler) {
				/// <param name="handler" type="function(eventArgs: Word.Interfaces.ContentControlEventArgs)">Handler for the event.</param>
				return;
			}
		};

		ContentControl.prototype.track = function() {
			/// <summary>
			/// Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
			/// </summary>
			/// <returns type="Word.ContentControl"/>
		}

		ContentControl.prototype.untrack = function() {
			/// <summary>
			/// Release the memory associated with this object, if has previous been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
			/// </summary>
			/// <returns type="Word.ContentControl"/>
		}

		return ContentControl;
	})(OfficeExtension.ClientObject);
	Word.ContentControl = ContentControl;
})(Word || (Word = {__proto__: null}));

var Word;
(function (Word) {
	/// <summary> ContentControl appearance [Api set: WordApi] </summary>
	var ContentControlAppearance = {
		__proto__: null,
		"boundingBox": "boundingBox",
		"tags": "tags",
		"hidden": "hidden",
	}
	Word.ContentControlAppearance = ContentControlAppearance;
})(Word || (Word = {__proto__: null}));

var Word;
(function (Word) {
	var ContentControlCollection = (function(_super) {
		__extends(ContentControlCollection, _super);
		function ContentControlCollection() {
			/// <summary> Contains a collection of {@link Word.ContentControl} objects. Content controls are bounded and potentially labeled regions in a document that serve as containers for specific types of content. Individual content controls may contain contents such as images, tables, or paragraphs of formatted text. Currently, only rich text content controls are supported. [Api set: WordApi 1.1] </summary>
			/// <field name="context" type="Word.RequestContext">The request context associated with this object.</field>
			/// <field name="isNull" type="Boolean">Returns a boolean value for whether the corresponding object is null. You must call "context.sync()" before reading the isNull property.</field>
			/// <field name="items" type="Array" elementType="Word.ContentControl">Gets the loaded child items in this collection.</field>
		}

		ContentControlCollection.prototype.load = function(option) {
			/// <summary>
			/// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
			/// </summary>
			/// <param name="option" type="string | string[] | OfficeExtension.LoadOption"/>
			/// <returns type="Word.ContentControlCollection"/>
		}
		ContentControlCollection.prototype.getById = function(id) {
			/// <summary>
			/// Gets a content control by its identifier. Throws if there isn&apos;t a content control with the identifier in this collection. [Api set: WordApi 1.1]
			/// </summary>
			/// <param name="id" type="Number">Required. A content control identifier.</param>
			/// <returns type="Word.ContentControl"></returns>
		}
		ContentControlCollection.prototype.getByIdOrNullObject = function(id) {
			/// <summary>
			/// Gets a content control by its identifier. Returns a null object if there isn&apos;t a content control with the identifier in this collection. [Api set: WordApi 1.3]
			/// </summary>
			/// <param name="id" type="Number">Required. A content control identifier.</param>
			/// <returns type="Word.ContentControl"></returns>
		}
		ContentControlCollection.prototype.getByTag = function(tag) {
			/// <summary>
			/// Gets the content controls that have the specified tag. [Api set: WordApi 1.1]
			/// </summary>
			/// <param name="tag" type="String">Required. A tag set on a content control.</param>
			/// <returns type="Word.ContentControlCollection"></returns>
		}
		ContentControlCollection.prototype.getByTitle = function(title) {
			/// <summary>
			/// Gets the content controls that have the specified title. [Api set: WordApi 1.1]
			/// </summary>
			/// <param name="title" type="String">Required. The title of a content control.</param>
			/// <returns type="Word.ContentControlCollection"></returns>
		}
		ContentControlCollection.prototype.getByTypes = function(types) {
			/// <summary>
			/// Gets the content controls that have the specified types and/or subtypes. [Api set: WordApi 1.3]
			/// </summary>
			/// <param name="types" type="Array" elementType="String">Required. An array of content control types and/or subtypes.</param>
			/// <returns type="Word.ContentControlCollection"></returns>
		}
		ContentControlCollection.prototype.getFirst = function() {
			/// <summary>
			/// Gets the first content control in this collection. Throws if this collection is empty. [Api set: WordApi 1.3]
			/// </summary>
			/// <returns type="Word.ContentControl"></returns>
		}
		ContentControlCollection.prototype.getFirstOrNullObject = function() {
			/// <summary>
			/// Gets the first content control in this collection. Returns a null object if this collection is empty. [Api set: WordApi 1.3]
			/// </summary>
			/// <returns type="Word.ContentControl"></returns>
		}
		ContentControlCollection.prototype.getItem = function(index) {
			/// <summary>
			/// Gets a content control by its index in the collection. [Api set: WordApi 1.1]
			/// </summary>
			/// <param name="index" >The index.</param>
			/// <returns type="Word.ContentControl"></returns>
		}

		ContentControlCollection.prototype.track = function() {
			/// <summary>
			/// Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
			/// </summary>
			/// <returns type="Word.ContentControlCollection"/>
		}

		ContentControlCollection.prototype.untrack = function() {
			/// <summary>
			/// Release the memory associated with this object, if has previous been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
			/// </summary>
			/// <returns type="Word.ContentControlCollection"/>
		}

		return ContentControlCollection;
	})(OfficeExtension.ClientObject);
	Word.ContentControlCollection = ContentControlCollection;
})(Word || (Word = {__proto__: null}));

var Word;
(function (Word) {
	var Interfaces;
	(function (Interfaces) {
		var ContentControlEventArgs = (function() {
			function ContentControlEventArgs() {
				/// <summary> Provides information about the content control that raised an event. [Api set: WordApi 1.4] </summary>
				/// <field name="contentControl" type="Word.ContentControl">The object that raised the event. Load this object to get its properties. [Api set: WordApi 1.4]</field>
				/// <field name="eventType" type="String">The event type. See Word.EventType for details. [Api set: WordApi 1.4]</field>
			}
			return ContentControlEventArgs;
		})();
		Interfaces.ContentControlEventArgs.__proto__ = null;
		Interfaces.ContentControlEventArgs = ContentControlEventArgs;
	})(Interfaces = Word.Interfaces || (Word.Interfaces = { __proto__: null}));
})(Word || (Word = {__proto__: null}));

var Word;
(function (Word) {
	/// <summary> Specifies supported content control types and subtypes. [Api set: WordApi] </summary>
	var ContentControlType = {
		__proto__: null,
		"unknown": "unknown",
		"richTextInline": "richTextInline",
		"richTextParagraphs": "richTextParagraphs",
		"richTextTableCell": "richTextTableCell",
		"richTextTableRow": "richTextTableRow",
		"richTextTable": "richTextTable",
		"plainTextInline": "plainTextInline",
		"plainTextParagraph": "plainTextParagraph",
		"picture": "picture",
		"buildingBlockGallery": "buildingBlockGallery",
		"checkBox": "checkBox",
		"comboBox": "comboBox",
		"dropDownList": "dropDownList",
		"datePicker": "datePicker",
		"repeatingSection": "repeatingSection",
		"richText": "richText",
		"plainText": "plainText",
	}
	Word.ContentControlType = ContentControlType;
})(Word || (Word = {__proto__: null}));

var Word;
(function (Word) {
	var CustomProperty = (function(_super) {
		__extends(CustomProperty, _super);
		function CustomProperty() {
			/// <summary> Represents a custom property. [Api set: WordApi 1.3] </summary>
			/// <field name="context" type="Word.RequestContext">The request context associated with this object.</field>
			/// <field name="isNull" type="Boolean">Returns a boolean value for whether the corresponding object is null. You must call "context.sync()" before reading the isNull property.</field>
			/// <field name="key" type="String">Gets the key of the custom property. Read only. [Api set: WordApi 1.3]</field>
			/// <field name="type" type="String">Gets the value type of the custom property. Possible values are: String, Number, Date, Boolean. Read only. [Api set: WordApi 1.3]</field>
			/// <field name="value" >Gets or sets the value of the custom property. Note that even though Word Online and the docx file format allow these properties to be arbitrarily long, the desktop version of Word will truncate string values to 255 16-bit chars (possibly creating invalid unicode by breaking up a surrogate pair). [Api set: WordApi 1.3]</field>
		}

		CustomProperty.prototype.load = function(option) {
			/// <summary>
			/// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
			/// </summary>
			/// <param name="option" type="string | string[] | OfficeExtension.LoadOption"/>
			/// <returns type="Word.CustomProperty"/>
		}

		CustomProperty.prototype.set = function() {
			/// <signature>
			/// <summary>Sets multiple properties on the object at the same time, based on JSON input.</summary>
			/// <param name="properties" type="Word.Interfaces.CustomPropertyUpdateData">Properties described by the Word.Interfaces.CustomPropertyUpdateData interface.</param>
			/// <param name="options" type="string">Options of the form { throwOnReadOnly?: boolean }
			/// <br />
			/// * throwOnReadOnly: Throw an error if the passed-in property list includes read-only properties (default = true).
			/// </param>
			/// </signature>
			/// <signature>
			/// <summary>Sets multiple properties on the object at the same time, based on an existing loaded object.</summary>
			/// <param name="properties" type="CustomProperty">An existing CustomProperty object, with properties that have already been loaded and synced.</param>
			/// </signature>
		}
		CustomProperty.prototype.delete = function() {
			/// <summary>
			/// Deletes the custom property. [Api set: WordApi 1.3]
			/// </summary>
			/// <returns ></returns>
		}

		CustomProperty.prototype.track = function() {
			/// <summary>
			/// Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
			/// </summary>
			/// <returns type="Word.CustomProperty"/>
		}

		CustomProperty.prototype.untrack = function() {
			/// <summary>
			/// Release the memory associated with this object, if has previous been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
			/// </summary>
			/// <returns type="Word.CustomProperty"/>
		}

		return CustomProperty;
	})(OfficeExtension.ClientObject);
	Word.CustomProperty = CustomProperty;
})(Word || (Word = {__proto__: null}));

var Word;
(function (Word) {
	var CustomPropertyCollection = (function(_super) {
		__extends(CustomPropertyCollection, _super);
		function CustomPropertyCollection() {
			/// <summary> Contains the collection of {@link Word.CustomProperty} objects. [Api set: WordApi 1.3] </summary>
			/// <field name="context" type="Word.RequestContext">The request context associated with this object.</field>
			/// <field name="isNull" type="Boolean">Returns a boolean value for whether the corresponding object is null. You must call "context.sync()" before reading the isNull property.</field>
			/// <field name="items" type="Array" elementType="Word.CustomProperty">Gets the loaded child items in this collection.</field>
		}

		CustomPropertyCollection.prototype.load = function(option) {
			/// <summary>
			/// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
			/// </summary>
			/// <param name="option" type="string | string[] | OfficeExtension.LoadOption"/>
			/// <returns type="Word.CustomPropertyCollection"/>
		}
		CustomPropertyCollection.prototype.add = function(key, value) {
			/// <summary>
			/// Creates a new or sets an existing custom property. [Api set: WordApi 1.3]
			/// </summary>
			/// <param name="key" type="String">Required. The custom property&apos;s key, which is case-insensitive.</param>
			/// <param name="value" >Required. The custom property&apos;s value.</param>
			/// <returns type="Word.CustomProperty"></returns>
		}
		CustomPropertyCollection.prototype.deleteAll = function() {
			/// <summary>
			/// Deletes all custom properties in this collection. [Api set: WordApi 1.3]
			/// </summary>
			/// <returns ></returns>
		}
		CustomPropertyCollection.prototype.getCount = function() {
			/// <summary>
			/// Gets the count of custom properties. [Api set: WordApi 1.3]
			/// </summary>
			/// <returns type="OfficeExtension.ClientResult&lt;number&gt;"></returns>
			var result = new OfficeExtension.ClientResult();
			result.__proto__ = null;
			result.value = 0;
			return result;
		}
		CustomPropertyCollection.prototype.getItem = function(key) {
			/// <summary>
			/// Gets a custom property object by its key, which is case-insensitive. Throws if the custom property does not exist. [Api set: WordApi 1.3]
			/// </summary>
			/// <param name="key" type="String">The key that identifies the custom property object.</param>
			/// <returns type="Word.CustomProperty"></returns>
		}
		CustomPropertyCollection.prototype.getItemOrNullObject = function(key) {
			/// <summary>
			/// Gets a custom property object by its key, which is case-insensitive. Returns a null object if the custom property does not exist. [Api set: WordApi 1.3]
			/// </summary>
			/// <param name="key" type="String">Required. The key that identifies the custom property object.</param>
			/// <returns type="Word.CustomProperty"></returns>
		}

		CustomPropertyCollection.prototype.track = function() {
			/// <summary>
			/// Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
			/// </summary>
			/// <returns type="Word.CustomPropertyCollection"/>
		}

		CustomPropertyCollection.prototype.untrack = function() {
			/// <summary>
			/// Release the memory associated with this object, if has previous been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
			/// </summary>
			/// <returns type="Word.CustomPropertyCollection"/>
		}

		return CustomPropertyCollection;
	})(OfficeExtension.ClientObject);
	Word.CustomPropertyCollection = CustomPropertyCollection;
})(Word || (Word = {__proto__: null}));

var Word;
(function (Word) {
	var CustomXmlPart = (function(_super) {
		__extends(CustomXmlPart, _super);
		function CustomXmlPart() {
			/// <summary> Represents a custom XML part. [Api set: WordApi 1.4] </summary>
			/// <field name="context" type="Word.RequestContext">The request context associated with this object.</field>
			/// <field name="isNull" type="Boolean">Returns a boolean value for whether the corresponding object is null. You must call "context.sync()" before reading the isNull property.</field>
			/// <field name="id" type="String">Gets the ID of the custom XML part. Read only. [Api set: WordApi 1.4]</field>
			/// <field name="namespaceUri" type="String">Gets the namespace URI of the custom XML part. Read only. [Api set: WordApi 1.4]</field>
		}

		CustomXmlPart.prototype.load = function(option) {
			/// <summary>
			/// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
			/// </summary>
			/// <param name="option" type="string | string[] | OfficeExtension.LoadOption"/>
			/// <returns type="Word.CustomXmlPart"/>
		}
		CustomXmlPart.prototype.delete = function() {
			/// <summary>
			/// Deletes the custom XML part. [Api set: WordApi 1.4]
			/// </summary>
			/// <returns ></returns>
		}
		CustomXmlPart.prototype.deleteAttribute = function(xpath, namespaceMappings, name) {
			/// <summary>
			/// Deletes an attribute with the given name from the element identified by xpath. [Api set: WordApi 1.4]
			/// </summary>
			/// <param name="xpath" type="String">Required. Absolute path to the single element in XPath notation.</param>
			/// <param name="namespaceMappings" >Required. An object whose properties represent namespace aliases and the values are the actual namespace URIs.</param>
			/// <param name="name" type="String">Required. Name of the attribute.</param>
			/// <returns ></returns>
		}
		CustomXmlPart.prototype.deleteElement = function(xpath, namespaceMappings) {
			/// <summary>
			/// Deletes the element identified by xpath. [Api set: WordApi 1.4]
			/// </summary>
			/// <param name="xpath" type="String">Required. Absolute path to the single element in XPath notation.</param>
			/// <param name="namespaceMappings" >Required. An object whose properties represent namespace aliases and the values are the actual namespace URIs.</param>
			/// <returns ></returns>
		}
		CustomXmlPart.prototype.getXml = function() {
			/// <summary>
			/// Gets the full XML content of the custom XML part. [Api set: WordApi 1.4]
			/// </summary>
			/// <returns type="OfficeExtension.ClientResult&lt;string&gt;"></returns>
			var result = new OfficeExtension.ClientResult();
			result.__proto__ = null;
			result.value = '';
			return result;
		}
		CustomXmlPart.prototype.insertAttribute = function(xpath, namespaceMappings, name, value) {
			/// <summary>
			/// Inserts an attribute with the given name and value to the element identified by xpath. [Api set: WordApi 1.4]
			/// </summary>
			/// <param name="xpath" type="String">Required. Absolute path to the single element in XPath notation.</param>
			/// <param name="namespaceMappings" >Required. An object whose properties represent namespace aliases and the values are the actual namespace URIs.</param>
			/// <param name="name" type="String">Required. Name of the attribute.</param>
			/// <param name="value" type="String">Required. Value of the attribute.</param>
			/// <returns ></returns>
		}
		CustomXmlPart.prototype.insertElement = function(xpath, xml, namespaceMappings, index) {
			/// <summary>
			/// Inserts the given XML under the parent element identified by xpath at child position index. [Api set: WordApi 1.4]
			/// </summary>
			/// <param name="xpath" type="String">Required. Absolute path to the single parent element in XPath notation.</param>
			/// <param name="xml" type="String">Required. XML content to be inserted.</param>
			/// <param name="namespaceMappings" >Required. An object whose properties represent namespace aliases and the values are the actual namespace URIs.</param>
			/// <param name="index" type="Number" optional="true">Optional. Zero-based position at which the new XML to be inserted. If omitted, the XML will be appended as the last child of this parent.</param>
			/// <returns ></returns>
		}
		CustomXmlPart.prototype.query = function(xpath, namespaceMappings) {
			/// <summary>
			/// Queries the XML content of the custom XML part. [Api set: WordApi 1.4]
			/// </summary>
			/// <param name="xpath" type="String">Required. An XPath query.</param>
			/// <param name="namespaceMappings" >Required. An object whose properties represent namespace aliases and the values are the actual namespace URIs.</param>
			/// <returns type="OfficeExtension.ClientResult&lt;string[]&gt;">An array where each item represents an entry matched by the XPath query.</returns>
			var result = new OfficeExtension.ClientResult();
			result.__proto__ = null;
			result.value = [];
			return result;
		}
		CustomXmlPart.prototype.setXml = function(xml) {
			/// <summary>
			/// Sets the full XML content of the custom XML part. [Api set: WordApi 1.4]
			/// </summary>
			/// <param name="xml" type="String">Required. XML content to be set.</param>
			/// <returns ></returns>
		}
		CustomXmlPart.prototype.updateAttribute = function(xpath, namespaceMappings, name, value) {
			/// <summary>
			/// Updates the value of an attribute with the given name of the element identified by xpath. [Api set: WordApi 1.4]
			/// </summary>
			/// <param name="xpath" type="String">Required. Absolute path to the single element in XPath notation.</param>
			/// <param name="namespaceMappings" >Required. An object whose properties represent namespace aliases and the values are the actual namespace URIs.</param>
			/// <param name="name" type="String">Required. Name of the attribute.</param>
			/// <param name="value" type="String">Required. New value of the attribute.</param>
			/// <returns ></returns>
		}
		CustomXmlPart.prototype.updateElement = function(xpath, xml, namespaceMappings) {
			/// <summary>
			/// Updates the XML of the element identified by xpath. [Api set: WordApi 1.4]
			/// </summary>
			/// <param name="xpath" type="String">Required. Absolute path to the single element in XPath notation.</param>
			/// <param name="xml" type="String">Required. New XML content to be stored.</param>
			/// <param name="namespaceMappings" >Required. An object whose properties represent namespace aliases and the values are the actual namespace URIs.</param>
			/// <returns ></returns>
		}

		CustomXmlPart.prototype.track = function() {
			/// <summary>
			/// Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
			/// </summary>
			/// <returns type="Word.CustomXmlPart"/>
		}

		CustomXmlPart.prototype.untrack = function() {
			/// <summary>
			/// Release the memory associated with this object, if has previous been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
			/// </summary>
			/// <returns type="Word.CustomXmlPart"/>
		}

		return CustomXmlPart;
	})(OfficeExtension.ClientObject);
	Word.CustomXmlPart = CustomXmlPart;
})(Word || (Word = {__proto__: null}));

var Word;
(function (Word) {
	var CustomXmlPartCollection = (function(_super) {
		__extends(CustomXmlPartCollection, _super);
		function CustomXmlPartCollection() {
			/// <summary> Contains the collection of {@link Word.CustomXmlPart} objects. [Api set: WordApi 1.4] </summary>
			/// <field name="context" type="Word.RequestContext">The request context associated with this object.</field>
			/// <field name="isNull" type="Boolean">Returns a boolean value for whether the corresponding object is null. You must call "context.sync()" before reading the isNull property.</field>
			/// <field name="items" type="Array" elementType="Word.CustomXmlPart">Gets the loaded child items in this collection.</field>
		}

		CustomXmlPartCollection.prototype.load = function(option) {
			/// <summary>
			/// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
			/// </summary>
			/// <param name="option" type="string | string[] | OfficeExtension.LoadOption"/>
			/// <returns type="Word.CustomXmlPartCollection"/>
		}
		CustomXmlPartCollection.prototype.add = function(xml) {
			/// <summary>
			/// Adds a new custom XML part to the document. [Api set: WordApi 1.4]
			/// </summary>
			/// <param name="xml" type="String">Required. XML content. Must be a valid XML fragment.</param>
			/// <returns type="Word.CustomXmlPart"></returns>
		}
		CustomXmlPartCollection.prototype.getByNamespace = function(namespaceUri) {
			/// <summary>
			/// Gets a new scoped collection of custom XML parts whose namespaces match the given namespace. [Api set: WordApi 1.4]
			/// </summary>
			/// <param name="namespaceUri" type="String">Required. The namespace URI.</param>
			/// <returns type="Word.CustomXmlPartScopedCollection"></returns>
		}
		CustomXmlPartCollection.prototype.getCount = function() {
			/// <summary>
			/// Gets the number of items in the collection. [Api set: WordApi 1.4]
			/// </summary>
			/// <returns type="OfficeExtension.ClientResult&lt;number&gt;"></returns>
			var result = new OfficeExtension.ClientResult();
			result.__proto__ = null;
			result.value = 0;
			return result;
		}
		CustomXmlPartCollection.prototype.getItem = function(id) {
			/// <summary>
			/// Gets a custom XML part based on its ID. Read only. [Api set: WordApi 1.4]
			/// </summary>
			/// <param name="id" type="String">ID or index of the custom XML part to be retrieved.</param>
			/// <returns type="Word.CustomXmlPart"></returns>
		}
		CustomXmlPartCollection.prototype.getItemOrNullObject = function(id) {
			/// <summary>
			/// Gets a custom XML part based on its ID. Returns a null object if the CustomXmlPart does not exist. [Api set: WordApi 1.4]
			/// </summary>
			/// <param name="id" type="String">Required. ID of the object to be retrieved.</param>
			/// <returns type="Word.CustomXmlPart"></returns>
		}

		CustomXmlPartCollection.prototype.track = function() {
			/// <summary>
			/// Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
			/// </summary>
			/// <returns type="Word.CustomXmlPartCollection"/>
		}

		CustomXmlPartCollection.prototype.untrack = function() {
			/// <summary>
			/// Release the memory associated with this object, if has previous been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
			/// </summary>
			/// <returns type="Word.CustomXmlPartCollection"/>
		}

		return CustomXmlPartCollection;
	})(OfficeExtension.ClientObject);
	Word.CustomXmlPartCollection = CustomXmlPartCollection;
})(Word || (Word = {__proto__: null}));

var Word;
(function (Word) {
	var CustomXmlPartScopedCollection = (function(_super) {
		__extends(CustomXmlPartScopedCollection, _super);
		function CustomXmlPartScopedCollection() {
			/// <summary> Contains the collection of {@link Word.CustomXmlPart} objects with a specific namespace. [Api set: WordApi 1.4] </summary>
			/// <field name="context" type="Word.RequestContext">The request context associated with this object.</field>
			/// <field name="isNull" type="Boolean">Returns a boolean value for whether the corresponding object is null. You must call "context.sync()" before reading the isNull property.</field>
			/// <field name="items" type="Array" elementType="Word.CustomXmlPart">Gets the loaded child items in this collection.</field>
		}

		CustomXmlPartScopedCollection.prototype.load = function(option) {
			/// <summary>
			/// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
			/// </summary>
			/// <param name="option" type="string | string[] | OfficeExtension.LoadOption"/>
			/// <returns type="Word.CustomXmlPartScopedCollection"/>
		}
		CustomXmlPartScopedCollection.prototype.getCount = function() {
			/// <summary>
			/// Gets the number of items in the collection. [Api set: WordApi 1.4]
			/// </summary>
			/// <returns type="OfficeExtension.ClientResult&lt;number&gt;"></returns>
			var result = new OfficeExtension.ClientResult();
			result.__proto__ = null;
			result.value = 0;
			return result;
		}
		CustomXmlPartScopedCollection.prototype.getItem = function(id) {
			/// <summary>
			/// Gets a custom XML part based on its ID. Read only. [Api set: WordApi 1.4]
			/// </summary>
			/// <param name="id" type="String">ID of the custom XML part to be retrieved.</param>
			/// <returns type="Word.CustomXmlPart"></returns>
		}
		CustomXmlPartScopedCollection.prototype.getItemOrNullObject = function(id) {
			/// <summary>
			/// Gets a custom XML part based on its ID. Returns a null object if the CustomXmlPart does not exist in the collection. [Api set: WordApi 1.4]
			/// </summary>
			/// <param name="id" type="String">Required. ID of the object to be retrieved.</param>
			/// <returns type="Word.CustomXmlPart"></returns>
		}
		CustomXmlPartScopedCollection.prototype.getOnlyItem = function() {
			/// <summary>
			/// If the collection contains exactly one item, this method returns it. Otherwise, this method produces an error. [Api set: WordApi 1.4]
			/// </summary>
			/// <returns type="Word.CustomXmlPart"></returns>
		}
		CustomXmlPartScopedCollection.prototype.getOnlyItemOrNullObject = function() {
			/// <summary>
			/// If the collection contains exactly one item, this method returns it. Otherwise, this method returns a null object. [Api set: WordApi 1.4]
			/// </summary>
			/// <returns type="Word.CustomXmlPart"></returns>
		}

		CustomXmlPartScopedCollection.prototype.track = function() {
			/// <summary>
			/// Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
			/// </summary>
			/// <returns type="Word.CustomXmlPartScopedCollection"/>
		}

		CustomXmlPartScopedCollection.prototype.untrack = function() {
			/// <summary>
			/// Release the memory associated with this object, if has previous been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
			/// </summary>
			/// <returns type="Word.CustomXmlPartScopedCollection"/>
		}

		return CustomXmlPartScopedCollection;
	})(OfficeExtension.ClientObject);
	Word.CustomXmlPartScopedCollection = CustomXmlPartScopedCollection;
})(Word || (Word = {__proto__: null}));

var Word;
(function (Word) {
	var Document = (function(_super) {
		__extends(Document, _super);
		function Document() {
			/// <summary> The Document object is the top level object. A Document object contains one or more sections, content controls, and the body that contains the contents of the document. [Api set: WordApi 1.1] </summary>
			/// <field name="context" type="Word.RequestContext">The request context associated with this object.</field>
			/// <field name="isNull" type="Boolean">Returns a boolean value for whether the corresponding object is null. You must call "context.sync()" before reading the isNull property.</field>
			/// <field name="body" type="Word.Body">Gets the body object of the document. The body is the text that excludes headers, footers, footnotes, textboxes, etc.. Read-only. [Api set: WordApi 1.1]</field>
			/// <field name="contentControls" type="Word.ContentControlCollection">Gets the collection of content control objects in the document. This includes content controls in the body of the document, headers, footers, textboxes, etc.. Read-only. [Api set: WordApi 1.1]</field>
			/// <field name="customXmlParts" type="Word.CustomXmlPartCollection">Gets the custom XML parts in the document. Read-only. [Api set: WordApi 1.4]</field>
			/// <field name="properties" type="Word.DocumentProperties">Gets the properties of the document. Read-only. [Api set: WordApi 1.3]</field>
			/// <field name="saved" type="Boolean">Indicates whether the changes in the document have been saved. A value of true indicates that the document hasn&apos;t changed since it was saved. Read-only. [Api set: WordApi 1.1]</field>
			/// <field name="sections" type="Word.SectionCollection">Gets the collection of section objects in the document. Read-only. [Api set: WordApi 1.1]</field>
			/// <field name="settings" type="Word.SettingCollection">Gets the add-in&apos;s settings in the document. Read-only. [Api set: WordApi 1.4]</field>
			/// <field name="onContentControlAdded" type="OfficeExtension.EventHandlers">Occurs when a content control is added. Run context.sync() in the handler to get the new content control&apos;s properties. [Api set: WordApi 1.4]</field>
		}

		Document.prototype.load = function(option) {
			/// <summary>
			/// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
			/// </summary>
			/// <param name="option" type="string | string[] | OfficeExtension.LoadOption"/>
			/// <returns type="Word.Document"/>
		}

		Document.prototype.set = function() {
			/// <signature>
			/// <summary>Sets multiple properties on the object at the same time, based on JSON input.</summary>
			/// <param name="properties" type="Word.Interfaces.DocumentUpdateData">Properties described by the Word.Interfaces.DocumentUpdateData interface.</param>
			/// <param name="options" type="string">Options of the form { throwOnReadOnly?: boolean }
			/// <br />
			/// * throwOnReadOnly: Throw an error if the passed-in property list includes read-only properties (default = true).
			/// </param>
			/// </signature>
			/// <signature>
			/// <summary>Sets multiple properties on the object at the same time, based on an existing loaded object.</summary>
			/// <param name="properties" type="Document">An existing Document object, with properties that have already been loaded and synced.</param>
			/// </signature>
		}
		Document.prototype.deleteBookmark = function(name) {
			/// <summary>
			/// Deletes a bookmark, if exists, from the document. [Api set: WordApi 1.4]
			/// </summary>
			/// <param name="name" type="String">Required. The bookmark name, which is case-insensitive.</param>
			/// <returns ></returns>
		}
		Document.prototype.getBookmarkRange = function(name) {
			/// <summary>
			/// Gets a bookmark&apos;s range. Throws if the bookmark does not exist. [Api set: WordApi 1.4]
			/// </summary>
			/// <param name="name" type="String">Required. The bookmark name, which is case-insensitive.</param>
			/// <returns type="Word.Range"></returns>
		}
		Document.prototype.getBookmarkRangeOrNullObject = function(name) {
			/// <summary>
			/// Gets a bookmark&apos;s range. Returns a null object if the bookmark does not exist. [Api set: WordApi 1.4]
			/// </summary>
			/// <param name="name" type="String">Required. The bookmark name, which is case-insensitive.</param>
			/// <returns type="Word.Range"></returns>
		}
		Document.prototype.getSelection = function() {
			/// <summary>
			/// Gets the current selection of the document. Multiple selections are not supported. [Api set: WordApi 1.1]
			/// </summary>
			/// <returns type="Word.Range"></returns>
		}
		Document.prototype.save = function() {
			/// <summary>
			/// Saves the document. This will use the Word default file naming convention if the document has not been saved before. [Api set: WordApi 1.1]
			/// </summary>
			/// <returns ></returns>
		}
		Document.prototype.onContentControlAdded = {
			__proto__: null,
			add: function (handler) {
				/// <param name="handler" type="function(eventArgs: Word.Interfaces.ContentControlEventArgs)">Handler for the event. EventArgs: Provides information about the content control that raised an event. </param>
				/// <returns type="OfficeExtension.EventHandlerResult"></returns>
				var eventInfo = new Word.Interfaces.ContentControlEventArgs();
				eventInfo.__proto__ = null;
				handler(eventInfo);
			},
			remove: function (handler) {
				/// <param name="handler" type="function(eventArgs: Word.Interfaces.ContentControlEventArgs)">Handler for the event.</param>
				return;
			}
		};

		Document.prototype.track = function() {
			/// <summary>
			/// Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
			/// </summary>
			/// <returns type="Word.Document"/>
		}

		Document.prototype.untrack = function() {
			/// <summary>
			/// Release the memory associated with this object, if has previous been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
			/// </summary>
			/// <returns type="Word.Document"/>
		}

		return Document;
	})(OfficeExtension.ClientObject);
	Word.Document = Document;
})(Word || (Word = {__proto__: null}));

var Word;
(function (Word) {
	var DocumentCreated = (function(_super) {
		__extends(DocumentCreated, _super);
		function DocumentCreated() {
			/// <summary> The DocumentCreated object is the top level object created by Application.CreateDocument. A DocumentCreated object is a special Document object. [Api set: WordApi 1.3] </summary>
			/// <field name="context" type="Word.RequestContext">The request context associated with this object.</field>
			/// <field name="isNull" type="Boolean">Returns a boolean value for whether the corresponding object is null. You must call "context.sync()" before reading the isNull property.</field>
			/// <field name="body" type="Word.Body">Gets the body object of the document. The body is the text that excludes headers, footers, footnotes, textboxes, etc.. Read-only. [Api set: WordApiHiddenDocument 1.3]</field>
			/// <field name="contentControls" type="Word.ContentControlCollection">Gets the collection of content control objects in the document. This includes content controls in the body of the document, headers, footers, textboxes, etc.. Read-only. [Api set: WordApiHiddenDocument 1.3]</field>
			/// <field name="customXmlParts" type="Word.CustomXmlPartCollection">Gets the custom XML parts in the document. Read-only. [Api set: WordApiHiddenDocument 1.4]</field>
			/// <field name="properties" type="Word.DocumentProperties">Gets the properties of the document. Read-only. [Api set: WordApiHiddenDocument 1.3]</field>
			/// <field name="saved" type="Boolean">Indicates whether the changes in the document have been saved. A value of true indicates that the document hasn&apos;t changed since it was saved. Read-only. [Api set: WordApiHiddenDocument 1.3]</field>
			/// <field name="sections" type="Word.SectionCollection">Gets the collection of section objects in the document. Read-only. [Api set: WordApiHiddenDocument 1.3]</field>
			/// <field name="settings" type="Word.SettingCollection">Gets the add-in&apos;s settings in the document. Read-only. [Api set: WordApiHiddenDocument 1.4]</field>
		}

		DocumentCreated.prototype.load = function(option) {
			/// <summary>
			/// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
			/// </summary>
			/// <param name="option" type="string | string[] | OfficeExtension.LoadOption"/>
			/// <returns type="Word.DocumentCreated"/>
		}

		DocumentCreated.prototype.set = function() {
			/// <signature>
			/// <summary>Sets multiple properties on the object at the same time, based on JSON input.</summary>
			/// <param name="properties" type="Word.Interfaces.DocumentCreatedUpdateData">Properties described by the Word.Interfaces.DocumentCreatedUpdateData interface.</param>
			/// <param name="options" type="string">Options of the form { throwOnReadOnly?: boolean }
			/// <br />
			/// * throwOnReadOnly: Throw an error if the passed-in property list includes read-only properties (default = true).
			/// </param>
			/// </signature>
			/// <signature>
			/// <summary>Sets multiple properties on the object at the same time, based on an existing loaded object.</summary>
			/// <param name="properties" type="DocumentCreated">An existing DocumentCreated object, with properties that have already been loaded and synced.</param>
			/// </signature>
		}
		DocumentCreated.prototype.deleteBookmark = function(name) {
			/// <summary>
			/// Deletes a bookmark, if exists, from the document. [Api set: WordApiHiddenDocument 1.4]
			/// </summary>
			/// <param name="name" type="String">Required. The bookmark name, which is case-insensitive.</param>
			/// <returns ></returns>
		}
		DocumentCreated.prototype.getBookmarkRange = function(name) {
			/// <summary>
			/// Gets a bookmark&apos;s range. Throws if the bookmark does not exist. [Api set: WordApiHiddenDocument 1.4]
			/// </summary>
			/// <param name="name" type="String">Required. The bookmark name, which is case-insensitive.</param>
			/// <returns type="Word.Range"></returns>
		}
		DocumentCreated.prototype.getBookmarkRangeOrNullObject = function(name) {
			/// <summary>
			/// Gets a bookmark&apos;s range. Returns a null object if the bookmark does not exist. [Api set: WordApiHiddenDocument 1.4]
			/// </summary>
			/// <param name="name" type="String">Required. The bookmark name, which is case-insensitive.</param>
			/// <returns type="Word.Range"></returns>
		}
		DocumentCreated.prototype.open = function() {
			/// <summary>
			/// Opens the document. [Api set: WordApi 1.3]
			/// </summary>
			/// <returns ></returns>
		}
		DocumentCreated.prototype.save = function() {
			/// <summary>
			/// Saves the document. This will use the Word default file naming convention if the document has not been saved before. [Api set: WordApiHiddenDocument 1.3]
			/// </summary>
			/// <returns ></returns>
		}

		DocumentCreated.prototype.track = function() {
			/// <summary>
			/// Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
			/// </summary>
			/// <returns type="Word.DocumentCreated"/>
		}

		DocumentCreated.prototype.untrack = function() {
			/// <summary>
			/// Release the memory associated with this object, if has previous been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
			/// </summary>
			/// <returns type="Word.DocumentCreated"/>
		}

		return DocumentCreated;
	})(OfficeExtension.ClientObject);
	Word.DocumentCreated = DocumentCreated;
})(Word || (Word = {__proto__: null}));

var Word;
(function (Word) {
	var DocumentProperties = (function(_super) {
		__extends(DocumentProperties, _super);
		function DocumentProperties() {
			/// <summary> Represents document properties. [Api set: WordApi 1.3] </summary>
			/// <field name="context" type="Word.RequestContext">The request context associated with this object.</field>
			/// <field name="isNull" type="Boolean">Returns a boolean value for whether the corresponding object is null. You must call "context.sync()" before reading the isNull property.</field>
			/// <field name="applicationName" type="String">Gets the application name of the document. Read only. [Api set: WordApi 1.3]</field>
			/// <field name="author" type="String">Gets or sets the author of the document. [Api set: WordApi 1.3]</field>
			/// <field name="category" type="String">Gets or sets the category of the document. [Api set: WordApi 1.3]</field>
			/// <field name="comments" type="String">Gets or sets the comments of the document. [Api set: WordApi 1.3]</field>
			/// <field name="company" type="String">Gets or sets the company of the document. [Api set: WordApi 1.3]</field>
			/// <field name="creationDate" type="Date">Gets the creation date of the document. Read only. [Api set: WordApi 1.3]</field>
			/// <field name="customProperties" type="Word.CustomPropertyCollection">Gets the collection of custom properties of the document. Read only. [Api set: WordApi 1.3]</field>
			/// <field name="format" type="String">Gets or sets the format of the document. [Api set: WordApi 1.3]</field>
			/// <field name="keywords" type="String">Gets or sets the keywords of the document. [Api set: WordApi 1.3]</field>
			/// <field name="lastAuthor" type="String">Gets the last author of the document. Read only. [Api set: WordApi 1.3]</field>
			/// <field name="lastPrintDate" type="Date">Gets the last print date of the document. Read only. [Api set: WordApi 1.3]</field>
			/// <field name="lastSaveTime" type="Date">Gets the last save time of the document. Read only. [Api set: WordApi 1.3]</field>
			/// <field name="manager" type="String">Gets or sets the manager of the document. [Api set: WordApi 1.3]</field>
			/// <field name="revisionNumber" type="String">Gets the revision number of the document. Read only. [Api set: WordApi 1.3]</field>
			/// <field name="security" type="Number">Gets the security of the document. Read only. [Api set: WordApi 1.3]</field>
			/// <field name="subject" type="String">Gets or sets the subject of the document. [Api set: WordApi 1.3]</field>
			/// <field name="template" type="String">Gets the template of the document. Read only. [Api set: WordApi 1.3]</field>
			/// <field name="title" type="String">Gets or sets the title of the document. [Api set: WordApi 1.3]</field>
		}

		DocumentProperties.prototype.load = function(option) {
			/// <summary>
			/// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
			/// </summary>
			/// <param name="option" type="string | string[] | OfficeExtension.LoadOption"/>
			/// <returns type="Word.DocumentProperties"/>
		}

		DocumentProperties.prototype.set = function() {
			/// <signature>
			/// <summary>Sets multiple properties on the object at the same time, based on JSON input.</summary>
			/// <param name="properties" type="Word.Interfaces.DocumentPropertiesUpdateData">Properties described by the Word.Interfaces.DocumentPropertiesUpdateData interface.</param>
			/// <param name="options" type="string">Options of the form { throwOnReadOnly?: boolean }
			/// <br />
			/// * throwOnReadOnly: Throw an error if the passed-in property list includes read-only properties (default = true).
			/// </param>
			/// </signature>
			/// <signature>
			/// <summary>Sets multiple properties on the object at the same time, based on an existing loaded object.</summary>
			/// <param name="properties" type="DocumentProperties">An existing DocumentProperties object, with properties that have already been loaded and synced.</param>
			/// </signature>
		}

		DocumentProperties.prototype.track = function() {
			/// <summary>
			/// Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
			/// </summary>
			/// <returns type="Word.DocumentProperties"/>
		}

		DocumentProperties.prototype.untrack = function() {
			/// <summary>
			/// Release the memory associated with this object, if has previous been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
			/// </summary>
			/// <returns type="Word.DocumentProperties"/>
		}

		return DocumentProperties;
	})(OfficeExtension.ClientObject);
	Word.DocumentProperties = DocumentProperties;
})(Word || (Word = {__proto__: null}));

var Word;
(function (Word) {
	/// <summary> [Api set: WordApi] </summary>
	var DocumentPropertyType = {
		__proto__: null,
		"string": "string",
		"number": "number",
		"date": "date",
		"boolean": "boolean",
	}
	Word.DocumentPropertyType = DocumentPropertyType;
})(Word || (Word = {__proto__: null}));

var Word;
(function (Word) {
	/// <summary> Provides information about the type of a raised event. For each object type, please keep the order of: deleted, selection changed, data changed, added. [Api set: WordApi] </summary>
	var EventType = {
		__proto__: null,
		"contentControlDeleted": "contentControlDeleted",
		"contentControlSelectionChanged": "contentControlSelectionChanged",
		"contentControlDataChanged": "contentControlDataChanged",
		"contentControlAdded": "contentControlAdded",
	}
	Word.EventType = EventType;
})(Word || (Word = {__proto__: null}));

var Word;
(function (Word) {
	/// <summary> [Api set: WordApi] </summary>
	var FileContentFormat = {
		__proto__: null,
		"base64": "base64",
		"html": "html",
		"ooxml": "ooxml",
	}
	Word.FileContentFormat = FileContentFormat;
})(Word || (Word = {__proto__: null}));

var Word;
(function (Word) {
	var Font = (function(_super) {
		__extends(Font, _super);
		function Font() {
			/// <summary> Represents a font. [Api set: WordApi 1.1] </summary>
			/// <field name="context" type="Word.RequestContext">The request context associated with this object.</field>
			/// <field name="isNull" type="Boolean">Returns a boolean value for whether the corresponding object is null. You must call "context.sync()" before reading the isNull property.</field>
			/// <field name="bold" type="Boolean">Gets or sets a value that indicates whether the font is bold. True if the font is formatted as bold, otherwise, false. [Api set: WordApi 1.1]</field>
			/// <field name="color" type="String">Gets or sets the color for the specified font. You can provide the value in the &apos;#RRGGBB&apos; format or the color name. [Api set: WordApi 1.1]</field>
			/// <field name="doubleStrikeThrough" type="Boolean">Gets or sets a value that indicates whether the font has a double strikethrough. True if the font is formatted as double strikethrough text, otherwise, false. [Api set: WordApi 1.1]</field>
			/// <field name="highlightColor" type="String">Gets or sets the highlight color. To set it, use a value either in the &apos;#RRGGBB&apos; format or the color name. To remove highlight color, set it to null. The returned highlight color can be in the &apos;#RRGGBB&apos; format, an empty string for mixed highlight colors, or null for no highlight color. [Api set: WordApi 1.1]</field>
			/// <field name="italic" type="Boolean">Gets or sets a value that indicates whether the font is italicized. True if the font is italicized, otherwise, false. [Api set: WordApi 1.1]</field>
			/// <field name="name" type="String">Gets or sets a value that represents the name of the font. [Api set: WordApi 1.1]</field>
			/// <field name="size" type="Number">Gets or sets a value that represents the font size in points. [Api set: WordApi 1.1]</field>
			/// <field name="strikeThrough" type="Boolean">Gets or sets a value that indicates whether the font has a strikethrough. True if the font is formatted as strikethrough text, otherwise, false. [Api set: WordApi 1.1]</field>
			/// <field name="subscript" type="Boolean">Gets or sets a value that indicates whether the font is a subscript. True if the font is formatted as subscript, otherwise, false. [Api set: WordApi 1.1]</field>
			/// <field name="superscript" type="Boolean">Gets or sets a value that indicates whether the font is a superscript. True if the font is formatted as superscript, otherwise, false. [Api set: WordApi 1.1]</field>
			/// <field name="underline" type="String">Gets or sets a value that indicates the font&apos;s underline type. &apos;None&apos; if the font is not underlined. [Api set: WordApi 1.1]</field>
		}

		Font.prototype.load = function(option) {
			/// <summary>
			/// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
			/// </summary>
			/// <param name="option" type="string | string[] | OfficeExtension.LoadOption"/>
			/// <returns type="Word.Font"/>
		}

		Font.prototype.set = function() {
			/// <signature>
			/// <summary>Sets multiple properties on the object at the same time, based on JSON input.</summary>
			/// <param name="properties" type="Word.Interfaces.FontUpdateData">Properties described by the Word.Interfaces.FontUpdateData interface.</param>
			/// <param name="options" type="string">Options of the form { throwOnReadOnly?: boolean }
			/// <br />
			/// * throwOnReadOnly: Throw an error if the passed-in property list includes read-only properties (default = true).
			/// </param>
			/// </signature>
			/// <signature>
			/// <summary>Sets multiple properties on the object at the same time, based on an existing loaded object.</summary>
			/// <param name="properties" type="Font">An existing Font object, with properties that have already been loaded and synced.</param>
			/// </signature>
		}

		Font.prototype.track = function() {
			/// <summary>
			/// Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
			/// </summary>
			/// <returns type="Word.Font"/>
		}

		Font.prototype.untrack = function() {
			/// <summary>
			/// Release the memory associated with this object, if has previous been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
			/// </summary>
			/// <returns type="Word.Font"/>
		}

		return Font;
	})(OfficeExtension.ClientObject);
	Word.Font = Font;
})(Word || (Word = {__proto__: null}));

var Word;
(function (Word) {
	/// <summary> [Api set: WordApi] </summary>
	var HeaderFooterType = {
		__proto__: null,
		"primary": "primary",
		"firstPage": "firstPage",
		"evenPages": "evenPages",
	}
	Word.HeaderFooterType = HeaderFooterType;
})(Word || (Word = {__proto__: null}));

var Word;
(function (Word) {
	/// <summary> [Api set: WordApi] </summary>
	var ImageFormat = {
		__proto__: null,
		"unsupported": "unsupported",
		"undefined": "undefined",
		"bmp": "bmp",
		"jpeg": "jpeg",
		"gif": "gif",
		"tiff": "tiff",
		"png": "png",
		"icon": "icon",
		"exif": "exif",
		"wmf": "wmf",
		"emf": "emf",
		"pict": "pict",
		"pdf": "pdf",
		"svg": "svg",
	}
	Word.ImageFormat = ImageFormat;
})(Word || (Word = {__proto__: null}));

var Word;
(function (Word) {
	var InlinePicture = (function(_super) {
		__extends(InlinePicture, _super);
		function InlinePicture() {
			/// <summary> Represents an inline picture. [Api set: WordApi 1.1] </summary>
			/// <field name="context" type="Word.RequestContext">The request context associated with this object.</field>
			/// <field name="isNull" type="Boolean">Returns a boolean value for whether the corresponding object is null. You must call "context.sync()" before reading the isNull property.</field>
			/// <field name="altTextDescription" type="String">Gets or sets a string that represents the alternative text associated with the inline image. [Api set: WordApi 1.1]</field>
			/// <field name="altTextTitle" type="String">Gets or sets a string that contains the title for the inline image. [Api set: WordApi 1.1]</field>
			/// <field name="height" type="Number">Gets or sets a number that describes the height of the inline image. [Api set: WordApi 1.1]</field>
			/// <field name="hyperlink" type="String">Gets or sets a hyperlink on the image. Use a &apos;#&apos; to separate the address part from the optional location part. [Api set: WordApi 1.1]</field>
			/// <field name="imageFormat" type="String">Gets the format of the inline image. Read-only. [Api set: WordApi 1.4]</field>
			/// <field name="lockAspectRatio" type="Boolean">Gets or sets a value that indicates whether the inline image retains its original proportions when you resize it. [Api set: WordApi 1.1]</field>
			/// <field name="paragraph" type="Word.Paragraph">Gets the parent paragraph that contains the inline image. Read-only. [Api set: WordApi 1.2]</field>
			/// <field name="parentContentControl" type="Word.ContentControl">Gets the content control that contains the inline image. Throws if there isn&apos;t a parent content control. Read-only. [Api set: WordApi 1.1]</field>
			/// <field name="parentContentControlOrNullObject" type="Word.ContentControl">Gets the content control that contains the inline image. Returns a null object if there isn&apos;t a parent content control. Read-only. [Api set: WordApi 1.3]</field>
			/// <field name="parentTable" type="Word.Table">Gets the table that contains the inline image. Throws if it is not contained in a table. Read-only. [Api set: WordApi 1.3]</field>
			/// <field name="parentTableCell" type="Word.TableCell">Gets the table cell that contains the inline image. Throws if it is not contained in a table cell. Read-only. [Api set: WordApi 1.3]</field>
			/// <field name="parentTableCellOrNullObject" type="Word.TableCell">Gets the table cell that contains the inline image. Returns a null object if it is not contained in a table cell. Read-only. [Api set: WordApi 1.3]</field>
			/// <field name="parentTableOrNullObject" type="Word.Table">Gets the table that contains the inline image. Returns a null object if it is not contained in a table. Read-only. [Api set: WordApi 1.3]</field>
			/// <field name="width" type="Number">Gets or sets a number that describes the width of the inline image. [Api set: WordApi 1.1]</field>
		}

		InlinePicture.prototype.load = function(option) {
			/// <summary>
			/// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
			/// </summary>
			/// <param name="option" type="string | string[] | OfficeExtension.LoadOption"/>
			/// <returns type="Word.InlinePicture"/>
		}

		InlinePicture.prototype.set = function() {
			/// <signature>
			/// <summary>Sets multiple properties on the object at the same time, based on JSON input.</summary>
			/// <param name="properties" type="Word.Interfaces.InlinePictureUpdateData">Properties described by the Word.Interfaces.InlinePictureUpdateData interface.</param>
			/// <param name="options" type="string">Options of the form { throwOnReadOnly?: boolean }
			/// <br />
			/// * throwOnReadOnly: Throw an error if the passed-in property list includes read-only properties (default = true).
			/// </param>
			/// </signature>
			/// <signature>
			/// <summary>Sets multiple properties on the object at the same time, based on an existing loaded object.</summary>
			/// <param name="properties" type="InlinePicture">An existing InlinePicture object, with properties that have already been loaded and synced.</param>
			/// </signature>
		}
		InlinePicture.prototype.delete = function() {
			/// <summary>
			/// Deletes the inline picture from the document. [Api set: WordApi 1.2]
			/// </summary>
			/// <returns ></returns>
		}
		InlinePicture.prototype.getBase64ImageSrc = function() {
			/// <summary>
			/// Gets the base64 encoded string representation of the inline image. [Api set: WordApi 1.1]
			/// </summary>
			/// <returns type="OfficeExtension.ClientResult&lt;string&gt;"></returns>
			var result = new OfficeExtension.ClientResult();
			result.__proto__ = null;
			result.value = '';
			return result;
		}
		InlinePicture.prototype.getNext = function() {
			/// <summary>
			/// Gets the next inline image. Throws if this inline image is the last one. [Api set: WordApi 1.3]
			/// </summary>
			/// <returns type="Word.InlinePicture"></returns>
		}
		InlinePicture.prototype.getNextOrNullObject = function() {
			/// <summary>
			/// Gets the next inline image. Returns a null object if this inline image is the last one. [Api set: WordApi 1.3]
			/// </summary>
			/// <returns type="Word.InlinePicture"></returns>
		}
		InlinePicture.prototype.getRange = function(rangeLocation) {
			/// <summary>
			/// Gets the picture, or the starting or ending point of the picture, as a range. [Api set: WordApi 1.3]
			/// </summary>
			/// <param name="rangeLocation" type="String" optional="true">Optional. The range location can be &apos;Whole&apos;, &apos;Start&apos;, or &apos;End&apos;.</param>
			/// <returns type="Word.Range"></returns>
		}
		InlinePicture.prototype.insertBreak = function(breakType, insertLocation) {
			/// <summary>
			/// Inserts a break at the specified location in the main document. The insertLocation value can be &apos;Before&apos; or &apos;After&apos;. [Api set: WordApi 1.2]
			/// </summary>
			/// <param name="breakType" type="String">Required. The break type to add.</param>
			/// <param name="insertLocation" type="String">Required. The value can be &apos;Before&apos; or &apos;After&apos;.</param>
			/// <returns ></returns>
		}
		InlinePicture.prototype.insertContentControl = function() {
			/// <summary>
			/// Wraps the inline picture with a rich text content control. [Api set: WordApi 1.1]
			/// </summary>
			/// <returns type="Word.ContentControl"></returns>
		}
		InlinePicture.prototype.insertFileFromBase64 = function(base64File, insertLocation) {
			/// <summary>
			/// Inserts a document at the specified location. The insertLocation value can be &apos;Before&apos; or &apos;After&apos;. [Api set: WordApi 1.2]
			/// </summary>
			/// <param name="base64File" type="String">Required. The base64 encoded content of a .docx file.</param>
			/// <param name="insertLocation" type="String">Required. The value can be &apos;Before&apos; or &apos;After&apos;.</param>
			/// <returns type="Word.Range"></returns>
		}
		InlinePicture.prototype.insertHtml = function(html, insertLocation) {
			/// <summary>
			/// Inserts HTML at the specified location. The insertLocation value can be &apos;Before&apos; or &apos;After&apos;. [Api set: WordApi 1.2]
			/// </summary>
			/// <param name="html" type="String">Required. The HTML to be inserted.</param>
			/// <param name="insertLocation" type="String">Required. The value can be &apos;Before&apos; or &apos;After&apos;.</param>
			/// <returns type="Word.Range"></returns>
		}
		InlinePicture.prototype.insertInlinePictureFromBase64 = function(base64EncodedImage, insertLocation) {
			/// <summary>
			/// Inserts an inline picture at the specified location. The insertLocation value can be &apos;Replace&apos;, &apos;Before&apos;, or &apos;After&apos;. [Api set: WordApi 1.2]
			/// </summary>
			/// <param name="base64EncodedImage" type="String">Required. The base64 encoded image to be inserted.</param>
			/// <param name="insertLocation" type="String">Required. The value can be &apos;Replace&apos;, &apos;Before&apos;, or &apos;After&apos;.</param>
			/// <returns type="Word.InlinePicture"></returns>
		}
		InlinePicture.prototype.insertOoxml = function(ooxml, insertLocation) {
			/// <summary>
			/// Inserts OOXML at the specified location.  The insertLocation value can be &apos;Before&apos; or &apos;After&apos;. [Api set: WordApi 1.2]
			/// </summary>
			/// <param name="ooxml" type="String">Required. The OOXML to be inserted.</param>
			/// <param name="insertLocation" type="String">Required. The value can be &apos;Before&apos; or &apos;After&apos;.</param>
			/// <returns type="Word.Range"></returns>
		}
		InlinePicture.prototype.insertParagraph = function(paragraphText, insertLocation) {
			/// <summary>
			/// Inserts a paragraph at the specified location. The insertLocation value can be &apos;Before&apos; or &apos;After&apos;. [Api set: WordApi 1.2]
			/// </summary>
			/// <param name="paragraphText" type="String">Required. The paragraph text to be inserted.</param>
			/// <param name="insertLocation" type="String">Required. The value can be &apos;Before&apos; or &apos;After&apos;.</param>
			/// <returns type="Word.Paragraph"></returns>
		}
		InlinePicture.prototype.insertText = function(text, insertLocation) {
			/// <summary>
			/// Inserts text at the specified location. The insertLocation value can be &apos;Before&apos; or &apos;After&apos;. [Api set: WordApi 1.2]
			/// </summary>
			/// <param name="text" type="String">Required. Text to be inserted.</param>
			/// <param name="insertLocation" type="String">Required. The value can be &apos;Before&apos; or &apos;After&apos;.</param>
			/// <returns type="Word.Range"></returns>
		}
		InlinePicture.prototype.select = function(selectionMode) {
			/// <summary>
			/// Selects the inline picture. This causes Word to scroll to the selection. [Api set: WordApi 1.2]
			/// </summary>
			/// <param name="selectionMode" type="String" optional="true">Optional. The selection mode can be &apos;Select&apos;, &apos;Start&apos;, or &apos;End&apos;. &apos;Select&apos; is the default.</param>
			/// <returns ></returns>
		}

		InlinePicture.prototype.track = function() {
			/// <summary>
			/// Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
			/// </summary>
			/// <returns type="Word.InlinePicture"/>
		}

		InlinePicture.prototype.untrack = function() {
			/// <summary>
			/// Release the memory associated with this object, if has previous been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
			/// </summary>
			/// <returns type="Word.InlinePicture"/>
		}

		return InlinePicture;
	})(OfficeExtension.ClientObject);
	Word.InlinePicture = InlinePicture;
})(Word || (Word = {__proto__: null}));

var Word;
(function (Word) {
	var InlinePictureCollection = (function(_super) {
		__extends(InlinePictureCollection, _super);
		function InlinePictureCollection() {
			/// <summary> Contains a collection of {@link Word.InlinePicture} objects. [Api set: WordApi 1.1] </summary>
			/// <field name="context" type="Word.RequestContext">The request context associated with this object.</field>
			/// <field name="isNull" type="Boolean">Returns a boolean value for whether the corresponding object is null. You must call "context.sync()" before reading the isNull property.</field>
			/// <field name="items" type="Array" elementType="Word.InlinePicture">Gets the loaded child items in this collection.</field>
		}

		InlinePictureCollection.prototype.load = function(option) {
			/// <summary>
			/// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
			/// </summary>
			/// <param name="option" type="string | string[] | OfficeExtension.LoadOption"/>
			/// <returns type="Word.InlinePictureCollection"/>
		}
		InlinePictureCollection.prototype.getFirst = function() {
			/// <summary>
			/// Gets the first inline image in this collection. Throws if this collection is empty. [Api set: WordApi 1.3]
			/// </summary>
			/// <returns type="Word.InlinePicture"></returns>
		}
		InlinePictureCollection.prototype.getFirstOrNullObject = function() {
			/// <summary>
			/// Gets the first inline image in this collection. Returns a null object if this collection is empty. [Api set: WordApi 1.3]
			/// </summary>
			/// <returns type="Word.InlinePicture"></returns>
		}

		InlinePictureCollection.prototype.track = function() {
			/// <summary>
			/// Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
			/// </summary>
			/// <returns type="Word.InlinePictureCollection"/>
		}

		InlinePictureCollection.prototype.untrack = function() {
			/// <summary>
			/// Release the memory associated with this object, if has previous been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
			/// </summary>
			/// <returns type="Word.InlinePictureCollection"/>
		}

		return InlinePictureCollection;
	})(OfficeExtension.ClientObject);
	Word.InlinePictureCollection = InlinePictureCollection;
})(Word || (Word = {__proto__: null}));

var Word;
(function (Word) {
	/// <summary> The insertion location types [Api set: WordApi] </summary>
	var InsertLocation = {
		__proto__: null,
		"before": "before",
		"after": "after",
		"start": "start",
		"end": "end",
		"replace": "replace",
	}
	Word.InsertLocation = InsertLocation;
})(Word || (Word = {__proto__: null}));

var Word;
(function (Word) {
	var List = (function(_super) {
		__extends(List, _super);
		function List() {
			/// <summary> Contains a collection of {@link Word.Paragraph} objects. [Api set: WordApi 1.3] </summary>
			/// <field name="context" type="Word.RequestContext">The request context associated with this object.</field>
			/// <field name="isNull" type="Boolean">Returns a boolean value for whether the corresponding object is null. You must call "context.sync()" before reading the isNull property.</field>
			/// <field name="id" type="Number">Gets the list&apos;s id. [Api set: WordApi 1.3]</field>
			/// <field name="levelExistences" type="Array" elementType="Boolean">Checks whether each of the 9 levels exists in the list. A true value indicates the level exists, which means there is at least one list item at that level. Read-only. [Api set: WordApi 1.3]</field>
			/// <field name="levelTypes" type="Array" elementType="String">Gets all 9 level types in the list. Each type can be &apos;Bullet&apos;, &apos;Number&apos;, or &apos;Picture&apos;. Read-only. [Api set: WordApi 1.3]</field>
			/// <field name="paragraphs" type="Word.ParagraphCollection">Gets paragraphs in the list. Read-only. [Api set: WordApi 1.3]</field>
		}

		List.prototype.load = function(option) {
			/// <summary>
			/// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
			/// </summary>
			/// <param name="option" type="string | string[] | OfficeExtension.LoadOption"/>
			/// <returns type="Word.List"/>
		}
		List.prototype.getLevelFont = function(level) {
			/// <summary>
			/// Gets the font of the bullet, number or picture at the specified level in the list. [Api set: WordApi 1.4]
			/// </summary>
			/// <param name="level" type="Number">Required. The level in the list.</param>
			/// <returns type="Word.Font"></returns>
		}
		List.prototype.getLevelParagraphs = function(level) {
			/// <summary>
			/// Gets the paragraphs that occur at the specified level in the list. [Api set: WordApi 1.3]
			/// </summary>
			/// <param name="level" type="Number">Required. The level in the list.</param>
			/// <returns type="Word.ParagraphCollection"></returns>
		}
		List.prototype.getLevelPicture = function(level) {
			/// <summary>
			/// Gets the base64 encoded string representation of the picture at the specified level in the list. [Api set: WordApi 1.4]
			/// </summary>
			/// <param name="level" type="Number">Required. The level in the list.</param>
			/// <returns type="OfficeExtension.ClientResult&lt;string&gt;"></returns>
			var result = new OfficeExtension.ClientResult();
			result.__proto__ = null;
			result.value = '';
			return result;
		}
		List.prototype.getLevelString = function(level) {
			/// <summary>
			/// Gets the bullet, number or picture at the specified level as a string. [Api set: WordApi 1.3]
			/// </summary>
			/// <param name="level" type="Number">Required. The level in the list.</param>
			/// <returns type="OfficeExtension.ClientResult&lt;string&gt;"></returns>
			var result = new OfficeExtension.ClientResult();
			result.__proto__ = null;
			result.value = '';
			return result;
		}
		List.prototype.insertParagraph = function(paragraphText, insertLocation) {
			/// <summary>
			/// Inserts a paragraph at the specified location. The insertLocation value can be &apos;Start&apos;, &apos;End&apos;, &apos;Before&apos;, or &apos;After&apos;. [Api set: WordApi 1.3]
			/// </summary>
			/// <param name="paragraphText" type="String">Required. The paragraph text to be inserted.</param>
			/// <param name="insertLocation" type="String">Required. The value can be &apos;Start&apos;, &apos;End&apos;, &apos;Before&apos;, or &apos;After&apos;.</param>
			/// <returns type="Word.Paragraph"></returns>
		}
		List.prototype.resetLevelFont = function(level, resetFontName) {
			/// <summary>
			/// Resets the font of the bullet, number or picture at the specified level in the list. [Api set: WordApi 1.4]
			/// </summary>
			/// <param name="level" type="Number">Required. The level in the list.</param>
			/// <param name="resetFontName" type="Boolean" optional="true">Optional. Indicates whether to reset the font name. Default is false that indicates the font name is kept unchanged.</param>
			/// <returns ></returns>
		}
		List.prototype.setLevelAlignment = function(level, alignment) {
			/// <summary>
			/// Sets the alignment of the bullet, number or picture at the specified level in the list. [Api set: WordApi 1.3]
			/// </summary>
			/// <param name="level" type="Number">Required. The level in the list.</param>
			/// <param name="alignment" type="String">Required. The level alignment that can be &apos;Left&apos;, &apos;Centered&apos;, or &apos;Right&apos;.</param>
			/// <returns ></returns>
		}
		List.prototype.setLevelBullet = function(level, listBullet, charCode, fontName) {
			/// <summary>
			/// Sets the bullet format at the specified level in the list. If the bullet is &apos;Custom&apos;, the charCode is required. [Api set: WordApi 1.3]
			/// </summary>
			/// <param name="level" type="Number">Required. The level in the list.</param>
			/// <param name="listBullet" type="String">Required. The bullet.</param>
			/// <param name="charCode" type="Number" optional="true">Optional. The bullet character&apos;s code value. Used only if the bullet is &apos;Custom&apos;.</param>
			/// <param name="fontName" type="String" optional="true">Optional. The bullet&apos;s font name. Used only if the bullet is &apos;Custom&apos;.</param>
			/// <returns ></returns>
		}
		List.prototype.setLevelIndents = function(level, textIndent, bulletNumberPictureIndent) {
			/// <summary>
			/// Sets the two indents of the specified level in the list. [Api set: WordApi 1.3]
			/// </summary>
			/// <param name="level" type="Number">Required. The level in the list.</param>
			/// <param name="textIndent" type="Number">Required. The text indent in points. It is the same as paragraph left indent.</param>
			/// <param name="bulletNumberPictureIndent" type="Number">Required. The relative indent, in points, of the bullet, number or picture. It is the same as paragraph first line indent.</param>
			/// <returns ></returns>
		}
		List.prototype.setLevelNumbering = function(level, listNumbering, formatString) {
			/// <summary>
			/// Sets the numbering format at the specified level in the list. [Api set: WordApi 1.3]
			/// </summary>
			/// <param name="level" type="Number">Required. The level in the list.</param>
			/// <param name="listNumbering" type="String">Required. The ordinal format.</param>
			/// <param name="formatString" type="Array"  optional="true">Optional. The numbering string format defined as an array of strings and/or integers. Each integer is a level of number type that is higher than or equal to this level. For example, an array of [&quot;(&quot;, level - 1, &quot;.&quot;, level, &quot;)&quot;] can define the format of &quot;(2.c)&quot;, where 2 is the parent&apos;s item number and c is this level&apos;s item number.</param>
			/// <returns ></returns>
		}
		List.prototype.setLevelPicture = function(level, base64EncodedImage) {
			/// <summary>
			/// Sets the picture at the specified level in the list. [Api set: WordApi 1.4]
			/// </summary>
			/// <param name="level" type="Number">Required. The level in the list.</param>
			/// <param name="base64EncodedImage" type="String" optional="true">Optional. The base64 encoded image to be set. If not given, the default picture is set.</param>
			/// <returns ></returns>
		}
		List.prototype.setLevelStartingNumber = function(level, startingNumber) {
			/// <summary>
			/// Sets the starting number at the specified level in the list. Default value is 1. [Api set: WordApi 1.3]
			/// </summary>
			/// <param name="level" type="Number">Required. The level in the list.</param>
			/// <param name="startingNumber" type="Number">Required. The number to start with.</param>
			/// <returns ></returns>
		}

		List.prototype.track = function() {
			/// <summary>
			/// Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
			/// </summary>
			/// <returns type="Word.List"/>
		}

		List.prototype.untrack = function() {
			/// <summary>
			/// Release the memory associated with this object, if has previous been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
			/// </summary>
			/// <returns type="Word.List"/>
		}

		return List;
	})(OfficeExtension.ClientObject);
	Word.List = List;
})(Word || (Word = {__proto__: null}));

var Word;
(function (Word) {
	/// <summary> [Api set: WordApi] </summary>
	var ListBullet = {
		__proto__: null,
		"custom": "custom",
		"solid": "solid",
		"hollow": "hollow",
		"square": "square",
		"diamonds": "diamonds",
		"arrow": "arrow",
		"checkmark": "checkmark",
	}
	Word.ListBullet = ListBullet;
})(Word || (Word = {__proto__: null}));

var Word;
(function (Word) {
	var ListCollection = (function(_super) {
		__extends(ListCollection, _super);
		function ListCollection() {
			/// <summary> Contains a collection of {@link Word.List} objects. [Api set: WordApi 1.3] </summary>
			/// <field name="context" type="Word.RequestContext">The request context associated with this object.</field>
			/// <field name="isNull" type="Boolean">Returns a boolean value for whether the corresponding object is null. You must call "context.sync()" before reading the isNull property.</field>
			/// <field name="items" type="Array" elementType="Word.List">Gets the loaded child items in this collection.</field>
		}

		ListCollection.prototype.load = function(option) {
			/// <summary>
			/// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
			/// </summary>
			/// <param name="option" type="string | string[] | OfficeExtension.LoadOption"/>
			/// <returns type="Word.ListCollection"/>
		}
		ListCollection.prototype.getById = function(id) {
			/// <summary>
			/// Gets a list by its identifier. Throws if there isn&apos;t a list with the identifier in this collection. [Api set: WordApi 1.3]
			/// </summary>
			/// <param name="id" type="Number">Required. A list identifier.</param>
			/// <returns type="Word.List"></returns>
		}
		ListCollection.prototype.getByIdOrNullObject = function(id) {
			/// <summary>
			/// Gets a list by its identifier. Returns a null object if there isn&apos;t a list with the identifier in this collection. [Api set: WordApi 1.3]
			/// </summary>
			/// <param name="id" type="Number">Required. A list identifier.</param>
			/// <returns type="Word.List"></returns>
		}
		ListCollection.prototype.getFirst = function() {
			/// <summary>
			/// Gets the first list in this collection. Throws if this collection is empty. [Api set: WordApi 1.3]
			/// </summary>
			/// <returns type="Word.List"></returns>
		}
		ListCollection.prototype.getFirstOrNullObject = function() {
			/// <summary>
			/// Gets the first list in this collection. Returns a null object if this collection is empty. [Api set: WordApi 1.3]
			/// </summary>
			/// <returns type="Word.List"></returns>
		}
		ListCollection.prototype.getItem = function(index) {
			/// <summary>
			/// Gets a list object by its index in the collection. [Api set: WordApi 1.3]
			/// </summary>
			/// <param name="index" >A number that identifies the index location of a list object.</param>
			/// <returns type="Word.List"></returns>
		}

		ListCollection.prototype.track = function() {
			/// <summary>
			/// Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
			/// </summary>
			/// <returns type="Word.ListCollection"/>
		}

		ListCollection.prototype.untrack = function() {
			/// <summary>
			/// Release the memory associated with this object, if has previous been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
			/// </summary>
			/// <returns type="Word.ListCollection"/>
		}

		return ListCollection;
	})(OfficeExtension.ClientObject);
	Word.ListCollection = ListCollection;
})(Word || (Word = {__proto__: null}));

var Word;
(function (Word) {
	var ListItem = (function(_super) {
		__extends(ListItem, _super);
		function ListItem() {
			/// <summary> Represents the paragraph list item format. [Api set: WordApi 1.3] </summary>
			/// <field name="context" type="Word.RequestContext">The request context associated with this object.</field>
			/// <field name="isNull" type="Boolean">Returns a boolean value for whether the corresponding object is null. You must call "context.sync()" before reading the isNull property.</field>
			/// <field name="level" type="Number">Gets or sets the level of the item in the list. [Api set: WordApi 1.3]</field>
			/// <field name="listString" type="String">Gets the list item bullet, number, or picture as a string. Read-only. [Api set: WordApi 1.3]</field>
			/// <field name="siblingIndex" type="Number">Gets the list item order number in relation to its siblings. Read-only. [Api set: WordApi 1.3]</field>
		}

		ListItem.prototype.load = function(option) {
			/// <summary>
			/// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
			/// </summary>
			/// <param name="option" type="string | string[] | OfficeExtension.LoadOption"/>
			/// <returns type="Word.ListItem"/>
		}

		ListItem.prototype.set = function() {
			/// <signature>
			/// <summary>Sets multiple properties on the object at the same time, based on JSON input.</summary>
			/// <param name="properties" type="Word.Interfaces.ListItemUpdateData">Properties described by the Word.Interfaces.ListItemUpdateData interface.</param>
			/// <param name="options" type="string">Options of the form { throwOnReadOnly?: boolean }
			/// <br />
			/// * throwOnReadOnly: Throw an error if the passed-in property list includes read-only properties (default = true).
			/// </param>
			/// </signature>
			/// <signature>
			/// <summary>Sets multiple properties on the object at the same time, based on an existing loaded object.</summary>
			/// <param name="properties" type="ListItem">An existing ListItem object, with properties that have already been loaded and synced.</param>
			/// </signature>
		}
		ListItem.prototype.getAncestor = function(parentOnly) {
			/// <summary>
			/// Gets the list item parent, or the closest ancestor if the parent does not exist. Throws if the list item has no ancestor. [Api set: WordApi 1.3]
			/// </summary>
			/// <param name="parentOnly" type="Boolean" optional="true">Optional. Specifies only the list item&apos;s parent will be returned. The default is false that specifies to get the lowest ancestor.</param>
			/// <returns type="Word.Paragraph"></returns>
		}
		ListItem.prototype.getAncestorOrNullObject = function(parentOnly) {
			/// <summary>
			/// Gets the list item parent, or the closest ancestor if the parent does not exist. Returns a null object if the list item has no ancestor. [Api set: WordApi 1.3]
			/// </summary>
			/// <param name="parentOnly" type="Boolean" optional="true">Optional. Specifies only the list item&apos;s parent will be returned. The default is false that specifies to get the lowest ancestor.</param>
			/// <returns type="Word.Paragraph"></returns>
		}
		ListItem.prototype.getDescendants = function(directChildrenOnly) {
			/// <summary>
			/// Gets all descendant list items of the list item. [Api set: WordApi 1.3]
			/// </summary>
			/// <param name="directChildrenOnly" type="Boolean" optional="true">Optional. Specifies only the list item&apos;s direct children will be returned. The default is false that indicates to get all descendant items.</param>
			/// <returns type="Word.ParagraphCollection"></returns>
		}

		ListItem.prototype.track = function() {
			/// <summary>
			/// Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
			/// </summary>
			/// <returns type="Word.ListItem"/>
		}

		ListItem.prototype.untrack = function() {
			/// <summary>
			/// Release the memory associated with this object, if has previous been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
			/// </summary>
			/// <returns type="Word.ListItem"/>
		}

		return ListItem;
	})(OfficeExtension.ClientObject);
	Word.ListItem = ListItem;
})(Word || (Word = {__proto__: null}));

var Word;
(function (Word) {
	/// <summary> [Api set: WordApi] </summary>
	var ListLevelType = {
		__proto__: null,
		"bullet": "bullet",
		"number": "number",
		"picture": "picture",
	}
	Word.ListLevelType = ListLevelType;
})(Word || (Word = {__proto__: null}));

var Word;
(function (Word) {
	/// <summary> [Api set: WordApi] </summary>
	var ListNumbering = {
		__proto__: null,
		"none": "none",
		"arabic": "arabic",
		"upperRoman": "upperRoman",
		"lowerRoman": "lowerRoman",
		"upperLetter": "upperLetter",
		"lowerLetter": "lowerLetter",
	}
	Word.ListNumbering = ListNumbering;
})(Word || (Word = {__proto__: null}));

var Word;
(function (Word) {
	/// <summary> [Api set: WordApi] </summary>
	var LocationRelation = {
		__proto__: null,
		"unrelated": "unrelated",
		"equal": "equal",
		"containsStart": "containsStart",
		"containsEnd": "containsEnd",
		"contains": "contains",
		"insideStart": "insideStart",
		"insideEnd": "insideEnd",
		"inside": "inside",
		"adjacentBefore": "adjacentBefore",
		"overlapsBefore": "overlapsBefore",
		"before": "before",
		"adjacentAfter": "adjacentAfter",
		"overlapsAfter": "overlapsAfter",
		"after": "after",
	}
	Word.LocationRelation = LocationRelation;
})(Word || (Word = {__proto__: null}));

var Word;
(function (Word) {
	var Paragraph = (function(_super) {
		__extends(Paragraph, _super);
		function Paragraph() {
			/// <summary> Represents a single paragraph in a selection, range, content control, or document body. [Api set: WordApi 1.1] </summary>
			/// <field name="context" type="Word.RequestContext">The request context associated with this object.</field>
			/// <field name="isNull" type="Boolean">Returns a boolean value for whether the corresponding object is null. You must call "context.sync()" before reading the isNull property.</field>
			/// <field name="alignment" type="String">Gets or sets the alignment for a paragraph. The value can be &apos;left&apos;, &apos;centered&apos;, &apos;right&apos;, or &apos;justified&apos;. [Api set: WordApi 1.1]</field>
			/// <field name="contentControls" type="Word.ContentControlCollection">Gets the collection of content control objects in the paragraph. Read-only. [Api set: WordApi 1.1]</field>
			/// <field name="firstLineIndent" type="Number">Gets or sets the value, in points, for a first line or hanging indent. Use a positive value to set a first-line indent, and use a negative value to set a hanging indent. [Api set: WordApi 1.1]</field>
			/// <field name="font" type="Word.Font">Gets the text format of the paragraph. Use this to get and set font name, size, color, and other properties. Read-only. [Api set: WordApi 1.1]</field>
			/// <field name="inlinePictures" type="Word.InlinePictureCollection">Gets the collection of InlinePicture objects in the paragraph. The collection does not include floating images. Read-only. [Api set: WordApi 1.1]</field>
			/// <field name="isLastParagraph" type="Boolean">Indicates the paragraph is the last one inside its parent body. Read-only. [Api set: WordApi 1.3]</field>
			/// <field name="isListItem" type="Boolean">Checks whether the paragraph is a list item. Read-only. [Api set: WordApi 1.3]</field>
			/// <field name="leftIndent" type="Number">Gets or sets the left indent value, in points, for the paragraph. [Api set: WordApi 1.1]</field>
			/// <field name="lineSpacing" type="Number">Gets or sets the line spacing, in points, for the specified paragraph. In the Word UI, this value is divided by 12. [Api set: WordApi 1.1]</field>
			/// <field name="lineUnitAfter" type="Number">Gets or sets the amount of spacing, in grid lines, after the paragraph. [Api set: WordApi 1.1]</field>
			/// <field name="lineUnitBefore" type="Number">Gets or sets the amount of spacing, in grid lines, before the paragraph. [Api set: WordApi 1.1]</field>
			/// <field name="list" type="Word.List">Gets the List to which this paragraph belongs. Throws if the paragraph is not in a list. Read-only. [Api set: WordApi 1.3]</field>
			/// <field name="listItem" type="Word.ListItem">Gets the ListItem for the paragraph. Throws if the paragraph is not part of a list. Read-only. [Api set: WordApi 1.3]</field>
			/// <field name="listItemOrNullObject" type="Word.ListItem">Gets the ListItem for the paragraph. Returns a null object if the paragraph is not part of a list. Read-only. [Api set: WordApi 1.3]</field>
			/// <field name="listOrNullObject" type="Word.List">Gets the List to which this paragraph belongs. Returns a null object if the paragraph is not in a list. Read-only. [Api set: WordApi 1.3]</field>
			/// <field name="outlineLevel" type="Number">Gets or sets the outline level for the paragraph. [Api set: WordApi 1.1]</field>
			/// <field name="parentBody" type="Word.Body">Gets the parent body of the paragraph. Read-only. [Api set: WordApi 1.3]</field>
			/// <field name="parentContentControl" type="Word.ContentControl">Gets the content control that contains the paragraph. Throws if there isn&apos;t a parent content control. Read-only. [Api set: WordApi 1.1]</field>
			/// <field name="parentContentControlOrNullObject" type="Word.ContentControl">Gets the content control that contains the paragraph. Returns a null object if there isn&apos;t a parent content control. Read-only. [Api set: WordApi 1.3]</field>
			/// <field name="parentTable" type="Word.Table">Gets the table that contains the paragraph. Throws if it is not contained in a table. Read-only. [Api set: WordApi 1.3]</field>
			/// <field name="parentTableCell" type="Word.TableCell">Gets the table cell that contains the paragraph. Throws if it is not contained in a table cell. Read-only. [Api set: WordApi 1.3]</field>
			/// <field name="parentTableCellOrNullObject" type="Word.TableCell">Gets the table cell that contains the paragraph. Returns a null object if it is not contained in a table cell. Read-only. [Api set: WordApi 1.3]</field>
			/// <field name="parentTableOrNullObject" type="Word.Table">Gets the table that contains the paragraph. Returns a null object if it is not contained in a table. Read-only. [Api set: WordApi 1.3]</field>
			/// <field name="rightIndent" type="Number">Gets or sets the right indent value, in points, for the paragraph. [Api set: WordApi 1.1]</field>
			/// <field name="spaceAfter" type="Number">Gets or sets the spacing, in points, after the paragraph. [Api set: WordApi 1.1]</field>
			/// <field name="spaceBefore" type="Number">Gets or sets the spacing, in points, before the paragraph. [Api set: WordApi 1.1]</field>
			/// <field name="style" type="String">Gets or sets the style name for the paragraph. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the &quot;styleBuiltIn&quot; property. [Api set: WordApi 1.1]</field>
			/// <field name="styleBuiltIn" type="String">Gets or sets the built-in style name for the paragraph. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the &quot;style&quot; property. [Api set: WordApi 1.3]</field>
			/// <field name="tableNestingLevel" type="Number">Gets the level of the paragraph&apos;s table. It returns 0 if the paragraph is not in a table. Read-only. [Api set: WordApi 1.3]</field>
			/// <field name="text" type="String">Gets the text of the paragraph. Read-only. [Api set: WordApi 1.1]</field>
		}

		Paragraph.prototype.load = function(option) {
			/// <summary>
			/// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
			/// </summary>
			/// <param name="option" type="string | string[] | OfficeExtension.LoadOption"/>
			/// <returns type="Word.Paragraph"/>
		}

		Paragraph.prototype.set = function() {
			/// <signature>
			/// <summary>Sets multiple properties on the object at the same time, based on JSON input.</summary>
			/// <param name="properties" type="Word.Interfaces.ParagraphUpdateData">Properties described by the Word.Interfaces.ParagraphUpdateData interface.</param>
			/// <param name="options" type="string">Options of the form { throwOnReadOnly?: boolean }
			/// <br />
			/// * throwOnReadOnly: Throw an error if the passed-in property list includes read-only properties (default = true).
			/// </param>
			/// </signature>
			/// <signature>
			/// <summary>Sets multiple properties on the object at the same time, based on an existing loaded object.</summary>
			/// <param name="properties" type="Paragraph">An existing Paragraph object, with properties that have already been loaded and synced.</param>
			/// </signature>
		}
		Paragraph.prototype.attachToList = function(listId, level) {
			/// <summary>
			/// Lets the paragraph join an existing list at the specified level. Fails if the paragraph cannot join the list or if the paragraph is already a list item. [Api set: WordApi 1.3]
			/// </summary>
			/// <param name="listId" type="Number">Required. The ID of an existing list.</param>
			/// <param name="level" type="Number">Required. The level in the list.</param>
			/// <returns type="Word.List"></returns>
		}
		Paragraph.prototype.clear = function() {
			/// <summary>
			/// Clears the contents of the paragraph object. The user can perform the undo operation on the cleared content. [Api set: WordApi 1.1]
			/// </summary>
			/// <returns ></returns>
		}
		Paragraph.prototype.delete = function() {
			/// <summary>
			/// Deletes the paragraph and its content from the document. [Api set: WordApi 1.1]
			/// </summary>
			/// <returns ></returns>
		}
		Paragraph.prototype.detachFromList = function() {
			/// <summary>
			/// Moves this paragraph out of its list, if the paragraph is a list item. [Api set: WordApi 1.3]
			/// </summary>
			/// <returns ></returns>
		}
		Paragraph.prototype.getHtml = function() {
			/// <summary>
			/// Gets an HTML representation of the paragraph object. When rendered in a web page or HTML viewer, the formatting will be a close, but not exact, match for of the formatting of the document. This method does not return the exact same HTML for the same document on different platforms (Windows, Mac, Word Online, etc.). If you need exact fidelity, or consistency across platforms, use `Paragraph.getOoxml()` and convert the returned XML to HTML. [Api set: WordApi 1.1]
			/// </summary>
			/// <returns type="OfficeExtension.ClientResult&lt;string&gt;"></returns>
			var result = new OfficeExtension.ClientResult();
			result.__proto__ = null;
			result.value = '';
			return result;
		}
		Paragraph.prototype.getNext = function() {
			/// <summary>
			/// Gets the next paragraph. Throws if the paragraph is the last one. [Api set: WordApi 1.3]
			/// </summary>
			/// <returns type="Word.Paragraph"></returns>
		}
		Paragraph.prototype.getNextOrNullObject = function() {
			/// <summary>
			/// Gets the next paragraph. Returns a null object if the paragraph is the last one. [Api set: WordApi 1.3]
			/// </summary>
			/// <returns type="Word.Paragraph"></returns>
		}
		Paragraph.prototype.getOoxml = function() {
			/// <summary>
			/// Gets the Office Open XML (OOXML) representation of the paragraph object. [Api set: WordApi 1.1]
			/// </summary>
			/// <returns type="OfficeExtension.ClientResult&lt;string&gt;"></returns>
			var result = new OfficeExtension.ClientResult();
			result.__proto__ = null;
			result.value = '';
			return result;
		}
		Paragraph.prototype.getPrevious = function() {
			/// <summary>
			/// Gets the previous paragraph. Throws if the paragraph is the first one. [Api set: WordApi 1.3]
			/// </summary>
			/// <returns type="Word.Paragraph"></returns>
		}
		Paragraph.prototype.getPreviousOrNullObject = function() {
			/// <summary>
			/// Gets the previous paragraph. Returns a null object if the paragraph is the first one. [Api set: WordApi 1.3]
			/// </summary>
			/// <returns type="Word.Paragraph"></returns>
		}
		Paragraph.prototype.getRange = function(rangeLocation) {
			/// <summary>
			/// Gets the whole paragraph, or the starting or ending point of the paragraph, as a range. [Api set: WordApi 1.3]
			/// </summary>
			/// <param name="rangeLocation" type="String" optional="true">Optional. The range location can be &apos;Whole&apos;, &apos;Start&apos;, &apos;End&apos;, &apos;After&apos;, or &apos;Content&apos;.</param>
			/// <returns type="Word.Range"></returns>
		}
		Paragraph.prototype.getTextRanges = function(endingMarks, trimSpacing) {
			/// <summary>
			/// Gets the text ranges in the paragraph by using punctuation marks and/or other ending marks. [Api set: WordApi 1.3]
			/// </summary>
			/// <param name="endingMarks" type="Array" elementType="String">Required. The punctuation marks and/or other ending marks as an array of strings.</param>
			/// <param name="trimSpacing" type="Boolean" optional="true">Optional. Indicates whether to trim spacing characters (spaces, tabs, column breaks, and paragraph end marks) from the start and end of the ranges returned in the range collection. Default is false which indicates that spacing characters at the start and end of the ranges are included in the range collection.</param>
			/// <returns type="Word.RangeCollection"></returns>
		}
		Paragraph.prototype.insertBreak = function(breakType, insertLocation) {
			/// <summary>
			/// Inserts a break at the specified location in the main document. The insertLocation value can be &apos;Before&apos; or &apos;After&apos;. [Api set: WordApi 1.1]
			/// </summary>
			/// <param name="breakType" type="String">Required. The break type to add to the document.</param>
			/// <param name="insertLocation" type="String">Required. The value can be &apos;Before&apos; or &apos;After&apos;.</param>
			/// <returns ></returns>
		}
		Paragraph.prototype.insertContentControl = function() {
			/// <summary>
			/// Wraps the paragraph object with a rich text content control. [Api set: WordApi 1.1]
			/// </summary>
			/// <returns type="Word.ContentControl"></returns>
		}
		Paragraph.prototype.insertFileFromBase64 = function(base64File, insertLocation) {
			/// <summary>
			/// Inserts a document into the paragraph at the specified location. The insertLocation value can be &apos;Replace&apos;, &apos;Start&apos;, or &apos;End&apos;. [Api set: WordApi 1.1]
			/// </summary>
			/// <param name="base64File" type="String">Required. The base64 encoded content of a .docx file.</param>
			/// <param name="insertLocation" type="String">Required. The value can be &apos;Replace&apos;, &apos;Start&apos;, or &apos;End&apos;.</param>
			/// <returns type="Word.Range"></returns>
		}
		Paragraph.prototype.insertHtml = function(html, insertLocation) {
			/// <summary>
			/// Inserts HTML into the paragraph at the specified location. The insertLocation value can be &apos;Replace&apos;, &apos;Start&apos;, or &apos;End&apos;. [Api set: WordApi 1.1]
			/// </summary>
			/// <param name="html" type="String">Required. The HTML to be inserted in the paragraph.</param>
			/// <param name="insertLocation" type="String">Required. The value can be &apos;Replace&apos;, &apos;Start&apos;, or &apos;End&apos;.</param>
			/// <returns type="Word.Range"></returns>
		}
		Paragraph.prototype.insertInlinePictureFromBase64 = function(base64EncodedImage, insertLocation) {
			/// <summary>
			/// Inserts a picture into the paragraph at the specified location. The insertLocation value can be &apos;Replace&apos;, &apos;Start&apos;, or &apos;End&apos;. [Api set: WordApi 1.1]
			/// </summary>
			/// <param name="base64EncodedImage" type="String">Required. The base64 encoded image to be inserted.</param>
			/// <param name="insertLocation" type="String">Required. The value can be &apos;Replace&apos;, &apos;Start&apos;, or &apos;End&apos;.</param>
			/// <returns type="Word.InlinePicture"></returns>
		}
		Paragraph.prototype.insertOoxml = function(ooxml, insertLocation) {
			/// <summary>
			/// Inserts OOXML into the paragraph at the specified location. The insertLocation value can be &apos;Replace&apos;, &apos;Start&apos;, or &apos;End&apos;. [Api set: WordApi 1.1]
			/// </summary>
			/// <param name="ooxml" type="String">Required. The OOXML to be inserted in the paragraph.</param>
			/// <param name="insertLocation" type="String">Required. The value can be &apos;Replace&apos;, &apos;Start&apos;, or &apos;End&apos;.</param>
			/// <returns type="Word.Range"></returns>
		}
		Paragraph.prototype.insertParagraph = function(paragraphText, insertLocation) {
			/// <summary>
			/// Inserts a paragraph at the specified location. The insertLocation value can be &apos;Before&apos; or &apos;After&apos;. [Api set: WordApi 1.1]
			/// </summary>
			/// <param name="paragraphText" type="String">Required. The paragraph text to be inserted.</param>
			/// <param name="insertLocation" type="String">Required. The value can be &apos;Before&apos; or &apos;After&apos;.</param>
			/// <returns type="Word.Paragraph"></returns>
		}
		Paragraph.prototype.insertTable = function(rowCount, columnCount, insertLocation, values) {
			/// <summary>
			/// Inserts a table with the specified number of rows and columns. The insertLocation value can be &apos;Before&apos; or &apos;After&apos;. [Api set: WordApi 1.3]
			/// </summary>
			/// <param name="rowCount" type="Number">Required. The number of rows in the table.</param>
			/// <param name="columnCount" type="Number">Required. The number of columns in the table.</param>
			/// <param name="insertLocation" type="String">Required. The value can be &apos;Before&apos; or &apos;After&apos;.</param>
			/// <param name="values" type="Array" elementType="Array" optional="true">Optional 2D array. Cells are filled if the corresponding strings are specified in the array.</param>
			/// <returns type="Word.Table"></returns>
		}
		Paragraph.prototype.insertText = function(text, insertLocation) {
			/// <summary>
			/// Inserts text into the paragraph at the specified location. The insertLocation value can be &apos;Replace&apos;, &apos;Start&apos;, or &apos;End&apos;. [Api set: WordApi 1.1]
			/// </summary>
			/// <param name="text" type="String">Required. Text to be inserted.</param>
			/// <param name="insertLocation" type="String">Required. The value can be &apos;Replace&apos;, &apos;Start&apos;, or &apos;End&apos;.</param>
			/// <returns type="Word.Range"></returns>
		}
		Paragraph.prototype.search = function(searchText, searchOptions) {
			/// <summary>
			/// Performs a search with the specified SearchOptions on the scope of the paragraph object. The search results are a collection of range objects. [Api set: WordApi 1.1]
			/// </summary>
			/// <param name="searchText" type="String">Required. The search text.</param>
			/// <param name="searchOptions" type="Word.SearchOptions" optional="true">Optional. Options for the search.</param>
			/// <returns type="Word.RangeCollection"></returns>
		}
		Paragraph.prototype.select = function(selectionMode) {
			/// <summary>
			/// Selects and navigates the Word UI to the paragraph. [Api set: WordApi 1.1]
			/// </summary>
			/// <param name="selectionMode" type="String" optional="true">Optional. The selection mode can be &apos;Select&apos;, &apos;Start&apos;, or &apos;End&apos;. &apos;Select&apos; is the default.</param>
			/// <returns ></returns>
		}
		Paragraph.prototype.split = function(delimiters, trimDelimiters, trimSpacing) {
			/// <summary>
			/// Splits the paragraph into child ranges by using delimiters. [Api set: WordApi 1.3]
			/// </summary>
			/// <param name="delimiters" type="Array" elementType="String">Required. The delimiters as an array of strings.</param>
			/// <param name="trimDelimiters" type="Boolean" optional="true">Optional. Indicates whether to trim delimiters from the ranges in the range collection. Default is false which indicates that the delimiters are included in the ranges returned in the range collection.</param>
			/// <param name="trimSpacing" type="Boolean" optional="true">Optional. Indicates whether to trim spacing characters (spaces, tabs, column breaks, and paragraph end marks) from the start and end of the ranges returned in the range collection. Default is false which indicates that spacing characters at the start and end of the ranges are included in the range collection.</param>
			/// <returns type="Word.RangeCollection"></returns>
		}
		Paragraph.prototype.startNewList = function() {
			/// <summary>
			/// Starts a new list with this paragraph. Fails if the paragraph is already a list item. [Api set: WordApi 1.3]
			/// </summary>
			/// <returns type="Word.List"></returns>
		}

		Paragraph.prototype.track = function() {
			/// <summary>
			/// Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
			/// </summary>
			/// <returns type="Word.Paragraph"/>
		}

		Paragraph.prototype.untrack = function() {
			/// <summary>
			/// Release the memory associated with this object, if has previous been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
			/// </summary>
			/// <returns type="Word.Paragraph"/>
		}

		return Paragraph;
	})(OfficeExtension.ClientObject);
	Word.Paragraph = Paragraph;
})(Word || (Word = {__proto__: null}));

var Word;
(function (Word) {
	var ParagraphCollection = (function(_super) {
		__extends(ParagraphCollection, _super);
		function ParagraphCollection() {
			/// <summary> Contains a collection of {@link Word.Paragraph} objects. [Api set: WordApi 1.1] </summary>
			/// <field name="context" type="Word.RequestContext">The request context associated with this object.</field>
			/// <field name="isNull" type="Boolean">Returns a boolean value for whether the corresponding object is null. You must call "context.sync()" before reading the isNull property.</field>
			/// <field name="items" type="Array" elementType="Word.Paragraph">Gets the loaded child items in this collection.</field>
		}

		ParagraphCollection.prototype.load = function(option) {
			/// <summary>
			/// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
			/// </summary>
			/// <param name="option" type="string | string[] | OfficeExtension.LoadOption"/>
			/// <returns type="Word.ParagraphCollection"/>
		}
		ParagraphCollection.prototype.getFirst = function() {
			/// <summary>
			/// Gets the first paragraph in this collection. Throws if the collection is empty. [Api set: WordApi 1.3]
			/// </summary>
			/// <returns type="Word.Paragraph"></returns>
		}
		ParagraphCollection.prototype.getFirstOrNullObject = function() {
			/// <summary>
			/// Gets the first paragraph in this collection. Returns a null object if the collection is empty. [Api set: WordApi 1.3]
			/// </summary>
			/// <returns type="Word.Paragraph"></returns>
		}
		ParagraphCollection.prototype.getLast = function() {
			/// <summary>
			/// Gets the last paragraph in this collection. Throws if the collection is empty. [Api set: WordApi 1.3]
			/// </summary>
			/// <returns type="Word.Paragraph"></returns>
		}
		ParagraphCollection.prototype.getLastOrNullObject = function() {
			/// <summary>
			/// Gets the last paragraph in this collection. Returns a null object if the collection is empty. [Api set: WordApi 1.3]
			/// </summary>
			/// <returns type="Word.Paragraph"></returns>
		}

		ParagraphCollection.prototype.track = function() {
			/// <summary>
			/// Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
			/// </summary>
			/// <returns type="Word.ParagraphCollection"/>
		}

		ParagraphCollection.prototype.untrack = function() {
			/// <summary>
			/// Release the memory associated with this object, if has previous been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
			/// </summary>
			/// <returns type="Word.ParagraphCollection"/>
		}

		return ParagraphCollection;
	})(OfficeExtension.ClientObject);
	Word.ParagraphCollection = ParagraphCollection;
})(Word || (Word = {__proto__: null}));

var Word;
(function (Word) {
	var Range = (function(_super) {
		__extends(Range, _super);
		function Range() {
			/// <summary> Represents a contiguous area in a document. [Api set: WordApi 1.1] </summary>
			/// <field name="context" type="Word.RequestContext">The request context associated with this object.</field>
			/// <field name="isNull" type="Boolean">Returns a boolean value for whether the corresponding object is null. You must call "context.sync()" before reading the isNull property.</field>
			/// <field name="contentControls" type="Word.ContentControlCollection">Gets the collection of content control objects in the range. Read-only. [Api set: WordApi 1.1]</field>
			/// <field name="font" type="Word.Font">Gets the text format of the range. Use this to get and set font name, size, color, and other properties. Read-only. [Api set: WordApi 1.1]</field>
			/// <field name="hyperlink" type="String">Gets the first hyperlink in the range, or sets a hyperlink on the range. All hyperlinks in the range are deleted when you set a new hyperlink on the range. Use a &apos;#&apos; to separate the address part from the optional location part. [Api set: WordApi 1.3]</field>
			/// <field name="inlinePictures" type="Word.InlinePictureCollection">Gets the collection of inline picture objects in the range. Read-only. [Api set: WordApi 1.2]</field>
			/// <field name="isEmpty" type="Boolean">Checks whether the range length is zero. Read-only. [Api set: WordApi 1.3]</field>
			/// <field name="lists" type="Word.ListCollection">Gets the collection of list objects in the range. Read-only. [Api set: WordApi 1.3]</field>
			/// <field name="paragraphs" type="Word.ParagraphCollection">Gets the collection of paragraph objects in the range. Read-only. [Api set: WordApi 1.1]</field>
			/// <field name="parentBody" type="Word.Body">Gets the parent body of the range. Read-only. [Api set: WordApi 1.3]</field>
			/// <field name="parentContentControl" type="Word.ContentControl">Gets the content control that contains the range. Throws if there isn&apos;t a parent content control. Read-only. [Api set: WordApi 1.1]</field>
			/// <field name="parentContentControlOrNullObject" type="Word.ContentControl">Gets the content control that contains the range. Returns a null object if there isn&apos;t a parent content control. Read-only. [Api set: WordApi 1.3]</field>
			/// <field name="parentTable" type="Word.Table">Gets the table that contains the range. Throws if it is not contained in a table. Read-only. [Api set: WordApi 1.3]</field>
			/// <field name="parentTableCell" type="Word.TableCell">Gets the table cell that contains the range. Throws if it is not contained in a table cell. Read-only. [Api set: WordApi 1.3]</field>
			/// <field name="parentTableCellOrNullObject" type="Word.TableCell">Gets the table cell that contains the range. Returns a null object if it is not contained in a table cell. Read-only. [Api set: WordApi 1.3]</field>
			/// <field name="parentTableOrNullObject" type="Word.Table">Gets the table that contains the range. Returns a null object if it is not contained in a table. Read-only. [Api set: WordApi 1.3]</field>
			/// <field name="style" type="String">Gets or sets the style name for the range. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the &quot;styleBuiltIn&quot; property. [Api set: WordApi 1.1]</field>
			/// <field name="styleBuiltIn" type="String">Gets or sets the built-in style name for the range. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the &quot;style&quot; property. [Api set: WordApi 1.3]</field>
			/// <field name="tables" type="Word.TableCollection">Gets the collection of table objects in the range. Read-only. [Api set: WordApi 1.3]</field>
			/// <field name="text" type="String">Gets the text of the range. Read-only. [Api set: WordApi 1.1]</field>
		}

		Range.prototype.load = function(option) {
			/// <summary>
			/// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
			/// </summary>
			/// <param name="option" type="string | string[] | OfficeExtension.LoadOption"/>
			/// <returns type="Word.Range"/>
		}

		Range.prototype.set = function() {
			/// <signature>
			/// <summary>Sets multiple properties on the object at the same time, based on JSON input.</summary>
			/// <param name="properties" type="Word.Interfaces.RangeUpdateData">Properties described by the Word.Interfaces.RangeUpdateData interface.</param>
			/// <param name="options" type="string">Options of the form { throwOnReadOnly?: boolean }
			/// <br />
			/// * throwOnReadOnly: Throw an error if the passed-in property list includes read-only properties (default = true).
			/// </param>
			/// </signature>
			/// <signature>
			/// <summary>Sets multiple properties on the object at the same time, based on an existing loaded object.</summary>
			/// <param name="properties" type="Range">An existing Range object, with properties that have already been loaded and synced.</param>
			/// </signature>
		}
		Range.prototype.clear = function() {
			/// <summary>
			/// Clears the contents of the range object. The user can perform the undo operation on the cleared content. [Api set: WordApi 1.1]
			/// </summary>
			/// <returns ></returns>
		}
		Range.prototype.compareLocationWith = function(range) {
			/// <summary>
			/// Compares this range&apos;s location with another range&apos;s location. [Api set: WordApi 1.3]
			/// </summary>
			/// <param name="range" type="Word.Range">Required. The range to compare with this range.</param>
			/// <returns type="OfficeExtension.ClientResult&lt;Word.LocationRelation&gt;"></returns>
			var result = new OfficeExtension.ClientResult();
			result.__proto__ = null;
			result.value = '';
			return result;
		}
		Range.prototype.delete = function() {
			/// <summary>
			/// Deletes the range and its content from the document. [Api set: WordApi 1.1]
			/// </summary>
			/// <returns ></returns>
		}
		Range.prototype.expandTo = function(range) {
			/// <summary>
			/// Returns a new range that extends from this range in either direction to cover another range. This range is not changed. Throws if the two ranges do not have a union. [Api set: WordApi 1.3]
			/// </summary>
			/// <param name="range" type="Word.Range">Required. Another range.</param>
			/// <returns type="Word.Range"></returns>
		}
		Range.prototype.expandToOrNullObject = function(range) {
			/// <summary>
			/// Returns a new range that extends from this range in either direction to cover another range. This range is not changed. Returns a null object if the two ranges do not have a union. [Api set: WordApi 1.3]
			/// </summary>
			/// <param name="range" type="Word.Range">Required. Another range.</param>
			/// <returns type="Word.Range"></returns>
		}
		Range.prototype.getBookmarks = function(includeHidden, includeAdjacent) {
			/// <summary>
			/// Gets the names all bookmarks in or overlapping the range. A bookmark is hidden if its name starts with the underscore character. [Api set: WordApi 1.4]
			/// </summary>
			/// <param name="includeHidden" type="Boolean" optional="true">Optional. Indicates whether to include hidden bookmarks. Default is false which indicates that the hidden bookmarks are excluded.</param>
			/// <param name="includeAdjacent" type="Boolean" optional="true">Optional. Indicates whether to include bookmarks that are adjacent to the range. Default is false which indicates that the adjacent bookmarks are excluded.</param>
			/// <returns type="OfficeExtension.ClientResult&lt;string[]&gt;"></returns>
			var result = new OfficeExtension.ClientResult();
			result.__proto__ = null;
			result.value = [];
			return result;
		}
		Range.prototype.getHtml = function() {
			/// <summary>
			/// Gets an HTML representation of the range object. When rendered in a web page or HTML viewer, the formatting will be a close, but not exact, match for of the formatting of the document. This method does not return the exact same HTML for the same document on different platforms (Windows, Mac, Word Online, etc.). If you need exact fidelity, or consistency across platforms, use `Range.getOoxml()` and convert the returned XML to HTML. [Api set: WordApi 1.1]
			/// </summary>
			/// <returns type="OfficeExtension.ClientResult&lt;string&gt;"></returns>
			var result = new OfficeExtension.ClientResult();
			result.__proto__ = null;
			result.value = '';
			return result;
		}
		Range.prototype.getHyperlinkRanges = function() {
			/// <summary>
			/// Gets hyperlink child ranges within the range. [Api set: WordApi 1.3]
			/// </summary>
			/// <returns type="Word.RangeCollection"></returns>
		}
		Range.prototype.getNextTextRange = function(endingMarks, trimSpacing) {
			/// <summary>
			/// Gets the next text range by using punctuation marks and/or other ending marks. Throws if this text range is the last one. [Api set: WordApi 1.3]
			/// </summary>
			/// <param name="endingMarks" type="Array" elementType="String">Required. The punctuation marks and/or other ending marks as an array of strings.</param>
			/// <param name="trimSpacing" type="Boolean" optional="true">Optional. Indicates whether to trim spacing characters (spaces, tabs, column breaks, and paragraph end marks) from the start and end of the returned range. Default is false which indicates that spacing characters at the start and end of the range are included.</param>
			/// <returns type="Word.Range"></returns>
		}
		Range.prototype.getNextTextRangeOrNullObject = function(endingMarks, trimSpacing) {
			/// <summary>
			/// Gets the next text range by using punctuation marks and/or other ending marks. Returns a null object if this text range is the last one. [Api set: WordApi 1.3]
			/// </summary>
			/// <param name="endingMarks" type="Array" elementType="String">Required. The punctuation marks and/or other ending marks as an array of strings.</param>
			/// <param name="trimSpacing" type="Boolean" optional="true">Optional. Indicates whether to trim spacing characters (spaces, tabs, column breaks, and paragraph end marks) from the start and end of the returned range. Default is false which indicates that spacing characters at the start and end of the range are included.</param>
			/// <returns type="Word.Range"></returns>
		}
		Range.prototype.getOoxml = function() {
			/// <summary>
			/// Gets the OOXML representation of the range object. [Api set: WordApi 1.1]
			/// </summary>
			/// <returns type="OfficeExtension.ClientResult&lt;string&gt;"></returns>
			var result = new OfficeExtension.ClientResult();
			result.__proto__ = null;
			result.value = '';
			return result;
		}
		Range.prototype.getRange = function(rangeLocation) {
			/// <summary>
			/// Clones the range, or gets the starting or ending point of the range as a new range. [Api set: WordApi 1.3]
			/// </summary>
			/// <param name="rangeLocation" type="String" optional="true">Optional. The range location can be &apos;Whole&apos;, &apos;Start&apos;, &apos;End&apos;, &apos;After&apos;, or &apos;Content&apos;.</param>
			/// <returns type="Word.Range"></returns>
		}
		Range.prototype.getTextRanges = function(endingMarks, trimSpacing) {
			/// <summary>
			/// Gets the text child ranges in the range by using punctuation marks and/or other ending marks. [Api set: WordApi 1.3]
			/// </summary>
			/// <param name="endingMarks" type="Array" elementType="String">Required. The punctuation marks and/or other ending marks as an array of strings.</param>
			/// <param name="trimSpacing" type="Boolean" optional="true">Optional. Indicates whether to trim spacing characters (spaces, tabs, column breaks, and paragraph end marks) from the start and end of the ranges returned in the range collection. Default is false which indicates that spacing characters at the start and end of the ranges are included in the range collection.</param>
			/// <returns type="Word.RangeCollection"></returns>
		}
		Range.prototype.insertBookmark = function(name) {
			/// <summary>
			/// Inserts a bookmark on the range. If a bookmark of the same name exists somewhere, it is deleted first. [Api set: WordApi 1.4]
			/// </summary>
			/// <param name="name" type="String">Required. The bookmark name, which is case-insensitive. If the name starts with an underscore character, the bookmark is an hidden one.</param>
			/// <returns ></returns>
		}
		Range.prototype.insertBreak = function(breakType, insertLocation) {
			/// <summary>
			/// Inserts a break at the specified location in the main document. The insertLocation value can be &apos;Before&apos; or &apos;After&apos;. [Api set: WordApi 1.1]
			/// </summary>
			/// <param name="breakType" type="String">Required. The break type to add.</param>
			/// <param name="insertLocation" type="String">Required. The value can be &apos;Before&apos; or &apos;After&apos;.</param>
			/// <returns ></returns>
		}
		Range.prototype.insertContentControl = function() {
			/// <summary>
			/// Wraps the range object with a rich text content control. [Api set: WordApi 1.1]
			/// </summary>
			/// <returns type="Word.ContentControl"></returns>
		}
		Range.prototype.insertFileFromBase64 = function(base64File, insertLocation) {
			/// <summary>
			/// Inserts a document at the specified location. The insertLocation value can be &apos;Replace&apos;, &apos;Start&apos;, &apos;End&apos;, &apos;Before&apos;, or &apos;After&apos;. [Api set: WordApi 1.1]
			/// </summary>
			/// <param name="base64File" type="String">Required. The base64 encoded content of a .docx file.</param>
			/// <param name="insertLocation" type="String">Required. The value can be &apos;Replace&apos;, &apos;Start&apos;, &apos;End&apos;, &apos;Before&apos;, or &apos;After&apos;.</param>
			/// <returns type="Word.Range"></returns>
		}
		Range.prototype.insertHtml = function(html, insertLocation) {
			/// <summary>
			/// Inserts HTML at the specified location. The insertLocation value can be &apos;Replace&apos;, &apos;Start&apos;, &apos;End&apos;, &apos;Before&apos;, or &apos;After&apos;. [Api set: WordApi 1.1]
			/// </summary>
			/// <param name="html" type="String">Required. The HTML to be inserted.</param>
			/// <param name="insertLocation" type="String">Required. The value can be &apos;Replace&apos;, &apos;Start&apos;, &apos;End&apos;, &apos;Before&apos;, or &apos;After&apos;.</param>
			/// <returns type="Word.Range"></returns>
		}
		Range.prototype.insertInlinePictureFromBase64 = function(base64EncodedImage, insertLocation) {
			/// <summary>
			/// Inserts a picture at the specified location. The insertLocation value can be &apos;Replace&apos;, &apos;Start&apos;, &apos;End&apos;, &apos;Before&apos;, or &apos;After&apos;. [Api set: WordApi 1.2]
			/// </summary>
			/// <param name="base64EncodedImage" type="String">Required. The base64 encoded image to be inserted.</param>
			/// <param name="insertLocation" type="String">Required. The value can be &apos;Replace&apos;, &apos;Start&apos;, &apos;End&apos;, &apos;Before&apos;, or &apos;After&apos;.</param>
			/// <returns type="Word.InlinePicture"></returns>
		}
		Range.prototype.insertOoxml = function(ooxml, insertLocation) {
			/// <summary>
			/// Inserts OOXML at the specified location.  The insertLocation value can be &apos;Replace&apos;, &apos;Start&apos;, &apos;End&apos;, &apos;Before&apos;, or &apos;After&apos;. [Api set: WordApi 1.1]
			/// </summary>
			/// <param name="ooxml" type="String">Required. The OOXML to be inserted.</param>
			/// <param name="insertLocation" type="String">Required. The value can be &apos;Replace&apos;, &apos;Start&apos;, &apos;End&apos;, &apos;Before&apos;, or &apos;After&apos;.</param>
			/// <returns type="Word.Range"></returns>
		}
		Range.prototype.insertParagraph = function(paragraphText, insertLocation) {
			/// <summary>
			/// Inserts a paragraph at the specified location. The insertLocation value can be &apos;Before&apos; or &apos;After&apos;. [Api set: WordApi 1.1]
			/// </summary>
			/// <param name="paragraphText" type="String">Required. The paragraph text to be inserted.</param>
			/// <param name="insertLocation" type="String">Required. The value can be &apos;Before&apos; or &apos;After&apos;.</param>
			/// <returns type="Word.Paragraph"></returns>
		}
		Range.prototype.insertTable = function(rowCount, columnCount, insertLocation, values) {
			/// <summary>
			/// Inserts a table with the specified number of rows and columns. The insertLocation value can be &apos;Before&apos; or &apos;After&apos;. [Api set: WordApi 1.3]
			/// </summary>
			/// <param name="rowCount" type="Number">Required. The number of rows in the table.</param>
			/// <param name="columnCount" type="Number">Required. The number of columns in the table.</param>
			/// <param name="insertLocation" type="String">Required. The value can be &apos;Before&apos; or &apos;After&apos;.</param>
			/// <param name="values" type="Array" elementType="Array" optional="true">Optional 2D array. Cells are filled if the corresponding strings are specified in the array.</param>
			/// <returns type="Word.Table"></returns>
		}
		Range.prototype.insertText = function(text, insertLocation) {
			/// <summary>
			/// Inserts text at the specified location. The insertLocation value can be &apos;Replace&apos;, &apos;Start&apos;, &apos;End&apos;, &apos;Before&apos;, or &apos;After&apos;. [Api set: WordApi 1.1]
			/// </summary>
			/// <param name="text" type="String">Required. Text to be inserted.</param>
			/// <param name="insertLocation" type="String">Required. The value can be &apos;Replace&apos;, &apos;Start&apos;, &apos;End&apos;, &apos;Before&apos;, or &apos;After&apos;.</param>
			/// <returns type="Word.Range"></returns>
		}
		Range.prototype.intersectWith = function(range) {
			/// <summary>
			/// Returns a new range as the intersection of this range with another range. This range is not changed. Throws if the two ranges are not overlapped or adjacent. [Api set: WordApi 1.3]
			/// </summary>
			/// <param name="range" type="Word.Range">Required. Another range.</param>
			/// <returns type="Word.Range"></returns>
		}
		Range.prototype.intersectWithOrNullObject = function(range) {
			/// <summary>
			/// Returns a new range as the intersection of this range with another range. This range is not changed. Returns a null object if the two ranges are not overlapped or adjacent. [Api set: WordApi 1.3]
			/// </summary>
			/// <param name="range" type="Word.Range">Required. Another range.</param>
			/// <returns type="Word.Range"></returns>
		}
		Range.prototype.search = function(searchText, searchOptions) {
			/// <summary>
			/// Performs a search with the specified SearchOptions on the scope of the range object. The search results are a collection of range objects. [Api set: WordApi 1.1]
			/// </summary>
			/// <param name="searchText" type="String">Required. The search text.</param>
			/// <param name="searchOptions" type="Word.SearchOptions" optional="true">Optional. Options for the search.</param>
			/// <returns type="Word.RangeCollection"></returns>
		}
		Range.prototype.select = function(selectionMode) {
			/// <summary>
			/// Selects and navigates the Word UI to the range. [Api set: WordApi 1.1]
			/// </summary>
			/// <param name="selectionMode" type="String" optional="true">Optional. The selection mode can be &apos;Select&apos;, &apos;Start&apos;, or &apos;End&apos;. &apos;Select&apos; is the default.</param>
			/// <returns ></returns>
		}
		Range.prototype.split = function(delimiters, multiParagraphs, trimDelimiters, trimSpacing) {
			/// <summary>
			/// Splits the range into child ranges by using delimiters. [Api set: WordApi 1.3]
			/// </summary>
			/// <param name="delimiters" type="Array" elementType="String">Required. The delimiters as an array of strings.</param>
			/// <param name="multiParagraphs" type="Boolean" optional="true">Optional. Indicates whether a returned child range can cover multiple paragraphs. Default is false which indicates that the paragraph boundaries are also used as delimiters.</param>
			/// <param name="trimDelimiters" type="Boolean" optional="true">Optional. Indicates whether to trim delimiters from the ranges in the range collection. Default is false which indicates that the delimiters are included in the ranges returned in the range collection.</param>
			/// <param name="trimSpacing" type="Boolean" optional="true">Optional. Indicates whether to trim spacing characters (spaces, tabs, column breaks, and paragraph end marks) from the start and end of the ranges returned in the range collection. Default is false which indicates that spacing characters at the start and end of the ranges are included in the range collection.</param>
			/// <returns type="Word.RangeCollection"></returns>
		}

		Range.prototype.track = function() {
			/// <summary>
			/// Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
			/// </summary>
			/// <returns type="Word.Range"/>
		}

		Range.prototype.untrack = function() {
			/// <summary>
			/// Release the memory associated with this object, if has previous been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
			/// </summary>
			/// <returns type="Word.Range"/>
		}

		return Range;
	})(OfficeExtension.ClientObject);
	Word.Range = Range;
})(Word || (Word = {__proto__: null}));

var Word;
(function (Word) {
	var RangeCollection = (function(_super) {
		__extends(RangeCollection, _super);
		function RangeCollection() {
			/// <summary> Contains a collection of {@link Word.Range} objects. [Api set: WordApi 1.1] </summary>
			/// <field name="context" type="Word.RequestContext">The request context associated with this object.</field>
			/// <field name="isNull" type="Boolean">Returns a boolean value for whether the corresponding object is null. You must call "context.sync()" before reading the isNull property.</field>
			/// <field name="items" type="Array" elementType="Word.Range">Gets the loaded child items in this collection.</field>
		}

		RangeCollection.prototype.load = function(option) {
			/// <summary>
			/// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
			/// </summary>
			/// <param name="option" type="string | string[] | OfficeExtension.LoadOption"/>
			/// <returns type="Word.RangeCollection"/>
		}
		RangeCollection.prototype.getFirst = function() {
			/// <summary>
			/// Gets the first range in this collection. Throws if this collection is empty. [Api set: WordApi 1.3]
			/// </summary>
			/// <returns type="Word.Range"></returns>
		}
		RangeCollection.prototype.getFirstOrNullObject = function() {
			/// <summary>
			/// Gets the first range in this collection. Returns a null object if this collection is empty. [Api set: WordApi 1.3]
			/// </summary>
			/// <returns type="Word.Range"></returns>
		}

		RangeCollection.prototype.track = function() {
			/// <summary>
			/// Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
			/// </summary>
			/// <returns type="Word.RangeCollection"/>
		}

		RangeCollection.prototype.untrack = function() {
			/// <summary>
			/// Release the memory associated with this object, if has previous been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
			/// </summary>
			/// <returns type="Word.RangeCollection"/>
		}

		return RangeCollection;
	})(OfficeExtension.ClientObject);
	Word.RangeCollection = RangeCollection;
})(Word || (Word = {__proto__: null}));

var Word;
(function (Word) {
	/// <summary> [Api set: WordApi] </summary>
	var RangeLocation = {
		__proto__: null,
		"whole": "whole",
		"start": "start",
		"end": "end",
		"before": "before",
		"after": "after",
		"content": "content",
	}
	Word.RangeLocation = RangeLocation;
})(Word || (Word = {__proto__: null}));

var Word;
(function (Word) {
	var SearchOptions = (function(_super) {
		__extends(SearchOptions, _super);
		function SearchOptions() {
			/// <summary> Specifies the options to be included in a search operation. [Api set: WordApi 1.1] </summary>
			/// <field name="context" type="Word.RequestContext">The request context associated with this object.</field>
			/// <field name="isNull" type="Boolean">Returns a boolean value for whether the corresponding object is null. You must call "context.sync()" before reading the isNull property.</field>
			/// <field name="ignorePunct" type="Boolean">Gets or sets a value that indicates whether to ignore all punctuation characters between words. Corresponds to the Ignore punctuation check box in the Find and Replace dialog box. [Api set: WordApi 1.1]</field>
			/// <field name="ignoreSpace" type="Boolean">Gets or sets a value that indicates whether to ignore all whitespace between words. Corresponds to the Ignore whitespace characters check box in the Find and Replace dialog box. [Api set: WordApi 1.1]</field>
			/// <field name="matchCase" type="Boolean">Gets or sets a value that indicates whether to perform a case sensitive search. Corresponds to the Match case check box in the Find and Replace dialog box. [Api set: WordApi 1.1]</field>
			/// <field name="matchPrefix" type="Boolean">Gets or sets a value that indicates whether to match words that begin with the search string. Corresponds to the Match prefix check box in the Find and Replace dialog box. [Api set: WordApi 1.1]</field>
			/// <field name="matchSuffix" type="Boolean">Gets or sets a value that indicates whether to match words that end with the search string. Corresponds to the Match suffix check box in the Find and Replace dialog box. [Api set: WordApi 1.1]</field>
			/// <field name="matchWholeWord" type="Boolean">Gets or sets a value that indicates whether to find operation only entire words, not text that is part of a larger word. Corresponds to the Find whole words only check box in the Find and Replace dialog box. [Api set: WordApi 1.1]</field>
			/// <field name="matchWildcards" type="Boolean">Gets or sets a value that indicates whether the search will be performed using special search operators. Corresponds to the Use wildcards check box in the Find and Replace dialog box. [Api set: WordApi 1.1]</field>
		}

		SearchOptions.prototype.load = function(option) {
			/// <summary>
			/// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
			/// </summary>
			/// <param name="option" type="string | string[] | OfficeExtension.LoadOption"/>
			/// <returns type="Word.SearchOptions"/>
		}

		SearchOptions.prototype.set = function() {
			/// <signature>
			/// <summary>Sets multiple properties on the object at the same time, based on JSON input.</summary>
			/// <param name="properties" type="Word.Interfaces.SearchOptionsUpdateData">Properties described by the Word.Interfaces.SearchOptionsUpdateData interface.</param>
			/// <param name="options" type="string">Options of the form { throwOnReadOnly?: boolean }
			/// <br />
			/// * throwOnReadOnly: Throw an error if the passed-in property list includes read-only properties (default = true).
			/// </param>
			/// </signature>
			/// <signature>
			/// <summary>Sets multiple properties on the object at the same time, based on an existing loaded object.</summary>
			/// <param name="properties" type="SearchOptions">An existing SearchOptions object, with properties that have already been loaded and synced.</param>
			/// </signature>
		}

		return SearchOptions;
	})(OfficeExtension.ClientObject);
	Word.SearchOptions = SearchOptions;
})(Word || (Word = {__proto__: null}));

var Word;
(function (Word) {
	var Section = (function(_super) {
		__extends(Section, _super);
		function Section() {
			/// <summary> Represents a section in a Word document. [Api set: WordApi 1.1] </summary>
			/// <field name="context" type="Word.RequestContext">The request context associated with this object.</field>
			/// <field name="isNull" type="Boolean">Returns a boolean value for whether the corresponding object is null. You must call "context.sync()" before reading the isNull property.</field>
			/// <field name="body" type="Word.Body">Gets the body object of the section. This does not include the header/footer and other section metadata. Read-only. [Api set: WordApi 1.1]</field>
		}

		Section.prototype.load = function(option) {
			/// <summary>
			/// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
			/// </summary>
			/// <param name="option" type="string | string[] | OfficeExtension.LoadOption"/>
			/// <returns type="Word.Section"/>
		}

		Section.prototype.set = function() {
			/// <signature>
			/// <summary>Sets multiple properties on the object at the same time, based on JSON input.</summary>
			/// <param name="properties" type="Word.Interfaces.SectionUpdateData">Properties described by the Word.Interfaces.SectionUpdateData interface.</param>
			/// <param name="options" type="string">Options of the form { throwOnReadOnly?: boolean }
			/// <br />
			/// * throwOnReadOnly: Throw an error if the passed-in property list includes read-only properties (default = true).
			/// </param>
			/// </signature>
			/// <signature>
			/// <summary>Sets multiple properties on the object at the same time, based on an existing loaded object.</summary>
			/// <param name="properties" type="Section">An existing Section object, with properties that have already been loaded and synced.</param>
			/// </signature>
		}
		Section.prototype.getFooter = function(type) {
			/// <summary>
			/// Gets one of the section&apos;s footers. [Api set: WordApi 1.1]
			/// </summary>
			/// <param name="type" type="String">Required. The type of footer to return. This value can be: &apos;Primary&apos;, &apos;FirstPage&apos;, or &apos;EvenPages&apos;.</param>
			/// <returns type="Word.Body"></returns>
		}
		Section.prototype.getHeader = function(type) {
			/// <summary>
			/// Gets one of the section&apos;s headers. [Api set: WordApi 1.1]
			/// </summary>
			/// <param name="type" type="String">Required. The type of header to return. This value can be: &apos;Primary&apos;, &apos;FirstPage&apos;, or &apos;EvenPages&apos;.</param>
			/// <returns type="Word.Body"></returns>
		}
		Section.prototype.getNext = function() {
			/// <summary>
			/// Gets the next section. Throws if this section is the last one. [Api set: WordApi 1.3]
			/// </summary>
			/// <returns type="Word.Section"></returns>
		}
		Section.prototype.getNextOrNullObject = function() {
			/// <summary>
			/// Gets the next section. Returns a null object if this section is the last one. [Api set: WordApi 1.3]
			/// </summary>
			/// <returns type="Word.Section"></returns>
		}

		Section.prototype.track = function() {
			/// <summary>
			/// Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
			/// </summary>
			/// <returns type="Word.Section"/>
		}

		Section.prototype.untrack = function() {
			/// <summary>
			/// Release the memory associated with this object, if has previous been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
			/// </summary>
			/// <returns type="Word.Section"/>
		}

		return Section;
	})(OfficeExtension.ClientObject);
	Word.Section = Section;
})(Word || (Word = {__proto__: null}));

var Word;
(function (Word) {
	var SectionCollection = (function(_super) {
		__extends(SectionCollection, _super);
		function SectionCollection() {
			/// <summary> Contains the collection of the document&apos;s {@link Word.Section} objects. [Api set: WordApi 1.1] </summary>
			/// <field name="context" type="Word.RequestContext">The request context associated with this object.</field>
			/// <field name="isNull" type="Boolean">Returns a boolean value for whether the corresponding object is null. You must call "context.sync()" before reading the isNull property.</field>
			/// <field name="items" type="Array" elementType="Word.Section">Gets the loaded child items in this collection.</field>
		}

		SectionCollection.prototype.load = function(option) {
			/// <summary>
			/// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
			/// </summary>
			/// <param name="option" type="string | string[] | OfficeExtension.LoadOption"/>
			/// <returns type="Word.SectionCollection"/>
		}
		SectionCollection.prototype.getFirst = function() {
			/// <summary>
			/// Gets the first section in this collection. Throws if this collection is empty. [Api set: WordApi 1.3]
			/// </summary>
			/// <returns type="Word.Section"></returns>
		}
		SectionCollection.prototype.getFirstOrNullObject = function() {
			/// <summary>
			/// Gets the first section in this collection. Returns a null object if this collection is empty. [Api set: WordApi 1.3]
			/// </summary>
			/// <returns type="Word.Section"></returns>
		}

		SectionCollection.prototype.track = function() {
			/// <summary>
			/// Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
			/// </summary>
			/// <returns type="Word.SectionCollection"/>
		}

		SectionCollection.prototype.untrack = function() {
			/// <summary>
			/// Release the memory associated with this object, if has previous been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
			/// </summary>
			/// <returns type="Word.SectionCollection"/>
		}

		return SectionCollection;
	})(OfficeExtension.ClientObject);
	Word.SectionCollection = SectionCollection;
})(Word || (Word = {__proto__: null}));

var Word;
(function (Word) {
	/// <summary> [Api set: WordApi] </summary>
	var SelectionMode = {
		__proto__: null,
		"select": "select",
		"start": "start",
		"end": "end",
	}
	Word.SelectionMode = SelectionMode;
})(Word || (Word = {__proto__: null}));

var Word;
(function (Word) {
	var Setting = (function(_super) {
		__extends(Setting, _super);
		function Setting() {
			/// <summary> Represents a setting of the add-in. [Api set: WordApi 1.4] </summary>
			/// <field name="context" type="Word.RequestContext">The request context associated with this object.</field>
			/// <field name="isNull" type="Boolean">Returns a boolean value for whether the corresponding object is null. You must call "context.sync()" before reading the isNull property.</field>
			/// <field name="key" type="String">Gets the key of the setting. Read only. [Api set: WordApi 1.4]</field>
			/// <field name="value" >Gets or sets the value of the setting. [Api set: WordApi 1.4]</field>
		}

		Setting.prototype.load = function(option) {
			/// <summary>
			/// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
			/// </summary>
			/// <param name="option" type="string | string[] | OfficeExtension.LoadOption"/>
			/// <returns type="Word.Setting"/>
		}

		Setting.prototype.set = function() {
			/// <signature>
			/// <summary>Sets multiple properties on the object at the same time, based on JSON input.</summary>
			/// <param name="properties" type="Word.Interfaces.SettingUpdateData">Properties described by the Word.Interfaces.SettingUpdateData interface.</param>
			/// <param name="options" type="string">Options of the form { throwOnReadOnly?: boolean }
			/// <br />
			/// * throwOnReadOnly: Throw an error if the passed-in property list includes read-only properties (default = true).
			/// </param>
			/// </signature>
			/// <signature>
			/// <summary>Sets multiple properties on the object at the same time, based on an existing loaded object.</summary>
			/// <param name="properties" type="Setting">An existing Setting object, with properties that have already been loaded and synced.</param>
			/// </signature>
		}
		Setting.prototype.delete = function() {
			/// <summary>
			/// Deletes the setting. [Api set: WordApi 1.4]
			/// </summary>
			/// <returns ></returns>
		}

		Setting.prototype.track = function() {
			/// <summary>
			/// Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
			/// </summary>
			/// <returns type="Word.Setting"/>
		}

		Setting.prototype.untrack = function() {
			/// <summary>
			/// Release the memory associated with this object, if has previous been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
			/// </summary>
			/// <returns type="Word.Setting"/>
		}

		return Setting;
	})(OfficeExtension.ClientObject);
	Word.Setting = Setting;
})(Word || (Word = {__proto__: null}));

var Word;
(function (Word) {
	var SettingCollection = (function(_super) {
		__extends(SettingCollection, _super);
		function SettingCollection() {
			/// <summary> Contains the collection of {@link Word.Setting} objects. [Api set: WordApi 1.4] </summary>
			/// <field name="context" type="Word.RequestContext">The request context associated with this object.</field>
			/// <field name="isNull" type="Boolean">Returns a boolean value for whether the corresponding object is null. You must call "context.sync()" before reading the isNull property.</field>
			/// <field name="items" type="Array" elementType="Word.Setting">Gets the loaded child items in this collection.</field>
		}

		SettingCollection.prototype.load = function(option) {
			/// <summary>
			/// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
			/// </summary>
			/// <param name="option" type="string | string[] | OfficeExtension.LoadOption"/>
			/// <returns type="Word.SettingCollection"/>
		}
		SettingCollection.prototype.add = function(key, value) {
			/// <summary>
			/// Creates a new setting or sets an existing setting. [Api set: WordApi 1.4]
			/// </summary>
			/// <param name="key" type="String">Required. The setting&apos;s key, which is case-sensitive.</param>
			/// <param name="value" >Required. The setting&apos;s value.</param>
			/// <returns type="Word.Setting"></returns>
		}
		SettingCollection.prototype.deleteAll = function() {
			/// <summary>
			/// Deletes all settings in this add-in. [Api set: WordApi 1.4]
			/// </summary>
			/// <returns ></returns>
		}
		SettingCollection.prototype.getCount = function() {
			/// <summary>
			/// Gets the count of settings. [Api set: WordApi 1.4]
			/// </summary>
			/// <returns type="OfficeExtension.ClientResult&lt;number&gt;"></returns>
			var result = new OfficeExtension.ClientResult();
			result.__proto__ = null;
			result.value = 0;
			return result;
		}
		SettingCollection.prototype.getItem = function(key) {
			/// <summary>
			/// Gets a setting object by its key, which is case-sensitive. Throws if the setting does not exist. [Api set: WordApi 1.4]
			/// </summary>
			/// <param name="key" type="String">The key that identifies the setting object.</param>
			/// <returns type="Word.Setting"></returns>
		}
		SettingCollection.prototype.getItemOrNullObject = function(key) {
			/// <summary>
			/// Gets a setting object by its key, which is case-sensitive. Returns a null object if the setting does not exist. [Api set: WordApi 1.4]
			/// </summary>
			/// <param name="key" type="String">Required. The key that identifies the setting object.</param>
			/// <returns type="Word.Setting"></returns>
		}

		SettingCollection.prototype.track = function() {
			/// <summary>
			/// Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
			/// </summary>
			/// <returns type="Word.SettingCollection"/>
		}

		SettingCollection.prototype.untrack = function() {
			/// <summary>
			/// Release the memory associated with this object, if has previous been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
			/// </summary>
			/// <returns type="Word.SettingCollection"/>
		}

		return SettingCollection;
	})(OfficeExtension.ClientObject);
	Word.SettingCollection = SettingCollection;
})(Word || (Word = {__proto__: null}));

var Word;
(function (Word) {
	/// <summary> [Api set: WordApi] </summary>
	var Style = {
		__proto__: null,
		"other": "other",
		"normal": "normal",
		"heading1": "heading1",
		"heading2": "heading2",
		"heading3": "heading3",
		"heading4": "heading4",
		"heading5": "heading5",
		"heading6": "heading6",
		"heading7": "heading7",
		"heading8": "heading8",
		"heading9": "heading9",
		"toc1": "toc1",
		"toc2": "toc2",
		"toc3": "toc3",
		"toc4": "toc4",
		"toc5": "toc5",
		"toc6": "toc6",
		"toc7": "toc7",
		"toc8": "toc8",
		"toc9": "toc9",
		"footnoteText": "footnoteText",
		"header": "header",
		"footer": "footer",
		"caption": "caption",
		"footnoteReference": "footnoteReference",
		"endnoteReference": "endnoteReference",
		"endnoteText": "endnoteText",
		"title": "title",
		"subtitle": "subtitle",
		"hyperlink": "hyperlink",
		"strong": "strong",
		"emphasis": "emphasis",
		"noSpacing": "noSpacing",
		"listParagraph": "listParagraph",
		"quote": "quote",
		"intenseQuote": "intenseQuote",
		"subtleEmphasis": "subtleEmphasis",
		"intenseEmphasis": "intenseEmphasis",
		"subtleReference": "subtleReference",
		"intenseReference": "intenseReference",
		"bookTitle": "bookTitle",
		"bibliography": "bibliography",
		"tocHeading": "tocHeading",
		"tableGrid": "tableGrid",
		"plainTable1": "plainTable1",
		"plainTable2": "plainTable2",
		"plainTable3": "plainTable3",
		"plainTable4": "plainTable4",
		"plainTable5": "plainTable5",
		"tableGridLight": "tableGridLight",
		"gridTable1Light": "gridTable1Light",
		"gridTable1Light_Accent1": "gridTable1Light_Accent1",
		"gridTable1Light_Accent2": "gridTable1Light_Accent2",
		"gridTable1Light_Accent3": "gridTable1Light_Accent3",
		"gridTable1Light_Accent4": "gridTable1Light_Accent4",
		"gridTable1Light_Accent5": "gridTable1Light_Accent5",
		"gridTable1Light_Accent6": "gridTable1Light_Accent6",
		"gridTable2": "gridTable2",
		"gridTable2_Accent1": "gridTable2_Accent1",
		"gridTable2_Accent2": "gridTable2_Accent2",
		"gridTable2_Accent3": "gridTable2_Accent3",
		"gridTable2_Accent4": "gridTable2_Accent4",
		"gridTable2_Accent5": "gridTable2_Accent5",
		"gridTable2_Accent6": "gridTable2_Accent6",
		"gridTable3": "gridTable3",
		"gridTable3_Accent1": "gridTable3_Accent1",
		"gridTable3_Accent2": "gridTable3_Accent2",
		"gridTable3_Accent3": "gridTable3_Accent3",
		"gridTable3_Accent4": "gridTable3_Accent4",
		"gridTable3_Accent5": "gridTable3_Accent5",
		"gridTable3_Accent6": "gridTable3_Accent6",
		"gridTable4": "gridTable4",
		"gridTable4_Accent1": "gridTable4_Accent1",
		"gridTable4_Accent2": "gridTable4_Accent2",
		"gridTable4_Accent3": "gridTable4_Accent3",
		"gridTable4_Accent4": "gridTable4_Accent4",
		"gridTable4_Accent5": "gridTable4_Accent5",
		"gridTable4_Accent6": "gridTable4_Accent6",
		"gridTable5Dark": "gridTable5Dark",
		"gridTable5Dark_Accent1": "gridTable5Dark_Accent1",
		"gridTable5Dark_Accent2": "gridTable5Dark_Accent2",
		"gridTable5Dark_Accent3": "gridTable5Dark_Accent3",
		"gridTable5Dark_Accent4": "gridTable5Dark_Accent4",
		"gridTable5Dark_Accent5": "gridTable5Dark_Accent5",
		"gridTable5Dark_Accent6": "gridTable5Dark_Accent6",
		"gridTable6Colorful": "gridTable6Colorful",
		"gridTable6Colorful_Accent1": "gridTable6Colorful_Accent1",
		"gridTable6Colorful_Accent2": "gridTable6Colorful_Accent2",
		"gridTable6Colorful_Accent3": "gridTable6Colorful_Accent3",
		"gridTable6Colorful_Accent4": "gridTable6Colorful_Accent4",
		"gridTable6Colorful_Accent5": "gridTable6Colorful_Accent5",
		"gridTable6Colorful_Accent6": "gridTable6Colorful_Accent6",
		"gridTable7Colorful": "gridTable7Colorful",
		"gridTable7Colorful_Accent1": "gridTable7Colorful_Accent1",
		"gridTable7Colorful_Accent2": "gridTable7Colorful_Accent2",
		"gridTable7Colorful_Accent3": "gridTable7Colorful_Accent3",
		"gridTable7Colorful_Accent4": "gridTable7Colorful_Accent4",
		"gridTable7Colorful_Accent5": "gridTable7Colorful_Accent5",
		"gridTable7Colorful_Accent6": "gridTable7Colorful_Accent6",
		"listTable1Light": "listTable1Light",
		"listTable1Light_Accent1": "listTable1Light_Accent1",
		"listTable1Light_Accent2": "listTable1Light_Accent2",
		"listTable1Light_Accent3": "listTable1Light_Accent3",
		"listTable1Light_Accent4": "listTable1Light_Accent4",
		"listTable1Light_Accent5": "listTable1Light_Accent5",
		"listTable1Light_Accent6": "listTable1Light_Accent6",
		"listTable2": "listTable2",
		"listTable2_Accent1": "listTable2_Accent1",
		"listTable2_Accent2": "listTable2_Accent2",
		"listTable2_Accent3": "listTable2_Accent3",
		"listTable2_Accent4": "listTable2_Accent4",
		"listTable2_Accent5": "listTable2_Accent5",
		"listTable2_Accent6": "listTable2_Accent6",
		"listTable3": "listTable3",
		"listTable3_Accent1": "listTable3_Accent1",
		"listTable3_Accent2": "listTable3_Accent2",
		"listTable3_Accent3": "listTable3_Accent3",
		"listTable3_Accent4": "listTable3_Accent4",
		"listTable3_Accent5": "listTable3_Accent5",
		"listTable3_Accent6": "listTable3_Accent6",
		"listTable4": "listTable4",
		"listTable4_Accent1": "listTable4_Accent1",
		"listTable4_Accent2": "listTable4_Accent2",
		"listTable4_Accent3": "listTable4_Accent3",
		"listTable4_Accent4": "listTable4_Accent4",
		"listTable4_Accent5": "listTable4_Accent5",
		"listTable4_Accent6": "listTable4_Accent6",
		"listTable5Dark": "listTable5Dark",
		"listTable5Dark_Accent1": "listTable5Dark_Accent1",
		"listTable5Dark_Accent2": "listTable5Dark_Accent2",
		"listTable5Dark_Accent3": "listTable5Dark_Accent3",
		"listTable5Dark_Accent4": "listTable5Dark_Accent4",
		"listTable5Dark_Accent5": "listTable5Dark_Accent5",
		"listTable5Dark_Accent6": "listTable5Dark_Accent6",
		"listTable6Colorful": "listTable6Colorful",
		"listTable6Colorful_Accent1": "listTable6Colorful_Accent1",
		"listTable6Colorful_Accent2": "listTable6Colorful_Accent2",
		"listTable6Colorful_Accent3": "listTable6Colorful_Accent3",
		"listTable6Colorful_Accent4": "listTable6Colorful_Accent4",
		"listTable6Colorful_Accent5": "listTable6Colorful_Accent5",
		"listTable6Colorful_Accent6": "listTable6Colorful_Accent6",
		"listTable7Colorful": "listTable7Colorful",
		"listTable7Colorful_Accent1": "listTable7Colorful_Accent1",
		"listTable7Colorful_Accent2": "listTable7Colorful_Accent2",
		"listTable7Colorful_Accent3": "listTable7Colorful_Accent3",
		"listTable7Colorful_Accent4": "listTable7Colorful_Accent4",
		"listTable7Colorful_Accent5": "listTable7Colorful_Accent5",
		"listTable7Colorful_Accent6": "listTable7Colorful_Accent6",
	}
	Word.Style = Style;
})(Word || (Word = {__proto__: null}));

var Word;
(function (Word) {
	var Table = (function(_super) {
		__extends(Table, _super);
		function Table() {
			/// <summary> Represents a table in a Word document. [Api set: WordApi 1.3] </summary>
			/// <field name="context" type="Word.RequestContext">The request context associated with this object.</field>
			/// <field name="isNull" type="Boolean">Returns a boolean value for whether the corresponding object is null. You must call "context.sync()" before reading the isNull property.</field>
			/// <field name="alignment" type="String">Gets or sets the alignment of the table against the page column. The value can be &apos;Left&apos;, &apos;Centered&apos;, or &apos;Right&apos;. [Api set: WordApi 1.3]</field>
			/// <field name="font" type="Word.Font">Gets the font. Use this to get and set font name, size, color, and other properties. Read-only. [Api set: WordApi 1.3]</field>
			/// <field name="headerRowCount" type="Number">Gets and sets the number of header rows. [Api set: WordApi 1.3]</field>
			/// <field name="horizontalAlignment" type="String">Gets and sets the horizontal alignment of every cell in the table. The value can be &apos;Left&apos;, &apos;Centered&apos;, &apos;Right&apos;, or &apos;Justified&apos;. [Api set: WordApi 1.3]</field>
			/// <field name="isUniform" type="Boolean">Indicates whether all of the table rows are uniform. Read-only. [Api set: WordApi 1.3]</field>
			/// <field name="nestingLevel" type="Number">Gets the nesting level of the table. Top-level tables have level 1. Read-only. [Api set: WordApi 1.3]</field>
			/// <field name="parentBody" type="Word.Body">Gets the parent body of the table. Read-only. [Api set: WordApi 1.3]</field>
			/// <field name="parentContentControl" type="Word.ContentControl">Gets the content control that contains the table. Throws if there isn&apos;t a parent content control. Read-only. [Api set: WordApi 1.3]</field>
			/// <field name="parentContentControlOrNullObject" type="Word.ContentControl">Gets the content control that contains the table. Returns a null object if there isn&apos;t a parent content control. Read-only. [Api set: WordApi 1.3]</field>
			/// <field name="parentTable" type="Word.Table">Gets the table that contains this table. Throws if it is not contained in a table. Read-only. [Api set: WordApi 1.3]</field>
			/// <field name="parentTableCell" type="Word.TableCell">Gets the table cell that contains this table. Throws if it is not contained in a table cell. Read-only. [Api set: WordApi 1.3]</field>
			/// <field name="parentTableCellOrNullObject" type="Word.TableCell">Gets the table cell that contains this table. Returns a null object if it is not contained in a table cell. Read-only. [Api set: WordApi 1.3]</field>
			/// <field name="parentTableOrNullObject" type="Word.Table">Gets the table that contains this table. Returns a null object if it is not contained in a table. Read-only. [Api set: WordApi 1.3]</field>
			/// <field name="rowCount" type="Number">Gets the number of rows in the table. Read-only. [Api set: WordApi 1.3]</field>
			/// <field name="rows" type="Word.TableRowCollection">Gets all of the table rows. Read-only. [Api set: WordApi 1.3]</field>
			/// <field name="shadingColor" type="String">Gets and sets the shading color. Color is specified in &quot;#RRGGBB&quot; format or by using the color name. [Api set: WordApi 1.3]</field>
			/// <field name="style" type="String">Gets or sets the style name for the table. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the &quot;styleBuiltIn&quot; property. [Api set: WordApi 1.3]</field>
			/// <field name="styleBandedColumns" type="Boolean">Gets and sets whether the table has banded columns. [Api set: WordApi 1.3]</field>
			/// <field name="styleBandedRows" type="Boolean">Gets and sets whether the table has banded rows. [Api set: WordApi 1.3]</field>
			/// <field name="styleBuiltIn" type="String">Gets or sets the built-in style name for the table. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the &quot;style&quot; property. [Api set: WordApi 1.3]</field>
			/// <field name="styleFirstColumn" type="Boolean">Gets and sets whether the table has a first column with a special style. [Api set: WordApi 1.3]</field>
			/// <field name="styleLastColumn" type="Boolean">Gets and sets whether the table has a last column with a special style. [Api set: WordApi 1.3]</field>
			/// <field name="styleTotalRow" type="Boolean">Gets and sets whether the table has a total (last) row with a special style. [Api set: WordApi 1.3]</field>
			/// <field name="tables" type="Word.TableCollection">Gets the child tables nested one level deeper. Read-only. [Api set: WordApi 1.3]</field>
			/// <field name="values" type="Array" elementType="Array">Gets and sets the text values in the table, as a 2D Javascript array. [Api set: WordApi 1.3]</field>
			/// <field name="verticalAlignment" type="String">Gets and sets the vertical alignment of every cell in the table. The value can be &apos;Top&apos;, &apos;Center&apos;, or &apos;Bottom&apos;. [Api set: WordApi 1.3]</field>
			/// <field name="width" type="Number">Gets and sets the width of the table in points. [Api set: WordApi 1.3]</field>
		}

		Table.prototype.load = function(option) {
			/// <summary>
			/// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
			/// </summary>
			/// <param name="option" type="string | string[] | OfficeExtension.LoadOption"/>
			/// <returns type="Word.Table"/>
		}

		Table.prototype.set = function() {
			/// <signature>
			/// <summary>Sets multiple properties on the object at the same time, based on JSON input.</summary>
			/// <param name="properties" type="Word.Interfaces.TableUpdateData">Properties described by the Word.Interfaces.TableUpdateData interface.</param>
			/// <param name="options" type="string">Options of the form { throwOnReadOnly?: boolean }
			/// <br />
			/// * throwOnReadOnly: Throw an error if the passed-in property list includes read-only properties (default = true).
			/// </param>
			/// </signature>
			/// <signature>
			/// <summary>Sets multiple properties on the object at the same time, based on an existing loaded object.</summary>
			/// <param name="properties" type="Table">An existing Table object, with properties that have already been loaded and synced.</param>
			/// </signature>
		}
		Table.prototype.addColumns = function(insertLocation, columnCount, values) {
			/// <summary>
			/// Adds columns to the start or end of the table, using the first or last existing column as a template. This is applicable to uniform tables. The string values, if specified, are set in the newly inserted rows. [Api set: WordApi 1.3]
			/// </summary>
			/// <param name="insertLocation" type="String">Required. It can be &apos;Start&apos; or &apos;End&apos;, corresponding to the appropriate side of the table.</param>
			/// <param name="columnCount" type="Number">Required. Number of columns to add.</param>
			/// <param name="values" type="Array" elementType="Array" optional="true">Optional 2D array. Cells are filled if the corresponding strings are specified in the array.</param>
			/// <returns ></returns>
		}
		Table.prototype.addRows = function(insertLocation, rowCount, values) {
			/// <summary>
			/// Adds rows to the start or end of the table, using the first or last existing row as a template. The string values, if specified, are set in the newly inserted rows. [Api set: WordApi 1.3]
			/// </summary>
			/// <param name="insertLocation" type="String">Required. It can be &apos;Start&apos; or &apos;End&apos;.</param>
			/// <param name="rowCount" type="Number">Required. Number of rows to add.</param>
			/// <param name="values" type="Array" elementType="Array" optional="true">Optional 2D array. Cells are filled if the corresponding strings are specified in the array.</param>
			/// <returns type="Word.TableRowCollection"></returns>
		}
		Table.prototype.autoFitWindow = function() {
			/// <summary>
			/// Autofits the table columns to the width of the window. [Api set: WordApi 1.3]
			/// </summary>
			/// <returns ></returns>
		}
		Table.prototype.clear = function() {
			/// <summary>
			/// Clears the contents of the table. [Api set: WordApi 1.3]
			/// </summary>
			/// <returns ></returns>
		}
		Table.prototype.delete = function() {
			/// <summary>
			/// Deletes the entire table. [Api set: WordApi 1.3]
			/// </summary>
			/// <returns ></returns>
		}
		Table.prototype.deleteColumns = function(columnIndex, columnCount) {
			/// <summary>
			/// Deletes specific columns. This is applicable to uniform tables. [Api set: WordApi 1.3]
			/// </summary>
			/// <param name="columnIndex" type="Number">Required. The first column to delete.</param>
			/// <param name="columnCount" type="Number" optional="true">Optional. The number of columns to delete. Default 1.</param>
			/// <returns ></returns>
		}
		Table.prototype.deleteRows = function(rowIndex, rowCount) {
			/// <summary>
			/// Deletes specific rows. [Api set: WordApi 1.3]
			/// </summary>
			/// <param name="rowIndex" type="Number">Required. The first row to delete.</param>
			/// <param name="rowCount" type="Number" optional="true">Optional. The number of rows to delete. Default 1.</param>
			/// <returns ></returns>
		}
		Table.prototype.distributeColumns = function() {
			/// <summary>
			/// Distributes the column widths evenly. This is applicable to uniform tables. [Api set: WordApi 1.3]
			/// </summary>
			/// <returns ></returns>
		}
		Table.prototype.getBorder = function(borderLocation) {
			/// <summary>
			/// Gets the border style for the specified border. [Api set: WordApi 1.3]
			/// </summary>
			/// <param name="borderLocation" type="String">Required. The border location.</param>
			/// <returns type="Word.TableBorder"></returns>
		}
		Table.prototype.getCell = function(rowIndex, cellIndex) {
			/// <summary>
			/// Gets the table cell at a specified row and column. Throws if the specified table cell does not exist. [Api set: WordApi 1.3]
			/// </summary>
			/// <param name="rowIndex" type="Number">Required. The index of the row.</param>
			/// <param name="cellIndex" type="Number">Required. The index of the cell in the row.</param>
			/// <returns type="Word.TableCell"></returns>
		}
		Table.prototype.getCellOrNullObject = function(rowIndex, cellIndex) {
			/// <summary>
			/// Gets the table cell at a specified row and column. Returns a null object if the specified table cell does not exist. [Api set: WordApi 1.3]
			/// </summary>
			/// <param name="rowIndex" type="Number">Required. The index of the row.</param>
			/// <param name="cellIndex" type="Number">Required. The index of the cell in the row.</param>
			/// <returns type="Word.TableCell"></returns>
		}
		Table.prototype.getCellPadding = function(cellPaddingLocation) {
			/// <summary>
			/// Gets cell padding in points. [Api set: WordApi 1.3]
			/// </summary>
			/// <param name="cellPaddingLocation" type="String">Required. The cell padding location can be &apos;Top&apos;, &apos;Left&apos;, &apos;Bottom&apos;, or &apos;Right&apos;.</param>
			/// <returns type="OfficeExtension.ClientResult&lt;number&gt;"></returns>
			var result = new OfficeExtension.ClientResult();
			result.__proto__ = null;
			result.value = 0;
			return result;
		}
		Table.prototype.getNext = function() {
			/// <summary>
			/// Gets the next table. Throws if this table is the last one. [Api set: WordApi 1.3]
			/// </summary>
			/// <returns type="Word.Table"></returns>
		}
		Table.prototype.getNextOrNullObject = function() {
			/// <summary>
			/// Gets the next table. Returns a null object if this table is the last one. [Api set: WordApi 1.3]
			/// </summary>
			/// <returns type="Word.Table"></returns>
		}
		Table.prototype.getParagraphAfter = function() {
			/// <summary>
			/// Gets the paragraph after the table. Throws if there isn&apos;t a paragraph after the table. [Api set: WordApi 1.3]
			/// </summary>
			/// <returns type="Word.Paragraph"></returns>
		}
		Table.prototype.getParagraphAfterOrNullObject = function() {
			/// <summary>
			/// Gets the paragraph after the table. Returns a null object if there isn&apos;t a paragraph after the table. [Api set: WordApi 1.3]
			/// </summary>
			/// <returns type="Word.Paragraph"></returns>
		}
		Table.prototype.getParagraphBefore = function() {
			/// <summary>
			/// Gets the paragraph before the table. Throws if there isn&apos;t a paragraph before the table. [Api set: WordApi 1.3]
			/// </summary>
			/// <returns type="Word.Paragraph"></returns>
		}
		Table.prototype.getParagraphBeforeOrNullObject = function() {
			/// <summary>
			/// Gets the paragraph before the table. Returns a null object if there isn&apos;t a paragraph before the table. [Api set: WordApi 1.3]
			/// </summary>
			/// <returns type="Word.Paragraph"></returns>
		}
		Table.prototype.getRange = function(rangeLocation) {
			/// <summary>
			/// Gets the range that contains this table, or the range at the start or end of the table. [Api set: WordApi 1.3]
			/// </summary>
			/// <param name="rangeLocation" type="String" optional="true">Optional. The range location can be &apos;Whole&apos;, &apos;Start&apos;, &apos;End&apos;, or &apos;After&apos;.</param>
			/// <returns type="Word.Range"></returns>
		}
		Table.prototype.insertContentControl = function() {
			/// <summary>
			/// Inserts a content control on the table. [Api set: WordApi 1.3]
			/// </summary>
			/// <returns type="Word.ContentControl"></returns>
		}
		Table.prototype.insertParagraph = function(paragraphText, insertLocation) {
			/// <summary>
			/// Inserts a paragraph at the specified location. The insertLocation value can be &apos;Before&apos; or &apos;After&apos;. [Api set: WordApi 1.3]
			/// </summary>
			/// <param name="paragraphText" type="String">Required. The paragraph text to be inserted.</param>
			/// <param name="insertLocation" type="String">Required. The value can be &apos;Before&apos; or &apos;After&apos;.</param>
			/// <returns type="Word.Paragraph"></returns>
		}
		Table.prototype.insertTable = function(rowCount, columnCount, insertLocation, values) {
			/// <summary>
			/// Inserts a table with the specified number of rows and columns. The insertLocation value can be &apos;Before&apos; or &apos;After&apos;. [Api set: WordApi 1.3]
			/// </summary>
			/// <param name="rowCount" type="Number">Required. The number of rows in the table.</param>
			/// <param name="columnCount" type="Number">Required. The number of columns in the table.</param>
			/// <param name="insertLocation" type="String">Required. The value can be &apos;Before&apos; or &apos;After&apos;.</param>
			/// <param name="values" type="Array" elementType="Array" optional="true">Optional 2D array. Cells are filled if the corresponding strings are specified in the array.</param>
			/// <returns type="Word.Table"></returns>
		}
		Table.prototype.mergeCells = function(topRow, firstCell, bottomRow, lastCell) {
			/// <summary>
			/// Merges the cells bounded inclusively by a first and last cell. [Api set: WordApi 1.4]
			/// </summary>
			/// <param name="topRow" type="Number">Required. The row of the first cell</param>
			/// <param name="firstCell" type="Number">Required. The index of the first cell in its row</param>
			/// <param name="bottomRow" type="Number">Required. The row of the last cell</param>
			/// <param name="lastCell" type="Number">Required. The index of the last cell in its row</param>
			/// <returns type="Word.TableCell"></returns>
		}
		Table.prototype.search = function(searchText, searchOptions) {
			/// <summary>
			/// Performs a search with the specified SearchOptions on the scope of the table object. The search results are a collection of range objects. [Api set: WordApi 1.3]
			/// </summary>
			/// <param name="searchText" type="String">Required. The search text.</param>
			/// <param name="searchOptions" type="Word.SearchOptions" optional="true">Optional. Options for the search.</param>
			/// <returns type="Word.RangeCollection"></returns>
		}
		Table.prototype.select = function(selectionMode) {
			/// <summary>
			/// Selects the table, or the position at the start or end of the table, and navigates the Word UI to it. [Api set: WordApi 1.3]
			/// </summary>
			/// <param name="selectionMode" type="String" optional="true">Optional. The selection mode can be &apos;Select&apos;, &apos;Start&apos;, or &apos;End&apos;. &apos;Select&apos; is the default.</param>
			/// <returns ></returns>
		}
		Table.prototype.setCellPadding = function(cellPaddingLocation, cellPadding) {
			/// <summary>
			/// Sets cell padding in points. [Api set: WordApi 1.3]
			/// </summary>
			/// <param name="cellPaddingLocation" type="String">Required. The cell padding location can be &apos;Top&apos;, &apos;Left&apos;, &apos;Bottom&apos;, or &apos;Right&apos;.</param>
			/// <param name="cellPadding" type="Number">Required. The cell padding.</param>
			/// <returns ></returns>
		}

		Table.prototype.track = function() {
			/// <summary>
			/// Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
			/// </summary>
			/// <returns type="Word.Table"/>
		}

		Table.prototype.untrack = function() {
			/// <summary>
			/// Release the memory associated with this object, if has previous been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
			/// </summary>
			/// <returns type="Word.Table"/>
		}

		return Table;
	})(OfficeExtension.ClientObject);
	Word.Table = Table;
})(Word || (Word = {__proto__: null}));

var Word;
(function (Word) {
	var TableBorder = (function(_super) {
		__extends(TableBorder, _super);
		function TableBorder() {
			/// <summary> Specifies the border style. [Api set: WordApi 1.3] </summary>
			/// <field name="context" type="Word.RequestContext">The request context associated with this object.</field>
			/// <field name="isNull" type="Boolean">Returns a boolean value for whether the corresponding object is null. You must call "context.sync()" before reading the isNull property.</field>
			/// <field name="color" type="String">Gets or sets the table border color. [Api set: WordApi 1.3]</field>
			/// <field name="type" type="String">Gets or sets the type of the table border. [Api set: WordApi 1.3]</field>
			/// <field name="width" type="Number">Gets or sets the width, in points, of the table border. Not applicable to table border types that have fixed widths. [Api set: WordApi 1.3]</field>
		}

		TableBorder.prototype.load = function(option) {
			/// <summary>
			/// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
			/// </summary>
			/// <param name="option" type="string | string[] | OfficeExtension.LoadOption"/>
			/// <returns type="Word.TableBorder"/>
		}

		TableBorder.prototype.set = function() {
			/// <signature>
			/// <summary>Sets multiple properties on the object at the same time, based on JSON input.</summary>
			/// <param name="properties" type="Word.Interfaces.TableBorderUpdateData">Properties described by the Word.Interfaces.TableBorderUpdateData interface.</param>
			/// <param name="options" type="string">Options of the form { throwOnReadOnly?: boolean }
			/// <br />
			/// * throwOnReadOnly: Throw an error if the passed-in property list includes read-only properties (default = true).
			/// </param>
			/// </signature>
			/// <signature>
			/// <summary>Sets multiple properties on the object at the same time, based on an existing loaded object.</summary>
			/// <param name="properties" type="TableBorder">An existing TableBorder object, with properties that have already been loaded and synced.</param>
			/// </signature>
		}

		TableBorder.prototype.track = function() {
			/// <summary>
			/// Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
			/// </summary>
			/// <returns type="Word.TableBorder"/>
		}

		TableBorder.prototype.untrack = function() {
			/// <summary>
			/// Release the memory associated with this object, if has previous been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
			/// </summary>
			/// <returns type="Word.TableBorder"/>
		}

		return TableBorder;
	})(OfficeExtension.ClientObject);
	Word.TableBorder = TableBorder;
})(Word || (Word = {__proto__: null}));

var Word;
(function (Word) {
	var TableCell = (function(_super) {
		__extends(TableCell, _super);
		function TableCell() {
			/// <summary> Represents a table cell in a Word document. [Api set: WordApi 1.3] </summary>
			/// <field name="context" type="Word.RequestContext">The request context associated with this object.</field>
			/// <field name="isNull" type="Boolean">Returns a boolean value for whether the corresponding object is null. You must call "context.sync()" before reading the isNull property.</field>
			/// <field name="body" type="Word.Body">Gets the body object of the cell. Read-only. [Api set: WordApi 1.3]</field>
			/// <field name="cellIndex" type="Number">Gets the index of the cell in its row. Read-only. [Api set: WordApi 1.3]</field>
			/// <field name="columnWidth" type="Number">Gets and sets the width of the cell&apos;s column in points. This is applicable to uniform tables. [Api set: WordApi 1.3]</field>
			/// <field name="horizontalAlignment" type="String">Gets and sets the horizontal alignment of the cell. The value can be &apos;Left&apos;, &apos;Centered&apos;, &apos;Right&apos;, or &apos;Justified&apos;. [Api set: WordApi 1.3]</field>
			/// <field name="parentRow" type="Word.TableRow">Gets the parent row of the cell. Read-only. [Api set: WordApi 1.3]</field>
			/// <field name="parentTable" type="Word.Table">Gets the parent table of the cell. Read-only. [Api set: WordApi 1.3]</field>
			/// <field name="rowIndex" type="Number">Gets the index of the cell&apos;s row in the table. Read-only. [Api set: WordApi 1.3]</field>
			/// <field name="shadingColor" type="String">Gets or sets the shading color of the cell. Color is specified in &quot;#RRGGBB&quot; format or by using the color name. [Api set: WordApi 1.3]</field>
			/// <field name="value" type="String">Gets and sets the text of the cell. [Api set: WordApi 1.3]</field>
			/// <field name="verticalAlignment" type="String">Gets and sets the vertical alignment of the cell. The value can be &apos;Top&apos;, &apos;Center&apos;, or &apos;Bottom&apos;. [Api set: WordApi 1.3]</field>
			/// <field name="width" type="Number">Gets the width of the cell in points. Read-only. [Api set: WordApi 1.3]</field>
		}

		TableCell.prototype.load = function(option) {
			/// <summary>
			/// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
			/// </summary>
			/// <param name="option" type="string | string[] | OfficeExtension.LoadOption"/>
			/// <returns type="Word.TableCell"/>
		}

		TableCell.prototype.set = function() {
			/// <signature>
			/// <summary>Sets multiple properties on the object at the same time, based on JSON input.</summary>
			/// <param name="properties" type="Word.Interfaces.TableCellUpdateData">Properties described by the Word.Interfaces.TableCellUpdateData interface.</param>
			/// <param name="options" type="string">Options of the form { throwOnReadOnly?: boolean }
			/// <br />
			/// * throwOnReadOnly: Throw an error if the passed-in property list includes read-only properties (default = true).
			/// </param>
			/// </signature>
			/// <signature>
			/// <summary>Sets multiple properties on the object at the same time, based on an existing loaded object.</summary>
			/// <param name="properties" type="TableCell">An existing TableCell object, with properties that have already been loaded and synced.</param>
			/// </signature>
		}
		TableCell.prototype.deleteColumn = function() {
			/// <summary>
			/// Deletes the column containing this cell. This is applicable to uniform tables. [Api set: WordApi 1.3]
			/// </summary>
			/// <returns ></returns>
		}
		TableCell.prototype.deleteRow = function() {
			/// <summary>
			/// Deletes the row containing this cell. [Api set: WordApi 1.3]
			/// </summary>
			/// <returns ></returns>
		}
		TableCell.prototype.getBorder = function(borderLocation) {
			/// <summary>
			/// Gets the border style for the specified border. [Api set: WordApi 1.3]
			/// </summary>
			/// <param name="borderLocation" type="String">Required. The border location.</param>
			/// <returns type="Word.TableBorder"></returns>
		}
		TableCell.prototype.getCellPadding = function(cellPaddingLocation) {
			/// <summary>
			/// Gets cell padding in points. [Api set: WordApi 1.3]
			/// </summary>
			/// <param name="cellPaddingLocation" type="String">Required. The cell padding location can be &apos;Top&apos;, &apos;Left&apos;, &apos;Bottom&apos;, or &apos;Right&apos;.</param>
			/// <returns type="OfficeExtension.ClientResult&lt;number&gt;"></returns>
			var result = new OfficeExtension.ClientResult();
			result.__proto__ = null;
			result.value = 0;
			return result;
		}
		TableCell.prototype.getNext = function() {
			/// <summary>
			/// Gets the next cell. Throws if this cell is the last one. [Api set: WordApi 1.3]
			/// </summary>
			/// <returns type="Word.TableCell"></returns>
		}
		TableCell.prototype.getNextOrNullObject = function() {
			/// <summary>
			/// Gets the next cell. Returns a null object if this cell is the last one. [Api set: WordApi 1.3]
			/// </summary>
			/// <returns type="Word.TableCell"></returns>
		}
		TableCell.prototype.insertColumns = function(insertLocation, columnCount, values) {
			/// <summary>
			/// Adds columns to the left or right of the cell, using the cell&apos;s column as a template. This is applicable to uniform tables. The string values, if specified, are set in the newly inserted rows. [Api set: WordApi 1.3]
			/// </summary>
			/// <param name="insertLocation" type="String">Required. It can be &apos;Before&apos; or &apos;After&apos;.</param>
			/// <param name="columnCount" type="Number">Required. Number of columns to add.</param>
			/// <param name="values" type="Array" elementType="Array" optional="true">Optional 2D array. Cells are filled if the corresponding strings are specified in the array.</param>
			/// <returns ></returns>
		}
		TableCell.prototype.insertRows = function(insertLocation, rowCount, values) {
			/// <summary>
			/// Inserts rows above or below the cell, using the cell&apos;s row as a template. The string values, if specified, are set in the newly inserted rows. [Api set: WordApi 1.3]
			/// </summary>
			/// <param name="insertLocation" type="String">Required. It can be &apos;Before&apos; or &apos;After&apos;.</param>
			/// <param name="rowCount" type="Number">Required. Number of rows to add.</param>
			/// <param name="values" type="Array" elementType="Array" optional="true">Optional 2D array. Cells are filled if the corresponding strings are specified in the array.</param>
			/// <returns type="Word.TableRowCollection"></returns>
		}
		TableCell.prototype.setCellPadding = function(cellPaddingLocation, cellPadding) {
			/// <summary>
			/// Sets cell padding in points. [Api set: WordApi 1.3]
			/// </summary>
			/// <param name="cellPaddingLocation" type="String">Required. The cell padding location can be &apos;Top&apos;, &apos;Left&apos;, &apos;Bottom&apos;, or &apos;Right&apos;.</param>
			/// <param name="cellPadding" type="Number">Required. The cell padding.</param>
			/// <returns ></returns>
		}
		TableCell.prototype.split = function(rowCount, columnCount) {
			/// <summary>
			/// Splits the cell into the specified number of rows and columns. [Api set: WordApi 1.4]
			/// </summary>
			/// <param name="rowCount" type="Number">Required. The number of rows to split into. Must be a divisor of the number of underlying rows.</param>
			/// <param name="columnCount" type="Number">Required. The number of columns to split into.</param>
			/// <returns ></returns>
		}

		TableCell.prototype.track = function() {
			/// <summary>
			/// Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
			/// </summary>
			/// <returns type="Word.TableCell"/>
		}

		TableCell.prototype.untrack = function() {
			/// <summary>
			/// Release the memory associated with this object, if has previous been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
			/// </summary>
			/// <returns type="Word.TableCell"/>
		}

		return TableCell;
	})(OfficeExtension.ClientObject);
	Word.TableCell = TableCell;
})(Word || (Word = {__proto__: null}));

var Word;
(function (Word) {
	var TableCellCollection = (function(_super) {
		__extends(TableCellCollection, _super);
		function TableCellCollection() {
			/// <summary> Contains the collection of the document&apos;s TableCell objects. [Api set: WordApi 1.3] </summary>
			/// <field name="context" type="Word.RequestContext">The request context associated with this object.</field>
			/// <field name="isNull" type="Boolean">Returns a boolean value for whether the corresponding object is null. You must call "context.sync()" before reading the isNull property.</field>
			/// <field name="items" type="Array" elementType="Word.TableCell">Gets the loaded child items in this collection.</field>
		}

		TableCellCollection.prototype.load = function(option) {
			/// <summary>
			/// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
			/// </summary>
			/// <param name="option" type="string | string[] | OfficeExtension.LoadOption"/>
			/// <returns type="Word.TableCellCollection"/>
		}
		TableCellCollection.prototype.getFirst = function() {
			/// <summary>
			/// Gets the first table cell in this collection. Throws if this collection is empty. [Api set: WordApi 1.3]
			/// </summary>
			/// <returns type="Word.TableCell"></returns>
		}
		TableCellCollection.prototype.getFirstOrNullObject = function() {
			/// <summary>
			/// Gets the first table cell in this collection. Returns a null object if this collection is empty. [Api set: WordApi 1.3]
			/// </summary>
			/// <returns type="Word.TableCell"></returns>
		}

		TableCellCollection.prototype.track = function() {
			/// <summary>
			/// Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
			/// </summary>
			/// <returns type="Word.TableCellCollection"/>
		}

		TableCellCollection.prototype.untrack = function() {
			/// <summary>
			/// Release the memory associated with this object, if has previous been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
			/// </summary>
			/// <returns type="Word.TableCellCollection"/>
		}

		return TableCellCollection;
	})(OfficeExtension.ClientObject);
	Word.TableCellCollection = TableCellCollection;
})(Word || (Word = {__proto__: null}));

var Word;
(function (Word) {
	var TableCollection = (function(_super) {
		__extends(TableCollection, _super);
		function TableCollection() {
			/// <summary> Contains the collection of the document&apos;s Table objects. [Api set: WordApi 1.3] </summary>
			/// <field name="context" type="Word.RequestContext">The request context associated with this object.</field>
			/// <field name="isNull" type="Boolean">Returns a boolean value for whether the corresponding object is null. You must call "context.sync()" before reading the isNull property.</field>
			/// <field name="items" type="Array" elementType="Word.Table">Gets the loaded child items in this collection.</field>
		}

		TableCollection.prototype.load = function(option) {
			/// <summary>
			/// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
			/// </summary>
			/// <param name="option" type="string | string[] | OfficeExtension.LoadOption"/>
			/// <returns type="Word.TableCollection"/>
		}
		TableCollection.prototype.getFirst = function() {
			/// <summary>
			/// Gets the first table in this collection. Throws if this collection is empty. [Api set: WordApi 1.3]
			/// </summary>
			/// <returns type="Word.Table"></returns>
		}
		TableCollection.prototype.getFirstOrNullObject = function() {
			/// <summary>
			/// Gets the first table in this collection. Returns a null object if this collection is empty. [Api set: WordApi 1.3]
			/// </summary>
			/// <returns type="Word.Table"></returns>
		}

		TableCollection.prototype.track = function() {
			/// <summary>
			/// Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
			/// </summary>
			/// <returns type="Word.TableCollection"/>
		}

		TableCollection.prototype.untrack = function() {
			/// <summary>
			/// Release the memory associated with this object, if has previous been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
			/// </summary>
			/// <returns type="Word.TableCollection"/>
		}

		return TableCollection;
	})(OfficeExtension.ClientObject);
	Word.TableCollection = TableCollection;
})(Word || (Word = {__proto__: null}));

var Word;
(function (Word) {
	var TableRow = (function(_super) {
		__extends(TableRow, _super);
		function TableRow() {
			/// <summary> Represents a row in a Word document. [Api set: WordApi 1.3] </summary>
			/// <field name="context" type="Word.RequestContext">The request context associated with this object.</field>
			/// <field name="isNull" type="Boolean">Returns a boolean value for whether the corresponding object is null. You must call "context.sync()" before reading the isNull property.</field>
			/// <field name="cellCount" type="Number">Gets the number of cells in the row. Read-only. [Api set: WordApi 1.3]</field>
			/// <field name="cells" type="Word.TableCellCollection">Gets cells. Read-only. [Api set: WordApi 1.3]</field>
			/// <field name="font" type="Word.Font">Gets the font. Use this to get and set font name, size, color, and other properties. Read-only. [Api set: WordApi 1.3]</field>
			/// <field name="horizontalAlignment" type="String">Gets and sets the horizontal alignment of every cell in the row. The value can be &apos;Left&apos;, &apos;Centered&apos;, &apos;Right&apos;, or &apos;Justified&apos;. [Api set: WordApi 1.3]</field>
			/// <field name="isHeader" type="Boolean">Checks whether the row is a header row. Read-only. To set the number of header rows, use HeaderRowCount on the Table object. [Api set: WordApi 1.3]</field>
			/// <field name="parentTable" type="Word.Table">Gets parent table. Read-only. [Api set: WordApi 1.3]</field>
			/// <field name="preferredHeight" type="Number">Gets and sets the preferred height of the row in points. [Api set: WordApi 1.3]</field>
			/// <field name="rowIndex" type="Number">Gets the index of the row in its parent table. Read-only. [Api set: WordApi 1.3]</field>
			/// <field name="shadingColor" type="String">Gets and sets the shading color. Color is specified in &quot;#RRGGBB&quot; format or by using the color name. [Api set: WordApi 1.3]</field>
			/// <field name="values" type="Array" elementType="Array">Gets and sets the text values in the row, as a 2D Javascript array. [Api set: WordApi 1.3]</field>
			/// <field name="verticalAlignment" type="String">Gets and sets the vertical alignment of the cells in the row. The value can be &apos;Top&apos;, &apos;Center&apos;, or &apos;Bottom&apos;. [Api set: WordApi 1.3]</field>
		}

		TableRow.prototype.load = function(option) {
			/// <summary>
			/// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
			/// </summary>
			/// <param name="option" type="string | string[] | OfficeExtension.LoadOption"/>
			/// <returns type="Word.TableRow"/>
		}

		TableRow.prototype.set = function() {
			/// <signature>
			/// <summary>Sets multiple properties on the object at the same time, based on JSON input.</summary>
			/// <param name="properties" type="Word.Interfaces.TableRowUpdateData">Properties described by the Word.Interfaces.TableRowUpdateData interface.</param>
			/// <param name="options" type="string">Options of the form { throwOnReadOnly?: boolean }
			/// <br />
			/// * throwOnReadOnly: Throw an error if the passed-in property list includes read-only properties (default = true).
			/// </param>
			/// </signature>
			/// <signature>
			/// <summary>Sets multiple properties on the object at the same time, based on an existing loaded object.</summary>
			/// <param name="properties" type="TableRow">An existing TableRow object, with properties that have already been loaded and synced.</param>
			/// </signature>
		}
		TableRow.prototype.clear = function() {
			/// <summary>
			/// Clears the contents of the row. [Api set: WordApi 1.3]
			/// </summary>
			/// <returns ></returns>
		}
		TableRow.prototype.delete = function() {
			/// <summary>
			/// Deletes the entire row. [Api set: WordApi 1.3]
			/// </summary>
			/// <returns ></returns>
		}
		TableRow.prototype.getBorder = function(borderLocation) {
			/// <summary>
			/// Gets the border style of the cells in the row. [Api set: WordApi 1.3]
			/// </summary>
			/// <param name="borderLocation" type="String">Required. The border location.</param>
			/// <returns type="Word.TableBorder"></returns>
		}
		TableRow.prototype.getCellPadding = function(cellPaddingLocation) {
			/// <summary>
			/// Gets cell padding in points. [Api set: WordApi 1.3]
			/// </summary>
			/// <param name="cellPaddingLocation" type="String">Required. The cell padding location can be &apos;Top&apos;, &apos;Left&apos;, &apos;Bottom&apos;, or &apos;Right&apos;.</param>
			/// <returns type="OfficeExtension.ClientResult&lt;number&gt;"></returns>
			var result = new OfficeExtension.ClientResult();
			result.__proto__ = null;
			result.value = 0;
			return result;
		}
		TableRow.prototype.getNext = function() {
			/// <summary>
			/// Gets the next row. Throws if this row is the last one. [Api set: WordApi 1.3]
			/// </summary>
			/// <returns type="Word.TableRow"></returns>
		}
		TableRow.prototype.getNextOrNullObject = function() {
			/// <summary>
			/// Gets the next row. Returns a null object if this row is the last one. [Api set: WordApi 1.3]
			/// </summary>
			/// <returns type="Word.TableRow"></returns>
		}
		TableRow.prototype.insertContentControl = function() {
			/// <summary>
			/// Inserts a content control on the row. [Api set: WordApi 1.4]
			/// </summary>
			/// <returns type="Word.ContentControl"></returns>
		}
		TableRow.prototype.insertRows = function(insertLocation, rowCount, values) {
			/// <summary>
			/// Inserts rows using this row as a template. If values are specified, inserts the values into the new rows. [Api set: WordApi 1.3]
			/// </summary>
			/// <param name="insertLocation" type="String">Required. Where the new rows should be inserted, relative to the current row. It can be &apos;Before&apos; or &apos;After&apos;.</param>
			/// <param name="rowCount" type="Number">Required. Number of rows to add</param>
			/// <param name="values" type="Array" elementType="Array" optional="true">Optional. Strings to insert in the new rows, specified as a 2D array. The number of cells in each row must not exceed the number of cells in the existing row.</param>
			/// <returns type="Word.TableRowCollection"></returns>
		}
		TableRow.prototype.merge = function() {
			/// <summary>
			/// Merges the row into one cell. [Api set: WordApi 1.4]
			/// </summary>
			/// <returns type="Word.TableCell"></returns>
		}
		TableRow.prototype.search = function(searchText, searchOptions) {
			/// <summary>
			/// Performs a search with the specified SearchOptions on the scope of the row. The search results are a collection of range objects. [Api set: WordApi 1.3]
			/// </summary>
			/// <param name="searchText" type="String">Required. The search text.</param>
			/// <param name="searchOptions" type="Word.SearchOptions" optional="true">Optional. Options for the search.</param>
			/// <returns type="Word.RangeCollection"></returns>
		}
		TableRow.prototype.select = function(selectionMode) {
			/// <summary>
			/// Selects the row and navigates the Word UI to it. [Api set: WordApi 1.3]
			/// </summary>
			/// <param name="selectionMode" type="String" optional="true">Optional. The selection mode can be &apos;Select&apos;, &apos;Start&apos;, or &apos;End&apos;. &apos;Select&apos; is the default.</param>
			/// <returns ></returns>
		}
		TableRow.prototype.setCellPadding = function(cellPaddingLocation, cellPadding) {
			/// <summary>
			/// Sets cell padding in points. [Api set: WordApi 1.3]
			/// </summary>
			/// <param name="cellPaddingLocation" type="String">Required. The cell padding location can be &apos;Top&apos;, &apos;Left&apos;, &apos;Bottom&apos;, or &apos;Right&apos;.</param>
			/// <param name="cellPadding" type="Number">Required. The cell padding.</param>
			/// <returns ></returns>
		}

		TableRow.prototype.track = function() {
			/// <summary>
			/// Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
			/// </summary>
			/// <returns type="Word.TableRow"/>
		}

		TableRow.prototype.untrack = function() {
			/// <summary>
			/// Release the memory associated with this object, if has previous been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
			/// </summary>
			/// <returns type="Word.TableRow"/>
		}

		return TableRow;
	})(OfficeExtension.ClientObject);
	Word.TableRow = TableRow;
})(Word || (Word = {__proto__: null}));

var Word;
(function (Word) {
	var TableRowCollection = (function(_super) {
		__extends(TableRowCollection, _super);
		function TableRowCollection() {
			/// <summary> Contains the collection of the document&apos;s TableRow objects. [Api set: WordApi 1.3] </summary>
			/// <field name="context" type="Word.RequestContext">The request context associated with this object.</field>
			/// <field name="isNull" type="Boolean">Returns a boolean value for whether the corresponding object is null. You must call "context.sync()" before reading the isNull property.</field>
			/// <field name="items" type="Array" elementType="Word.TableRow">Gets the loaded child items in this collection.</field>
		}

		TableRowCollection.prototype.load = function(option) {
			/// <summary>
			/// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
			/// </summary>
			/// <param name="option" type="string | string[] | OfficeExtension.LoadOption"/>
			/// <returns type="Word.TableRowCollection"/>
		}
		TableRowCollection.prototype.getFirst = function() {
			/// <summary>
			/// Gets the first row in this collection. Throws if this collection is empty. [Api set: WordApi 1.3]
			/// </summary>
			/// <returns type="Word.TableRow"></returns>
		}
		TableRowCollection.prototype.getFirstOrNullObject = function() {
			/// <summary>
			/// Gets the first row in this collection. Returns a null object if this collection is empty. [Api set: WordApi 1.3]
			/// </summary>
			/// <returns type="Word.TableRow"></returns>
		}

		TableRowCollection.prototype.track = function() {
			/// <summary>
			/// Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
			/// </summary>
			/// <returns type="Word.TableRowCollection"/>
		}

		TableRowCollection.prototype.untrack = function() {
			/// <summary>
			/// Release the memory associated with this object, if has previous been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
			/// </summary>
			/// <returns type="Word.TableRowCollection"/>
		}

		return TableRowCollection;
	})(OfficeExtension.ClientObject);
	Word.TableRowCollection = TableRowCollection;
})(Word || (Word = {__proto__: null}));

var Word;
(function (Word) {
	/// <summary> [Api set: WordApi] </summary>
	var TapObjectType = {
		__proto__: null,
		"chart": "chart",
		"smartArt": "smartArt",
		"table": "table",
		"image": "image",
		"slide": "slide",
		"ole": "ole",
		"text": "text",
	}
	Word.TapObjectType = TapObjectType;
})(Word || (Word = {__proto__: null}));

var Word;
(function (Word) {
	/// <summary> Underline types [Api set: WordApi] </summary>
	var UnderlineType = {
		__proto__: null,
		"mixed": "mixed",
		"none": "none",
		"single": "single",
		"word": "word",
		"double": "double",
		"thick": "thick",
		"dotted": "dotted",
		"dottedHeavy": "dottedHeavy",
		"dashLine": "dashLine",
		"dashLineHeavy": "dashLineHeavy",
		"dashLineLong": "dashLineLong",
		"dashLineLongHeavy": "dashLineLongHeavy",
		"dotDashLine": "dotDashLine",
		"dotDashLineHeavy": "dotDashLineHeavy",
		"twoDotDashLine": "twoDotDashLine",
		"twoDotDashLineHeavy": "twoDotDashLineHeavy",
		"wave": "wave",
		"waveHeavy": "waveHeavy",
		"waveDouble": "waveDouble",
	}
	Word.UnderlineType = UnderlineType;
})(Word || (Word = {__proto__: null}));

var Word;
(function (Word) {
	/// <summary> [Api set: WordApi] </summary>
	var VerticalAlignment = {
		__proto__: null,
		"mixed": "mixed",
		"top": "top",
		"center": "center",
		"bottom": "bottom",
	}
	Word.VerticalAlignment = VerticalAlignment;
})(Word || (Word = {__proto__: null}));

var Word;
(function (Word) {
	var Interfaces;
	(function (Interfaces) {
		var BodyUpdateData = (function() {
			function BodyUpdateData() {
				/// <summary>An interface for updating data on the Body object, for use in "body.set({ ... })".</summary>
				/// <field name="font" type="Word.Interfaces.FontUpdateData">Gets the text format of the body. Use this to get and set font name, size, color and other properties. [Api set: WordApi 1.1]</field>
				/// <field name="style" type="String">Gets or sets the style name for the body. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the &quot;styleBuiltIn&quot; property. [Api set: WordApi 1.1]</field>;
				/// <field name="styleBuiltIn" type="String">Gets or sets the built-in style name for the body. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the &quot;style&quot; property. [Api set: WordApi 1.3]</field>;
			}
			return BodyUpdateData;
		})();
		Interfaces.BodyUpdateData.__proto__ = null;
		Interfaces.BodyUpdateData = BodyUpdateData;
	})(Interfaces = Word.Interfaces || (Word.Interfaces = { __proto__: null}));
})(Word || (Word = {__proto__: null}));

var Word;
(function (Word) {
	var Interfaces;
	(function (Interfaces) {
		var ContentControlUpdateData = (function() {
			function ContentControlUpdateData() {
				/// <summary>An interface for updating data on the ContentControl object, for use in "contentControl.set({ ... })".</summary>
				/// <field name="font" type="Word.Interfaces.FontUpdateData">Gets the text format of the content control. Use this to get and set font name, size, color, and other properties. [Api set: WordApi 1.1]</field>
				/// <field name="appearance" type="String">Gets or sets the appearance of the content control. The value can be &apos;BoundingBox&apos;, &apos;Tags&apos;, or &apos;Hidden&apos;. [Api set: WordApi 1.1]</field>;
				/// <field name="cannotDelete" type="Boolean">Gets or sets a value that indicates whether the user can delete the content control. Mutually exclusive with removeWhenEdited. [Api set: WordApi 1.1]</field>;
				/// <field name="cannotEdit" type="Boolean">Gets or sets a value that indicates whether the user can edit the contents of the content control. [Api set: WordApi 1.1]</field>;
				/// <field name="color" type="String">Gets or sets the color of the content control. Color is specified in &apos;#RRGGBB&apos; format or by using the color name. [Api set: WordApi 1.1]</field>;
				/// <field name="placeholderText" type="String">Gets or sets the placeholder text of the content control. Dimmed text will be displayed when the content control is empty. [Api set: WordApi 1.1]</field>;
				/// <field name="removeWhenEdited" type="Boolean">Gets or sets a value that indicates whether the content control is removed after it is edited. Mutually exclusive with cannotDelete. [Api set: WordApi 1.1]</field>;
				/// <field name="style" type="String">Gets or sets the style name for the content control. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the &quot;styleBuiltIn&quot; property. [Api set: WordApi 1.1]</field>;
				/// <field name="styleBuiltIn" type="String">Gets or sets the built-in style name for the content control. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the &quot;style&quot; property. [Api set: WordApi 1.3]</field>;
				/// <field name="tag" type="String">Gets or sets a tag to identify a content control. [Api set: WordApi 1.1]</field>;
				/// <field name="title" type="String">Gets or sets the title for a content control. [Api set: WordApi 1.1]</field>;
			}
			return ContentControlUpdateData;
		})();
		Interfaces.ContentControlUpdateData.__proto__ = null;
		Interfaces.ContentControlUpdateData = ContentControlUpdateData;
	})(Interfaces = Word.Interfaces || (Word.Interfaces = { __proto__: null}));
})(Word || (Word = {__proto__: null}));

var Word;
(function (Word) {
	var Interfaces;
	(function (Interfaces) {
		var CustomPropertyUpdateData = (function() {
			function CustomPropertyUpdateData() {
				/// <summary>An interface for updating data on the CustomProperty object, for use in "customProperty.set({ ... })".</summary>
				/// <field name="value" >Gets or sets the value of the custom property. Note that even though Word Online and the docx file format allow these properties to be arbitrarily long, the desktop version of Word will truncate string values to 255 16-bit chars (possibly creating invalid unicode by breaking up a surrogate pair). [Api set: WordApi 1.3]</field>;
			}
			return CustomPropertyUpdateData;
		})();
		Interfaces.CustomPropertyUpdateData.__proto__ = null;
		Interfaces.CustomPropertyUpdateData = CustomPropertyUpdateData;
	})(Interfaces = Word.Interfaces || (Word.Interfaces = { __proto__: null}));
})(Word || (Word = {__proto__: null}));

var Word;
(function (Word) {
	var Interfaces;
	(function (Interfaces) {
		var DocumentUpdateData = (function() {
			function DocumentUpdateData() {
				/// <summary>An interface for updating data on the Document object, for use in "document.set({ ... })".</summary>
				/// <field name="body" type="Word.Interfaces.BodyUpdateData">Gets the body object of the document. The body is the text that excludes headers, footers, footnotes, textboxes, etc.. [Api set: WordApi 1.1]</field>
				/// <field name="properties" type="Word.Interfaces.DocumentPropertiesUpdateData">Gets the properties of the document. [Api set: WordApi 1.3]</field>
				/// <field name="allowCloseOnUntitled" type="Boolean">Gets or sets a value that indicates that, when opening a new document, whether it is allowed to close this document even if this document is untitled. True to close, false otherwise. [Api set: WordApi]</field>;
			}
			return DocumentUpdateData;
		})();
		Interfaces.DocumentUpdateData.__proto__ = null;
		Interfaces.DocumentUpdateData = DocumentUpdateData;
	})(Interfaces = Word.Interfaces || (Word.Interfaces = { __proto__: null}));
})(Word || (Word = {__proto__: null}));

var Word;
(function (Word) {
	var Interfaces;
	(function (Interfaces) {
		var DocumentCreatedUpdateData = (function() {
			function DocumentCreatedUpdateData() {
				/// <summary>An interface for updating data on the DocumentCreated object, for use in "documentCreated.set({ ... })".</summary>
				/// <field name="body" type="Word.Interfaces.BodyUpdateData">Gets the body object of the document. The body is the text that excludes headers, footers, footnotes, textboxes, etc.. [Api set: WordApiHiddenDocument 1.3]</field>
				/// <field name="properties" type="Word.Interfaces.DocumentPropertiesUpdateData">Gets the properties of the document. [Api set: WordApiHiddenDocument 1.3]</field>
			}
			return DocumentCreatedUpdateData;
		})();
		Interfaces.DocumentCreatedUpdateData.__proto__ = null;
		Interfaces.DocumentCreatedUpdateData = DocumentCreatedUpdateData;
	})(Interfaces = Word.Interfaces || (Word.Interfaces = { __proto__: null}));
})(Word || (Word = {__proto__: null}));

var Word;
(function (Word) {
	var Interfaces;
	(function (Interfaces) {
		var DocumentPropertiesUpdateData = (function() {
			function DocumentPropertiesUpdateData() {
				/// <summary>An interface for updating data on the DocumentProperties object, for use in "documentProperties.set({ ... })".</summary>
				/// <field name="author" type="String">Gets or sets the author of the document. [Api set: WordApi 1.3]</field>;
				/// <field name="category" type="String">Gets or sets the category of the document. [Api set: WordApi 1.3]</field>;
				/// <field name="comments" type="String">Gets or sets the comments of the document. [Api set: WordApi 1.3]</field>;
				/// <field name="company" type="String">Gets or sets the company of the document. [Api set: WordApi 1.3]</field>;
				/// <field name="format" type="String">Gets or sets the format of the document. [Api set: WordApi 1.3]</field>;
				/// <field name="keywords" type="String">Gets or sets the keywords of the document. [Api set: WordApi 1.3]</field>;
				/// <field name="manager" type="String">Gets or sets the manager of the document. [Api set: WordApi 1.3]</field>;
				/// <field name="subject" type="String">Gets or sets the subject of the document. [Api set: WordApi 1.3]</field>;
				/// <field name="title" type="String">Gets or sets the title of the document. [Api set: WordApi 1.3]</field>;
			}
			return DocumentPropertiesUpdateData;
		})();
		Interfaces.DocumentPropertiesUpdateData.__proto__ = null;
		Interfaces.DocumentPropertiesUpdateData = DocumentPropertiesUpdateData;
	})(Interfaces = Word.Interfaces || (Word.Interfaces = { __proto__: null}));
})(Word || (Word = {__proto__: null}));

var Word;
(function (Word) {
	var Interfaces;
	(function (Interfaces) {
		var FontUpdateData = (function() {
			function FontUpdateData() {
				/// <summary>An interface for updating data on the Font object, for use in "font.set({ ... })".</summary>
				/// <field name="bold" type="Boolean">Gets or sets a value that indicates whether the font is bold. True if the font is formatted as bold, otherwise, false. [Api set: WordApi 1.1]</field>;
				/// <field name="color" type="String">Gets or sets the color for the specified font. You can provide the value in the &apos;#RRGGBB&apos; format or the color name. [Api set: WordApi 1.1]</field>;
				/// <field name="doubleStrikeThrough" type="Boolean">Gets or sets a value that indicates whether the font has a double strikethrough. True if the font is formatted as double strikethrough text, otherwise, false. [Api set: WordApi 1.1]</field>;
				/// <field name="highlightColor" type="String">Gets or sets the highlight color. To set it, use a value either in the &apos;#RRGGBB&apos; format or the color name. To remove highlight color, set it to null. The returned highlight color can be in the &apos;#RRGGBB&apos; format, an empty string for mixed highlight colors, or null for no highlight color. [Api set: WordApi 1.1]</field>;
				/// <field name="italic" type="Boolean">Gets or sets a value that indicates whether the font is italicized. True if the font is italicized, otherwise, false. [Api set: WordApi 1.1]</field>;
				/// <field name="name" type="String">Gets or sets a value that represents the name of the font. [Api set: WordApi 1.1]</field>;
				/// <field name="size" type="Number">Gets or sets a value that represents the font size in points. [Api set: WordApi 1.1]</field>;
				/// <field name="strikeThrough" type="Boolean">Gets or sets a value that indicates whether the font has a strikethrough. True if the font is formatted as strikethrough text, otherwise, false. [Api set: WordApi 1.1]</field>;
				/// <field name="subscript" type="Boolean">Gets or sets a value that indicates whether the font is a subscript. True if the font is formatted as subscript, otherwise, false. [Api set: WordApi 1.1]</field>;
				/// <field name="superscript" type="Boolean">Gets or sets a value that indicates whether the font is a superscript. True if the font is formatted as superscript, otherwise, false. [Api set: WordApi 1.1]</field>;
				/// <field name="underline" type="String">Gets or sets a value that indicates the font&apos;s underline type. &apos;None&apos; if the font is not underlined. [Api set: WordApi 1.1]</field>;
			}
			return FontUpdateData;
		})();
		Interfaces.FontUpdateData.__proto__ = null;
		Interfaces.FontUpdateData = FontUpdateData;
	})(Interfaces = Word.Interfaces || (Word.Interfaces = { __proto__: null}));
})(Word || (Word = {__proto__: null}));

var Word;
(function (Word) {
	var Interfaces;
	(function (Interfaces) {
		var InlinePictureUpdateData = (function() {
			function InlinePictureUpdateData() {
				/// <summary>An interface for updating data on the InlinePicture object, for use in "inlinePicture.set({ ... })".</summary>
				/// <field name="altTextDescription" type="String">Gets or sets a string that represents the alternative text associated with the inline image. [Api set: WordApi 1.1]</field>;
				/// <field name="altTextTitle" type="String">Gets or sets a string that contains the title for the inline image. [Api set: WordApi 1.1]</field>;
				/// <field name="height" type="Number">Gets or sets a number that describes the height of the inline image. [Api set: WordApi 1.1]</field>;
				/// <field name="hyperlink" type="String">Gets or sets a hyperlink on the image. Use a &apos;#&apos; to separate the address part from the optional location part. [Api set: WordApi 1.1]</field>;
				/// <field name="lockAspectRatio" type="Boolean">Gets or sets a value that indicates whether the inline image retains its original proportions when you resize it. [Api set: WordApi 1.1]</field>;
				/// <field name="width" type="Number">Gets or sets a number that describes the width of the inline image. [Api set: WordApi 1.1]</field>;
			}
			return InlinePictureUpdateData;
		})();
		Interfaces.InlinePictureUpdateData.__proto__ = null;
		Interfaces.InlinePictureUpdateData = InlinePictureUpdateData;
	})(Interfaces = Word.Interfaces || (Word.Interfaces = { __proto__: null}));
})(Word || (Word = {__proto__: null}));

var Word;
(function (Word) {
	var Interfaces;
	(function (Interfaces) {
		var ListItemUpdateData = (function() {
			function ListItemUpdateData() {
				/// <summary>An interface for updating data on the ListItem object, for use in "listItem.set({ ... })".</summary>
				/// <field name="level" type="Number">Gets or sets the level of the item in the list. [Api set: WordApi 1.3]</field>;
			}
			return ListItemUpdateData;
		})();
		Interfaces.ListItemUpdateData.__proto__ = null;
		Interfaces.ListItemUpdateData = ListItemUpdateData;
	})(Interfaces = Word.Interfaces || (Word.Interfaces = { __proto__: null}));
})(Word || (Word = {__proto__: null}));

var Word;
(function (Word) {
	var Interfaces;
	(function (Interfaces) {
		var ParagraphUpdateData = (function() {
			function ParagraphUpdateData() {
				/// <summary>An interface for updating data on the Paragraph object, for use in "paragraph.set({ ... })".</summary>
				/// <field name="font" type="Word.Interfaces.FontUpdateData">Gets the text format of the paragraph. Use this to get and set font name, size, color, and other properties. [Api set: WordApi 1.1]</field>
				/// <field name="listItem" type="Word.Interfaces.ListItemUpdateData">Gets the ListItem for the paragraph. Throws if the paragraph is not part of a list. [Api set: WordApi 1.3]</field>
				/// <field name="listItemOrNullObject" type="Word.Interfaces.ListItemUpdateData">Gets the ListItem for the paragraph. Returns a null object if the paragraph is not part of a list. [Api set: WordApi 1.3]</field>
				/// <field name="alignment" type="String">Gets or sets the alignment for a paragraph. The value can be &apos;left&apos;, &apos;centered&apos;, &apos;right&apos;, or &apos;justified&apos;. [Api set: WordApi 1.1]</field>;
				/// <field name="firstLineIndent" type="Number">Gets or sets the value, in points, for a first line or hanging indent. Use a positive value to set a first-line indent, and use a negative value to set a hanging indent. [Api set: WordApi 1.1]</field>;
				/// <field name="leftIndent" type="Number">Gets or sets the left indent value, in points, for the paragraph. [Api set: WordApi 1.1]</field>;
				/// <field name="lineSpacing" type="Number">Gets or sets the line spacing, in points, for the specified paragraph. In the Word UI, this value is divided by 12. [Api set: WordApi 1.1]</field>;
				/// <field name="lineUnitAfter" type="Number">Gets or sets the amount of spacing, in grid lines, after the paragraph. [Api set: WordApi 1.1]</field>;
				/// <field name="lineUnitBefore" type="Number">Gets or sets the amount of spacing, in grid lines, before the paragraph. [Api set: WordApi 1.1]</field>;
				/// <field name="outlineLevel" type="Number">Gets or sets the outline level for the paragraph. [Api set: WordApi 1.1]</field>;
				/// <field name="rightIndent" type="Number">Gets or sets the right indent value, in points, for the paragraph. [Api set: WordApi 1.1]</field>;
				/// <field name="spaceAfter" type="Number">Gets or sets the spacing, in points, after the paragraph. [Api set: WordApi 1.1]</field>;
				/// <field name="spaceBefore" type="Number">Gets or sets the spacing, in points, before the paragraph. [Api set: WordApi 1.1]</field>;
				/// <field name="style" type="String">Gets or sets the style name for the paragraph. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the &quot;styleBuiltIn&quot; property. [Api set: WordApi 1.1]</field>;
				/// <field name="styleBuiltIn" type="String">Gets or sets the built-in style name for the paragraph. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the &quot;style&quot; property. [Api set: WordApi 1.3]</field>;
			}
			return ParagraphUpdateData;
		})();
		Interfaces.ParagraphUpdateData.__proto__ = null;
		Interfaces.ParagraphUpdateData = ParagraphUpdateData;
	})(Interfaces = Word.Interfaces || (Word.Interfaces = { __proto__: null}));
})(Word || (Word = {__proto__: null}));

var Word;
(function (Word) {
	var Interfaces;
	(function (Interfaces) {
		var RangeUpdateData = (function() {
			function RangeUpdateData() {
				/// <summary>An interface for updating data on the Range object, for use in "range.set({ ... })".</summary>
				/// <field name="font" type="Word.Interfaces.FontUpdateData">Gets the text format of the range. Use this to get and set font name, size, color, and other properties. [Api set: WordApi 1.1]</field>
				/// <field name="hyperlink" type="String">Gets the first hyperlink in the range, or sets a hyperlink on the range. All hyperlinks in the range are deleted when you set a new hyperlink on the range. Use a &apos;#&apos; to separate the address part from the optional location part. [Api set: WordApi 1.3]</field>;
				/// <field name="style" type="String">Gets or sets the style name for the range. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the &quot;styleBuiltIn&quot; property. [Api set: WordApi 1.1]</field>;
				/// <field name="styleBuiltIn" type="String">Gets or sets the built-in style name for the range. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the &quot;style&quot; property. [Api set: WordApi 1.3]</field>;
			}
			return RangeUpdateData;
		})();
		Interfaces.RangeUpdateData.__proto__ = null;
		Interfaces.RangeUpdateData = RangeUpdateData;
	})(Interfaces = Word.Interfaces || (Word.Interfaces = { __proto__: null}));
})(Word || (Word = {__proto__: null}));

var Word;
(function (Word) {
	var Interfaces;
	(function (Interfaces) {
		var SearchOptionsUpdateData = (function() {
			function SearchOptionsUpdateData() {
				/// <summary>An interface for updating data on the SearchOptions object, for use in "searchOptions.set({ ... })".</summary>
				/// <field name="ignorePunct" type="Boolean">Gets or sets a value that indicates whether to ignore all punctuation characters between words. Corresponds to the Ignore punctuation check box in the Find and Replace dialog box. [Api set: WordApi 1.1]</field>;
				/// <field name="ignoreSpace" type="Boolean">Gets or sets a value that indicates whether to ignore all whitespace between words. Corresponds to the Ignore whitespace characters check box in the Find and Replace dialog box. [Api set: WordApi 1.1]</field>;
				/// <field name="matchCase" type="Boolean">Gets or sets a value that indicates whether to perform a case sensitive search. Corresponds to the Match case check box in the Find and Replace dialog box. [Api set: WordApi 1.1]</field>;
				/// <field name="matchPrefix" type="Boolean">Gets or sets a value that indicates whether to match words that begin with the search string. Corresponds to the Match prefix check box in the Find and Replace dialog box. [Api set: WordApi 1.1]</field>;
				/// <field name="matchSuffix" type="Boolean">Gets or sets a value that indicates whether to match words that end with the search string. Corresponds to the Match suffix check box in the Find and Replace dialog box. [Api set: WordApi 1.1]</field>;
				/// <field name="matchWholeWord" type="Boolean">Gets or sets a value that indicates whether to find operation only entire words, not text that is part of a larger word. Corresponds to the Find whole words only check box in the Find and Replace dialog box. [Api set: WordApi 1.1]</field>;
				/// <field name="matchWildcards" type="Boolean">Gets or sets a value that indicates whether the search will be performed using special search operators. Corresponds to the Use wildcards check box in the Find and Replace dialog box. [Api set: WordApi 1.1]</field>;
			}
			return SearchOptionsUpdateData;
		})();
		Interfaces.SearchOptionsUpdateData.__proto__ = null;
		Interfaces.SearchOptionsUpdateData = SearchOptionsUpdateData;
	})(Interfaces = Word.Interfaces || (Word.Interfaces = { __proto__: null}));
})(Word || (Word = {__proto__: null}));

var Word;
(function (Word) {
	var Interfaces;
	(function (Interfaces) {
		var SectionUpdateData = (function() {
			function SectionUpdateData() {
				/// <summary>An interface for updating data on the Section object, for use in "section.set({ ... })".</summary>
				/// <field name="body" type="Word.Interfaces.BodyUpdateData">Gets the body object of the section. This does not include the header/footer and other section metadata. [Api set: WordApi 1.1]</field>
			}
			return SectionUpdateData;
		})();
		Interfaces.SectionUpdateData.__proto__ = null;
		Interfaces.SectionUpdateData = SectionUpdateData;
	})(Interfaces = Word.Interfaces || (Word.Interfaces = { __proto__: null}));
})(Word || (Word = {__proto__: null}));

var Word;
(function (Word) {
	var Interfaces;
	(function (Interfaces) {
		var SettingUpdateData = (function() {
			function SettingUpdateData() {
				/// <summary>An interface for updating data on the Setting object, for use in "setting.set({ ... })".</summary>
				/// <field name="value" >Gets or sets the value of the setting. [Api set: WordApi 1.4]</field>;
			}
			return SettingUpdateData;
		})();
		Interfaces.SettingUpdateData.__proto__ = null;
		Interfaces.SettingUpdateData = SettingUpdateData;
	})(Interfaces = Word.Interfaces || (Word.Interfaces = { __proto__: null}));
})(Word || (Word = {__proto__: null}));

var Word;
(function (Word) {
	var Interfaces;
	(function (Interfaces) {
		var TableUpdateData = (function() {
			function TableUpdateData() {
				/// <summary>An interface for updating data on the Table object, for use in "table.set({ ... })".</summary>
				/// <field name="font" type="Word.Interfaces.FontUpdateData">Gets the font. Use this to get and set font name, size, color, and other properties. [Api set: WordApi 1.3]</field>
				/// <field name="alignment" type="String">Gets or sets the alignment of the table against the page column. The value can be &apos;Left&apos;, &apos;Centered&apos;, or &apos;Right&apos;. [Api set: WordApi 1.3]</field>;
				/// <field name="headerRowCount" type="Number">Gets and sets the number of header rows. [Api set: WordApi 1.3]</field>;
				/// <field name="horizontalAlignment" type="String">Gets and sets the horizontal alignment of every cell in the table. The value can be &apos;Left&apos;, &apos;Centered&apos;, &apos;Right&apos;, or &apos;Justified&apos;. [Api set: WordApi 1.3]</field>;
				/// <field name="shadingColor" type="String">Gets and sets the shading color. Color is specified in &quot;#RRGGBB&quot; format or by using the color name. [Api set: WordApi 1.3]</field>;
				/// <field name="style" type="String">Gets or sets the style name for the table. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the &quot;styleBuiltIn&quot; property. [Api set: WordApi 1.3]</field>;
				/// <field name="styleBandedColumns" type="Boolean">Gets and sets whether the table has banded columns. [Api set: WordApi 1.3]</field>;
				/// <field name="styleBandedRows" type="Boolean">Gets and sets whether the table has banded rows. [Api set: WordApi 1.3]</field>;
				/// <field name="styleBuiltIn" type="String">Gets or sets the built-in style name for the table. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the &quot;style&quot; property. [Api set: WordApi 1.3]</field>;
				/// <field name="styleFirstColumn" type="Boolean">Gets and sets whether the table has a first column with a special style. [Api set: WordApi 1.3]</field>;
				/// <field name="styleLastColumn" type="Boolean">Gets and sets whether the table has a last column with a special style. [Api set: WordApi 1.3]</field>;
				/// <field name="styleTotalRow" type="Boolean">Gets and sets whether the table has a total (last) row with a special style. [Api set: WordApi 1.3]</field>;
				/// <field name="values" type="Array" elementType="Array">Gets and sets the text values in the table, as a 2D Javascript array. [Api set: WordApi 1.3]</field>;
				/// <field name="verticalAlignment" type="String">Gets and sets the vertical alignment of every cell in the table. The value can be &apos;Top&apos;, &apos;Center&apos;, or &apos;Bottom&apos;. [Api set: WordApi 1.3]</field>;
				/// <field name="width" type="Number">Gets and sets the width of the table in points. [Api set: WordApi 1.3]</field>;
			}
			return TableUpdateData;
		})();
		Interfaces.TableUpdateData.__proto__ = null;
		Interfaces.TableUpdateData = TableUpdateData;
	})(Interfaces = Word.Interfaces || (Word.Interfaces = { __proto__: null}));
})(Word || (Word = {__proto__: null}));

var Word;
(function (Word) {
	var Interfaces;
	(function (Interfaces) {
		var TableRowUpdateData = (function() {
			function TableRowUpdateData() {
				/// <summary>An interface for updating data on the TableRow object, for use in "tableRow.set({ ... })".</summary>
				/// <field name="font" type="Word.Interfaces.FontUpdateData">Gets the font. Use this to get and set font name, size, color, and other properties. [Api set: WordApi 1.3]</field>
				/// <field name="horizontalAlignment" type="String">Gets and sets the horizontal alignment of every cell in the row. The value can be &apos;Left&apos;, &apos;Centered&apos;, &apos;Right&apos;, or &apos;Justified&apos;. [Api set: WordApi 1.3]</field>;
				/// <field name="preferredHeight" type="Number">Gets and sets the preferred height of the row in points. [Api set: WordApi 1.3]</field>;
				/// <field name="shadingColor" type="String">Gets and sets the shading color. Color is specified in &quot;#RRGGBB&quot; format or by using the color name. [Api set: WordApi 1.3]</field>;
				/// <field name="values" type="Array" elementType="Array">Gets and sets the text values in the row, as a 2D Javascript array. [Api set: WordApi 1.3]</field>;
				/// <field name="verticalAlignment" type="String">Gets and sets the vertical alignment of the cells in the row. The value can be &apos;Top&apos;, &apos;Center&apos;, or &apos;Bottom&apos;. [Api set: WordApi 1.3]</field>;
			}
			return TableRowUpdateData;
		})();
		Interfaces.TableRowUpdateData.__proto__ = null;
		Interfaces.TableRowUpdateData = TableRowUpdateData;
	})(Interfaces = Word.Interfaces || (Word.Interfaces = { __proto__: null}));
})(Word || (Word = {__proto__: null}));

var Word;
(function (Word) {
	var Interfaces;
	(function (Interfaces) {
		var TableCellUpdateData = (function() {
			function TableCellUpdateData() {
				/// <summary>An interface for updating data on the TableCell object, for use in "tableCell.set({ ... })".</summary>
				/// <field name="body" type="Word.Interfaces.BodyUpdateData">Gets the body object of the cell. [Api set: WordApi 1.3]</field>
				/// <field name="columnWidth" type="Number">Gets and sets the width of the cell&apos;s column in points. This is applicable to uniform tables. [Api set: WordApi 1.3]</field>;
				/// <field name="horizontalAlignment" type="String">Gets and sets the horizontal alignment of the cell. The value can be &apos;Left&apos;, &apos;Centered&apos;, &apos;Right&apos;, or &apos;Justified&apos;. [Api set: WordApi 1.3]</field>;
				/// <field name="shadingColor" type="String">Gets or sets the shading color of the cell. Color is specified in &quot;#RRGGBB&quot; format or by using the color name. [Api set: WordApi 1.3]</field>;
				/// <field name="value" type="String">Gets and sets the text of the cell. [Api set: WordApi 1.3]</field>;
				/// <field name="verticalAlignment" type="String">Gets and sets the vertical alignment of the cell. The value can be &apos;Top&apos;, &apos;Center&apos;, or &apos;Bottom&apos;. [Api set: WordApi 1.3]</field>;
			}
			return TableCellUpdateData;
		})();
		Interfaces.TableCellUpdateData.__proto__ = null;
		Interfaces.TableCellUpdateData = TableCellUpdateData;
	})(Interfaces = Word.Interfaces || (Word.Interfaces = { __proto__: null}));
})(Word || (Word = {__proto__: null}));

var Word;
(function (Word) {
	var Interfaces;
	(function (Interfaces) {
		var TableBorderUpdateData = (function() {
			function TableBorderUpdateData() {
				/// <summary>An interface for updating data on the TableBorder object, for use in "tableBorder.set({ ... })".</summary>
				/// <field name="color" type="String">Gets or sets the table border color. [Api set: WordApi 1.3]</field>;
				/// <field name="type" type="String">Gets or sets the type of the table border. [Api set: WordApi 1.3]</field>;
				/// <field name="width" type="Number">Gets or sets the width, in points, of the table border. Not applicable to table border types that have fixed widths. [Api set: WordApi 1.3]</field>;
			}
			return TableBorderUpdateData;
		})();
		Interfaces.TableBorderUpdateData.__proto__ = null;
		Interfaces.TableBorderUpdateData = TableBorderUpdateData;
	})(Interfaces = Word.Interfaces || (Word.Interfaces = { __proto__: null}));
})(Word || (Word = {__proto__: null}));
var Word;
(function (Word) {
	var RequestContext = (function (_super) {
		__extends(RequestContext, _super);
		function RequestContext() {
			/// <summary>
			/// The RequestContext object facilitates requests to the Word application. Since the Office add-in and the Word application run in two different processes, the request context is required to get access to the Word object model from the add-in.
			/// </summary>
				/// <field name="document" type="Word.Document">Root object for interacting with the document</field>
			_super.call(this, null);
		}
		return RequestContext;
	})(OfficeExtension.ClientRequestContext);
	Word.RequestContext = RequestContext;

	Word.run = function (batch) {
		/// <signature>
		/// <summary>
		/// Executes a batch script that performs actions on the Word object model, using a new RequestContext. When the promise is resolved, any tracked objects that were automatically allocated during execution will be released.
		/// </summary>
		/// <param name="batch" type="function(context) { ... }">
		/// A function that takes in a RequestContext and returns a promise (typically, just the result of "context.sync()").
		/// <br />
		/// The context parameter facilitates requests to the Word application. Since the Office add-in and the Word application run in two different processes, the RequestContext is required to get access to the Word object model from the add-in.
		/// </param>
		/// </signature>
		/// <signature>
		/// <summary>
		/// Executes a batch script that performs actions on the Word object model, using the RequestContext of a previously-created API object. When the promise is resolved, any tracked objects that were automatically allocated during execution will be released.
		/// </summary>
		/// <param name="object" type="OfficeExtension.ClientObject">
		/// A previously-created API object. The batch will use the same RequestContext as the passed-in object, which means that any changes applied to the object will be picked up by "context.sync()".
		/// </param>
		/// <param name="batch" type="function(context) { ... }">
		/// A function that takes in a RequestContext and returns a promise (typically, just the result of "context.sync()").
		/// <br />
		/// The context parameter facilitates requests to the Word application. Since the Office add-in and the Word application run in two different processes, the RequestContext is required to get access to the Word object model from the add-in.
		/// </param>
		/// </signature>
		/// <signature>
		/// <summary>
		/// Executes a batch script that performs actions on the Word object model, using the RequestContext of a previously-created API object. When the promise is resolved, any tracked objects that were automatically allocated during execution will be released.
		/// </summary>
		/// <param name="objects" type="Array&lt;OfficeExtension.ClientObject&gt;">
		/// An array of previously-created API objects. The array will be validated to make sure that all of the objects share the same context. The batch will use this shared RequestContext, which means that any changes applied to these objects will be picked up by "context.sync()".
		/// </param>
		/// <param name="batch" type="function(context) { ... }">
		/// A function that takes in a RequestContext and returns a promise (typically, just the result of "context.sync()").
		/// <br />
		/// The context parameter facilitates requests to the Word application. Since the Office add-in and the Word application run in two different processes, the RequestContext is required to get access to the Word object model from the add-in.
		/// </param>
		/// </signature>
		arguments[arguments.length - 1](new Word.RequestContext());
		return new OfficeExtension.Promise();
	}
})(Word || (Word = {__proto__: null}));
Word.__proto__ = null;

