///<overview>
///This file contains the javascript code to call the binding JSOM API.
///most of the functions read input data from textareas (input1, input2... input6) in html file 
///and write output data into another textarea (output)
///</overview>

/// <summary>
/// om object, access point to jsom api.
/// </summary>
var _OM;

/// <summary>
/// Settings object
/// </summary>
var _settings;

/// <summary>
/// the recent binding object.
/// </summary>
var _bindingObj;

/// <summary>
/// the type of binding, could be 
/// Microsoft.Office.WebExtension.BindingType.Text
/// Microsoft.Office.WebExtension.BindingType.Matrix
/// Microsoft.Office.WebExtension.BindingType.Table
/// </summary>
var _bindingType;

/// <summary>
/// the type of data, could be 
/// Microsoft.Office.WebExtension.CoercionType.Text
/// Microsoft.Office.WebExtension.CoercionType.Matrix
/// Microsoft.Office.WebExtension.CoercionType.Table
/// </summary>
var _coercionType;

/// <summary>
/// the type of data, could be 
/// Microsoft.Office.WebExtension.ValueFormat.Formatted
/// Microsoft.Office.WebExtension.ValueFormat.Unformatted
/// </summary>
var _valueFormat;

/// <summary>
/// the type of file, could be 
/// Microsoft.Office.FileType.Text
/// Microsoft.Office.FileType.Compressed
/// Microsoft.Office.FileType.PDF
/// </summary>
var _fileType;

/// <summary>
/// other parameters. const value
/// </summary>
var _filterType;
var _datachangedEvent;

/// <summary>
/// reason value, could be
/// inserted
/// documentOpened
/// </summary>
var _reason;

/// <summary>
/// variable to store binding promise object
/// </summary>
var _offSelect;

/// <summary>
/// variable to store binding value for binding promise
/// </summary>
var _bindSelectValue = "bindings#";

/// <summary>
/// variable to store binding promise id 
/// </summary>
var _bindSelectValueLocal;

/// <summary>
/// variable to store setting test data 
/// </summary>
var _setting;

/// <summary>
/// handle to dialog 
/// </summary>
var _dialog

/// <summary>
/// heightValue of dialog
/// </summary>
var _dialogHeight


/// <summary>
/// widthValue of dialog
/// </summary>
var _dialogWidth

// Init setting test data
function initSettingTestData() {
	var ob = new Object();
	ob.name = "object name";
	ob.value = "object value";
	ob.age = 15;
	ob.money = 300.56;
	ob.gf = null;
	ob.friends = [1, 2, 3, 4, 5];
	ob.isBoy = true;
	
	_setting = {
	names: ["TEST_SETTING_KEY", "1", "a", " ", "123.321", "5", "a test string.", "7", "8", "9", "10", "11", "12", "13"],
	values: ["TEST_SETTING_VALUE", "a test string", 123, -123, 0, 0.0123,
			 0.0000000000000000000000000000000005, 99999999999999999999999999999999,
			 null, true, false, [1, 2, 3, 4, 5], [['a', 'b'], [1, 2]], ob]
	};
}

function getParameterByName(name, url) {
	if (!url) url = window.location.href;
	name = name.replace(/[\[\]]/g, "\\$&");
	var regex = new RegExp("[?&]" + name + "(=([^&#]*)|&|#|$)"),
	results = regex.exec(url);
	if (!results) return null;
	if (!results[2]) return '';
	return decodeURIComponent(results[2].replace(/\+/g, " "));
}

Office.initialize = function (reason) {
	var dialogMessage = getParameterByName("dialogMessage"); 
	if (dialogMessage) {
		Office.context.ui.messageParent(dialogMessage);
	}
	
	if (getParameterByName("dialogClose") == "1") {
		setTimeout(function() {
				   window.close();
				   }, 2000);
	}
	
	_coercionType = Microsoft.Office.WebExtension.CoercionType.Text;
	_valueFormat = Microsoft.Office.WebExtension.ValueFormat.Unformatted;
	_filterType = Microsoft.Office.WebExtension.FilterType.All;
	_datachange = Microsoft.Office.WebExtension.EventType.BindingDataChanged;
	_OM = Office.context.document;
	_settings = Office.context.document.settings;
	_reason = reason;
	initSettingTestData();
	returnValueToAutomation('initialized');
	window.osfTestAgaveStatus = "initialized";
	passStep();
}

function returnValueToAutomation_backup(res) {
	if (res instanceof Microsoft.Office.WebExtension.TableData) {
		var rows = res.rows;
		var headers = res.headers
		document.getElementById("output").value = "rows: " + rows + "||headers: " + headers;
	} else {
		document.getElementById("output").value = res;
	}
}

function returnValueToAutomation(res) {
	var output;
	if (res instanceof Microsoft.Office.WebExtension.TableData) {
		output = StringifyTable(res);
	}
	else if (res instanceof Array) {
		output = StringifyMatrix(res);
	}
	else {
		output = res;
	}
	document.getElementById("output").value = output;
}

function allInputData(){
	var _input1 = document.getElementById("input1").value;
	var _input2 = document.getElementById("input2").value;
	var _input3 = document.getElementById("input3").value;
	var _input4 = document.getElementById("input4").value;
	var _input5 = document.getElementById("input5").value;
	var _input6 = document.getElementById("input6").value;
	
	return [_input1, _input2, _input3, _input4, _input5, _input6];
}

function inputData_backup() {
	var _obj;
	var _textData = document.getElementById('textData');
	var _matrixData = document.getElementById('matrixData');
	var _tableData = document.getElementById('tableData');
	var _imageData = document.getElementById('imageData');
	var _input1 = document.getElementById("input1");
	var _input2 = document.getElementById("input2");
	var _input3 = document.getElementById("input3");
	var _input4 = document.getElementById("input4");
	var _input5 = document.getElementById("input5");
	var _input6 = document.getElementById("input6");
	if (_textData.checked == 1 || _imageData.checked == 1) {
		_obj = _input1.value;
	} else if (_matrixData.checked == 1) {
		_obj = [[_input1.value, _input2.value], [_input3.value, _input4.value]];  //to test matrix 
	} else if (_tableData.checked == 1) {
		var rows = [[_input1.value, _input2.value], [_input3.value, _input4.value]];
		var headers = [_input5.value, _input6.value];
		_obj = new Microsoft.Office.WebExtension.TableData(rows, headers); //to test table
	} else {
		return null;
	}
	return _obj;
}

function inputData() {
	var _textData = document.getElementById('textData');
	var _matrixData = document.getElementById('matrixData');
	var _tableData = document.getElementById('tableData');
	var _imageData = document.getElementById('imageData');
	var _input1 = document.getElementById("input1");
	var input = _input1.value.split("|");
	if (_textData.checked == 1 || _imageData.checked == 1) {
		_obj = input[0];
	} else if (_matrixData.checked == 1) {
		_obj = eval(input[0]);  //to test matrix 
	} else if (_tableData.checked == 1) {
		var rows = eval(input[1]);
		var headers = eval(input[0]);
		_obj = new Microsoft.Office.WebExtension.TableData(rows, headers); //to test table
	} else {
		return null;
	}
	return _obj;
}

// Converts a matrix (2D array) into a string representation (which looks exactly like // a 2D array would appear in code: [[F,G],[H,I]]
function StringifyMatrix(matrix) {
	if (matrix == null) {
		return "null";
	}
	if (matrix.length == 0) {
		return "[]";
	}
	
	var val = "[";
	for (var i = 0; i < matrix.length; i++) {
		if (matrix[i] == null) {
			val += "null,";
		} else if (matrix[i].length == 0) {
			val += "[],";
		} else {
			val += "['" + matrix[i].join("','") + "'],";
		}
	}
	return val.substring(0, val.length - 1) + "]";
}

// Converts a table object into a string representation, with the headers first and body
// second, separated by a vertical bar.  Like this:
// [["header1","header2"]]|[["body1","body2"]]
function StringifyTable(table) {
	if (table == null) {
		return "null";
	}
	var headers = StringifyMatrix(table.headers);
	var rows = StringifyMatrix(table.rows);
	return headers + "|" + rows;
}

// Evaluate JavaScript
function executeJavaScript() {
	var item = document.getElementById("codeWindow");
	var output = document.getElementById("output");
	output.value = '';
	output.value = eval(item.value);
}

// Output
function output(str) {
	var textArea = document.getElementById("output");
	textArea.value = textArea.value + str;
}

function passStep() {
	var textArea = document.getElementById("outputStepResult");
	textArea.value = "PASS";
}

function failStep() {
	var textArea = document.getElementById("outputStepResult");
	textArea.value = "FAIL";
}

function failStepWithErrMsg(errMsg) {
	var textArea = document.getElementById("outputStepResult");
	textArea.value = "FAIL|" + errMsg;
}

function outputStepResult(asyncResult) {
	if (asyncResult.status == "failed") {
		failStepWithErrMsg('Error: ' + asyncResult.error.name + " : " + asyncResult.error.message);
	}
	else {
		passStep();
	}
}

function cleanOutputTextBox() {
	document.getElementById("outputStepResult").value = "";
	document.getElementById("output").value = "";
	document.getElementById("output1").value = "";
	document.getElementById("output2").value = "";
	document.getElementById("output3").value = "";
	document.getElementById("output4").value = "";
	document.getElementById("output5").value = "";
}

var PASS_KEYWORD = "Done all tests";

function executeOpenDialogFailTests() {
	var dialogUrl = window.location.protocol + "//" + window.location.host + window.location.pathname;
	dialogUrl = dialogUrl.replace("JSOMTestDialog.html", "JSOMTestChildDialog.html")
	Office.context.ui.displayDialogAsync(dialogUrl, {height: 80, width:80}, function(asyncResult){
		if(asyncResult.status == Office.AsyncResultStatus.Succeeded)
		{
			_dialog = asyncResult.value;
			Office.context.ui.displayDialogAsync(dialogUrl, {height: 80, width:80}, function(asyncResult){
				if(asyncResult.status !== Office.AsyncResultStatus.Succeeded)
				{
					output(PASS_KEYWORD);
					_dialog.close();
					_dialog = null;
				}
    		});
		}
    });
}

function executeOpenDialogTests() {
	var dialogUrl = window.location.protocol + "//" + window.location.host + window.location.pathname;
	dialogUrl = dialogUrl.replace("JSOMTestDialog.html", "JSOMTestChildDialog.html")
	Office.context.ui.displayDialogAsync(dialogUrl, {height: 80, width:80}, function(asyncResult){
		if(asyncResult.status == Office.AsyncResultStatus.Succeeded)
		{
			_dialog = asyncResult.value;
			output(PASS_KEYWORD);
			_dialog.addEventHandler(Microsoft.Office.WebExtension.EventType.DialogMessageReceived, function (arg) {
				var winParas = arg.message.split("+");
				_dialogHeight = parseFloat(winParas[0]);
    			_dialogWidth = parseFloat(winParas[1]);
				_dialog.close();
				_dialog = null;
			});
		}
    });
}

function executeOpenSmallerDialogTests() {
	var dialogUrl = window.location.protocol + "//" + window.location.host + window.location.pathname;
	dialogUrl = dialogUrl.replace("JSOMTestDialog.html", "JSOMTestChildDialog.html")
	Office.context.ui.displayDialogAsync(dialogUrl, {height: 40, width:80}, function(asyncResult){
		if(asyncResult.status !== Office.AsyncResultStatus.Succeeded)
		{
			_dialog = asyncResult.value;
			_dialog.addEventHandler(Microsoft.Office.WebExtension.EventType.DialogMessageReceived, function (arg) {
				var winParas = arg.message.split("+");
				height = parseFloat(winParas[0]);
    			width = parseFloat(winParas[1]);
    			if(height > (0.49 * _dialogHeight) && height < (0.51 * _dialogHeight) && width > (0.99 * _dialogWidth) && width < (1.01 * _dialogWidth))
    			{
					output(PASS_KEYWORD);
				}
				_dialog.close();
				_dialog = null;
			});
		}
    });
}

function executeOpenDialogTwoWayMessageTests() {
	var dialogUrl = window.location.protocol + "//" + window.location.host + window.location.pathname;
	dialogUrl = dialogUrl.replace("JSOMTestDialog.html", "JSOMTestChildDialog.html")
	Office.context.ui.displayDialogAsync(dialogUrl, {height: 80, width:80}, function(asyncResult){
		if(asyncResult.status == Office.AsyncResultStatus.Succeeded)
		{
			_dialog = asyncResult.value;
			_dialog.addEventHandler(Microsoft.Office.WebExtension.EventType.DialogMessageReceived, function (arg) {
				if (arg.message === PASS_KEYWORD) {
					_dialog.close();
					_dialog = null;
				} else {
					var winParas = arg.message.split("+");
					_dialogHeight = parseFloat(winParas[0]);
					_dialogWidth = parseFloat(winParas[1]);
				}
			});
			_dialog.messageChild(PASS_KEYWORD);
		}
    });
}
