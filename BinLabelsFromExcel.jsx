/* SET-UP */
alert("IMPORTANT:\nThis script uses InDesign to collect information from an Excel document. This requires opening Excel through InDeign. Please save and close all open Excel files before continuing as InDesign may close any open file without saving.\n\nAlso, because InDesign and Excel are not perfectly compatible, this script may error when it first tries to connect to Excel. If that happens, simply run the script again, and the connection should be successful.");
var excelFile = File.openDialog("Select excel file");
//var LabelFile = File.openDialog("Select Bin Label file");
var myDoc = app.activeDocument;
var myPath = myDoc.filePath;
var myData;

/* END SET-UP */
var t1 = excelFile.fsName;
var t2 = t1.split('/');
t2.shift();
t1 = t2.join(':');


myData = getExcelData();

//$.writeln(myData.length);
var fabric, piece, stature, size, sku, quantity, measurement, price, image, master;
var myRow, tempPage, tempMaster, tempItems, tempFrame, tempColor;
var docPages = myDoc.pages;
for (var i = 1; i < myData.length; i++) {
	myRow = myData[i];
	if (myData[i][0] != '') {
		fabric = myData[i][0];
	}
	if (myData[i][1] != '') {
		piece = myData[i][1];
	}
	if (myData[i][2] != '') {
		stature = myData[i][2];
	}
	size = myData[i][3];
	sku = myData[i][4].toString() + myData[i][5].toString();
	quantity = parseInt(myData[i][6]);
	measurement = myData[i][7];
	if (myData[i][8] != '') {
		price = myData[i][8];
	}
	if (myData[i][9] != '') {
		image = myData[i][9];
	}
	if (myData[i][10] != '') {
		master = myData[i][10];
	}
	
	
	
	for (var p = quantity; p > 0; p--) {
		tempMaster = getMasterPage(master);
		tempColor = getSwatch(master);
		tempPage = docPages.add(LocationOptions.AT_END);
		tempPage.appliedMaster = tempMaster;
		tempItems = tempPage.masterPageItems;
		for (var piCount = 0; piCount < tempItems.length; piCount++) {
			switch (tempItems[piCount].label) {
				case 'fabric':
				tempFrame = tempItems[piCount].override(tempPage);
				tempFrame.contents = fabric;
				tempFrame.paragraphs[0].fillColor = tempColor;
				break;
				case 'piece':
				tempFrame = tempItems[piCount].override(tempPage);
				tempFrame.contents = piece;
				break;
				case 'size':
				tempFrame = tempItems[piCount].override(tempPage);
				tempFrame.contents = size;
				tempFrame.paragraphs[0].fillColor = tempColor;
				break;
				case 'stature':
				tempFrame = tempItems[piCount].override(tempPage);
				tempFrame.contents = stature;
				tempFrame.paragraphs[0].fillColor = tempColor;
				break;
				case 'measurement':
				tempFrame = tempItems[piCount].override(tempPage);
				tempFrame.contents = measurement;
				break;
				case 'price':
					tempFrame = tempItems[piCount].override(tempPage);
					try{
						tempFrame.contents = price;
						break;
					}
					catch(e){
						tempFrame.contents = '';
						break;
					}
				case 'sku':
				tempFrame = tempItems[piCount].override(tempPage);
				tempFrame.contents = sku;
				break;
				case 'lineArt':
				tempFrame = tempItems[piCount].override(tempPage);
				tempFrame.place(File(image));
				tempFrame.fit(FitOptions.PROPORTIONALLY);
				break;
			}
		}
	}
}

if (myDoc.pages.length > 1) {
	myDoc.pages[0].remove();
}


var myFolder = new Folder(myPath + '/NEW/');
if (! myFolder.exists)
myFolder.create();
var myDate = new Date();
var myStamp = getDateFormat(myDate);
var presets = app.pdfExportPresets;
app.pdfExportPreferences.pageRange = PageRange.ALL_PAGES;
var myPreset = presets.itemByName('[LDS Print Standard]');
//var myPreset = presets.itemByName('[LDS Web Standard]');
myDoc.exportFile(ExportFormat.PDF_TYPE, File(myFolder + '/BinLabel' + myStamp), false, myPreset);

alert('DONE');


/* FUNCTIONS */
function getExcelData() {
	// The platform-specific full path name for the xlsx-file -- fsName
	// If you pass it as a string, make sure to double the backslashes in the path like in the line below
	var excelFilePath = t1;
	
	// [Optional] the character to use for splitting the rows in the spreadsheed.
	// If it isn't set, pipe (|) will be used by default
	var splitCharRows = "|";
	
	// [Optional] the character to use for splitting the columns in the spreadsheed: e.g. semicolon (;) or tab (\t)
	// If it isn't set, semicolon will be used by default
	var splitCharColumns = ";";
	
	// [Optional] the worksheet number: either string or number. If it isn't set, the first worksheet will be used by default
	var sheetNumber = "1";
	
	// Returns an array in case of success; null -- if something went wrong: e.g. called on PC, too old version of InDesign.
	var data = GetDataFromExcelMac(excelFilePath, splitCharRows, splitCharColumns, sheetNumber);
	return data;
}

function GetDataFromExcelMac(excelFilePath, splitCharRows, splitCharColumns, sheetNumber) {
	if (File.fs != "Macintosh") return null;
	if (typeof splitCharRows === "undefined") var splitCharRows = "|";
	if (typeof splitCharColumns === "undefined") var splitCharColumns = ";";
	if (typeof sheetNumber === "undefined") var sheetNumber = "1";
	
	var appVersion,
	appVersionNum = Number(String(app.version).split(".")[0]);
	
	switch (appVersionNum) {
		case 15:
		appVersion = "CC 2020";
		break;
		case 14:
		appVersion = "CC 2019";
		break;
		case 13:
		appVersion = "CC 2018";
		break;
		case 12:
		appVersion = "CC 2017";
		break;
		case 11:
		appVersion = "CC 2015";
		break;
		case 10:
		appVersion = "CC 2014";
		break;
		case 9:
		appVersion = "CC";
		break;
		case 8:
		appVersion = "CS 6";
		break;
		case 7:
		if (app.version.match(/^7\.5/) != null) {
			appVersion = "CS 5.5";
		} else {
			appVersion = "CS 5";
		}
		break;
		case 6:
		appVersion = "CS 4";
		break;
		case 5:
		appVersion = "CS 3";
		break;
		case 4:
		appVersion = "CS 2";
		break;
		case 3:
		appVersion = "CS";
		break;
		default:
		return null;
	}
	
	var as = 'tell application "Microsoft Excel"\r';
	as += 'open file \"' + excelFilePath + '\"\r';
	as += 'set theWorkbook to active workbook\r';
	as += 'set theSheet to sheet ' + sheetNumber + ' of theWorkbook\r';
	as += 'set theMatrix to value of used range of theSheet\r';
	as += 'set theRowCount to count theMatrix\r';
	as += 'set str to ""\r';
	as += 'set oldDelimiters to AppleScript\'s text item delimiters\r';
	as += 'repeat with countRows from 1 to theRowCount\r';
	as += 'set theRow to item countRows of theMatrix\r';
	as += 'set AppleScript\'s text item delimiters to \"' + splitCharColumns + '\"\r';
	as += 'set str to str & (theRow as string) & \"' + splitCharRows + '\"\r';
	as += 'end repeat\r';
	as += 'set AppleScript\'s text item delimiters to oldDelimiters\r';
	as += 'close theWorkbook saving no\r';
	as += 'end tell\r';
	as += 'tell application id "com.adobe.InDesign"\r';
	as += 'tell script args\r';
	as += 'set value name "excelData" value str\r';
	as += 'end tell\r';
	as += 'end tell';
	
	if (appVersionNum > 5) {
		// CS4 and above
		app.doScript(as, ScriptLanguage.APPLESCRIPT_LANGUAGE, undefined, UndoModes.ENTIRE_SCRIPT);
	} else {
		// CS3 and below
		app.doScript(as, ScriptLanguage.APPLESCRIPT_LANGUAGE);
	}
	
	var str = app.scriptArgs.getValue("excelData");
	app.scriptArgs.clear();
	
	var tempArrLine, line,
	data =[],
	tempArrData = str.split(splitCharRows);
	
	for (var i = 0; i < tempArrData.length; i++) {
		line = tempArrData[i];
		if (line == "") continue;
		tempArrLine = line.split(splitCharColumns);
		data.push(tempArrLine);
	}
	return data;
}

function getMasterPage(m) {
	for (var mCount = 0; mCount < myDoc.masterSpreads.length; mCount++) {
		if (myDoc.masterSpreads[mCount].baseName == m) {
			return myDoc.masterSpreads[mCount];
		}
	}
	return myDoc.masterSpreads[0];
}

function getSwatch(c) {
	return myDoc.swatches.itemByName(c);
}

function getDateFormat(d){
	var yr, mon, day, hrs, min, sec, ms;
	yr = d.getYear();
	yr1 = yr + 1900;
	mon = d.getMonth();
	mon1 = mon + 1;
	day = d.getDate().toString();
	hrs = d.getHours().toString();
    min = d.getMinutes().toString();
    sec = d.getSeconds().toString();
    ms = d.getMilliseconds().toString();

    return '_' + yr1 + '_' + mon1 + '_' + day + '_' + hrs + min; 
}
/* END FUNCTIONS */