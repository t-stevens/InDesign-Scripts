var myDoc = app.activeDocument;
var myPath = myDoc.filePath;



//set export preset
var presets = app.pdfExportPresets;
var myPreset = presets.itemByName('[LDS Web Standard]');
//change preferences

uia = app.scriptPreferences.userInteractionLevel;
app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;

//Root XML
var myRoot = myDoc.xmlElements.item(0);

//run main script
main();

//restore preferences
app.scriptPreferences.userInteractionLevel = uia

/*FUNCTIONS*/
function main() {
	var myLang = getLanguageCode();
	var myItems = myRoot.evaluateXPathExpression("//*[@export]");
	updateLanguage(myItems, myLang);
	var myType = chooseExport().toString();
    //$.writeln(myType);
	if (myType == 'PDF') {
		exportPdfs(myItems);
	} else  if (myType == 'JPEG'){
		exportJPEGs(myItems);
	}
    else  {
		exportPngs(myItems);
	}
	alert('Done');
}
function getLanguageCode(){
	var pubTest = myRoot.evaluateXPathExpression("//*[@publicationLanguage]");
	if(pubTest.length > 0){
		return pubTest[0].xmlAttributes.itemByName('publicationLanguage').value;
	}
	var nameString = myDoc.name;
	var usSplit = nameString.split('_');
	if(usSplit[1].length == 3){
		return usSplit[1];
	}
	else{
		return 'TMP';
	}
}

function updateLanguage(xmlItems, langCode){
	var myItem, exAttr, exName, langName;
	for(var iCount = 0; iCount < xmlItems.length; iCount++){
		myItem = xmlItems[iCount];
		exAttr = myItem.xmlAttributes.itemByName('export');
		exName = exAttr.value;
		langName = exName.replace('TMP', langCode);
		exAttr.value = langName;
	}
}

function chooseExport() {
	var fileTypes =[ 'PDF', 'PNG', 'JPEG'];
	
	//set up dialog window
	myDlg = new Window('dialog', 'Drop-down List');
	//choose rows or columns
	myDlg.orientation = 'column';
	myDlg.alignment = 'right';
	//drop-down setup
	myDlg.DDgroup1 = myDlg.add('group');
	myDlg.DDgroup1.orientation = 'row';
	
	//add static text
	myDlg.DDgroup1.add('statictext', undefined, "Select file type to export");
	//add drop down list
	myDlg.DDgroup1.DD = myDlg.DDgroup1.add('dropdownlist', undefined, undefined, {
		items: fileTypes
	})
	
	//add default selection
	myDlg.DDgroup1.DD.selection = 0;
	//add button
	myDlg.closeBtn = myDlg.add('button', undefined, 'OK');
	// add button functions
	myDlg.closeBtn.onClick = function () {
		//close window
		this.parent.close();
	}
	//show window and store drop down value
	result = myDlg.show();
	return myDlg.DDgroup1.DD.selection;
}

function exportPdfs(items) {
	//create folder for PDFs
	var pdfFolder = new Folder(myPath + '/PDFs/');
	if (! pdfFolder.exists)
	pdfFolder.create();
	var myStories =[];
	var myNames =[];
	for (var i = 0; i < items.length; i++) {
		myStories.push(items[i].parentStory);
		myNames.push(items[i].xmlAttributes.itemByName("export").value);
	}
	var myPages, myfileName;
	for (var i = 0; i < myStories.length; i++) {
		myPages = getPages(myStories[i]);
		myfileName = myNames[i];
		app.pdfExportPreferences.pageRange = myPages;
		myDoc.exportFile (ExportFormat.PDF_TYPE, File(pdfFolder + "/" + myfileName + '.pdf'), false, myPreset);
	}
}

function exportPngs(items) {
	//create folder for PNGs
	var pngFolder = new Folder(myPath + '/PNGs/');
	if (! pngFolder.exists)
	pngFolder.create();
	var myStories =[];
	var myNames =[];
	for (var i = 0; i < items.length; i++) {
		myStories.push(items[i].parentStory);
		myNames.push(items[i].xmlAttributes.itemByName("export").value);
	}
	var myPages, myfileName;
	for (var i = 0; i < myStories.length; i++) {
		myPages = getPages(myStories[i]);
		myfileName = myNames[i];
		app.pngExportPreferences.pageString = myPages;
         app.pngExportPreferences.pngExportRange = PNGExportRangeEnum.EXPORT_RANGE;
		app.pngExportPreferences.pngQuality = PNGQualityEnum.HIGH;
         app.pngExportPreferences.exportResolution = 300;
		myDoc.exportFile (ExportFormat.PNG_FORMAT, File(pngFolder + "/" + myfileName + '.png'), false, myPreset);
	}
}

function exportJPEGs(items) {
	//create folder for JPEGs
	var jpegFolder = new Folder(myPath + '/JPGs/');
	if (! jpegFolder.exists)
	jpegFolder.create();
	var myStories =[];
	var myNames =[];
	for (var i = 0; i < items.length; i++) {
		myStories.push(items[i].parentStory);
		myNames.push(items[i].xmlAttributes.itemByName("export").value);
	}
	var myPages, myfileName;
	for (var i = myStories.length-1; i >= 0; i--) {
		myPages = getPages(myStories[i]);
		myfileName = myNames[i];
		app.jpegExportPreferences.pageString = myPages;
         app.jpegExportPreferences.jpegExportRange = ExportRangeOrAllPages.EXPORT_RANGE;
		app.jpegExportPreferences.jpegQuality = JPEGOptionsQuality.HIGH;
         app.jpegExportPreferences.exportResolution = 300;
		myDoc.exportFile (ExportFormat.JPG, File(jpegFolder + "/" + myfileName + '.jpg'), false);
	}
}

function getPages(s) {
	var p = s.textContainers[0].parentPage.name + "-" + s.textContainers[s.textContainers.length -1].parentPage.name;
	return p;
}