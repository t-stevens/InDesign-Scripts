var dlgResults = setSaveFreq();
var myFreq = dlgResults[0];
var flag = dlgResults[1];
var sTime = setStartTime();

//change preferences
uia = app.scriptPreferences.userInteractionLevel;
app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;
app.pdfPlacePreferences.pdfCrop = PDFCrop.CROP_TRIM;
var refPoint = app.activeWindow.transformReferencePoint;
app.activeWindow.transformReferencePoint = AnchorPoint.CENTER_ANCHOR;
//END change preferences

var myDoc = app.activeDocument;
var myRoot = myDoc.xmlElements.item(0);

var stories = sortStories(myRoot);
if (myDoc.xmlInstructions.count() == 0) {
	var imgString = "img";
	var idString = "file-id";
} else {
	var imgString = "image";
	var idString = "fileID";
}

var myFileIDElem = myRoot.evaluateXPathExpression("//*[@" + idString + "]");
var myFileID = myFileIDElem[0].xmlAttributes.itemByName(idString).value;
$.writeln(myFileID);

var tempName = myFileID.split('_');
var PDNum = tempName[0];
var langCode = tempName[1];
var docURI = myRoot.xmlElements.item(0).xmlAttributes.itemByName('uri').value;
var tempURI = docURI.split('/');
var myTitle = tempURI[2];
var myPath = myDoc.filePath + '/' + PDNum + '_' + langCode + '_' + myTitle + '.indd';

var oversize = myDoc.masterSpreads.itemByName('X-Oversized');
var body = myDoc.masterSpreads.itemByName('B-Body');

if (flag) {
	myDoc.save(File(myPath));
	main();
} else {
	alert('Execution cancelled');
}

function main() {
	for (var i = 0; i < stories.length; i++) {
		try {
			$.writeln('fileID: ' + stories[i].xmlAttributes.itemByName(idString).value + '\tStory: ' + (i + 1) + '/' + stories.length);
		}
		catch (e) {
			$.writeln('No fileID');
		}
		var f = resizeAndPlace(stories[i]);
		paragraphStarts(f);
		sizeImages(stories[i]);
		convertToOval(stories[i]);
		anchorText(stories[i]);
		remapStyles();
		clear(f);
		assignMasterPage(stories[i], f);
		adjustTableWidths(stories[i], f);
		//relinkXML();
		addPages(f);
		setMinLines(stories[i]);
		resizeAnchoredFrames(f);
		if (i != stories.length -1) {
			myDoc.pages.add(LocationOptions.AT_END);
		}
		stories[i].xmlAttributes.itemByName('newStory').remove();
		if ((myFreq != 'none') && ((i + 1) % myFreq == 0)) {
			myDoc.save(File(myPath));
			$.writeln('SAVED');
		}
	}
	setMasterPages();
	removeBlankPages();
	rotateImages();
	imgShading();
	myDoc.save(File(myPath));
	var totalTime = getTime(sTime);
	var dispTime = msToTime(totalTime);
	alert('Done\nElapsed time: ' + dispTime);
}

//restore preferences
app.scriptPreferences.userInteractionLevel = uia
app.activeWindow.transformReferencePoint = refPoint;



/*----FUNCTIONS----*/

//get stories with the @newStory  --  takes root element
function sortStories(myRoot) {
	var sortFn = function (a, b) {
		return a.id > b.id;
	}
	var stories = myRoot.evaluateXPathExpression("//bodyMatter//*[@newStory]");
	if (stories.length == 0) {
		stories = myRoot.evaluateXPathExpression("//*[@newStory]");
	}
	stories.sort(sortFn);
	return stories;
}

//resize page and place xml  --  takes xmlElement
function resizeAndPlace(s) {
	var p = myDoc.pages.lastItem();
	if (p.allPageItems.length != 0) {
		p = myDoc.pages.add(LocationOptions.AT_END);
	}
	p.appliedMaster = oversize;
	var myFrame = p.textFrames.add();
	myFrame.geometricBounds = p.bounds;
	s.placeXML(myFrame);
	return myFrame;
}

//add pages to overset stories  --  takes textFrame
function addPages(tFrame) {
	var p1, t0, t1, m1;
	t0 = tFrame
	while (tFrame.parentStory.overflows) {
		p1 = myDoc.pages.add(LocationOptions.AT_END);
		t1 = p1.textFrames.add();
		m1 = p1.marginPreferences;
		if (p1.side == PageSideOptions.RIGHT_HAND) {
			t1.geometricBounds =[p1.bounds[0] + m1.top, p1.bounds[1] + m1.left, p1.bounds[2] - m1.bottom, p1.bounds[3] - m1.right];
		} else if (p1.side == PageSideOptions.LEFT_HAND) {
			t1.geometricBounds =[p1.bounds[0] + m1.top, p1.bounds[1] + m1.right, p1.bounds[2] - m1.bottom, p1.bounds[3] - m1.left];
		} else if (p1.side == PageSideOptions.SINGLE_SIDED) {
			t1.geometricBounds =[p1.bounds[0] + m1.top, p1.bounds[1] + m1.right, p1.bounds[2] - m1.bottom, p1.bounds[3] - m1.left];
		}
		t0.nextTextFrame = t1;
		t0 = t1;
	}
}

//resize images based on @bounds  --  takes xmlElement
function sizeImages(s) {
	var imgArray = s.evaluateXPathExpression("//" + imgString + "[@bounds]");
	var mRectangle, bounds, myObjStyle, myImgName, temp1, temp2;
	for (var i = imgArray.length -1; i >= 0; i--) {
		bounds = getBounds(imgArray[i].xmlAttributes.itemByName('bounds').value.toLowerCase());
		temp1 = imgArray[i].xmlAttributes.itemByName('href').value;
		temp2 = temp1.split('/');
		myImgName = temp2[temp2.length -1];
		if ((bounds[0] != 'x') && (bounds[1] != 'x')) {
			try {
				mRectangle = imgArray[i].xmlContent.parent;
				mBounds = mRectangle.geometricBounds;
				mRectangle.geometricBounds =[mBounds[0], mBounds[1], mBounds[0] + bounds[1], mBounds[1] + bounds[0]];
			}
			catch (e) {
				$.writeln('ERROR in image: ' + myImgName);
			}
		} else if ((bounds[0] != 'x') && (bounds[1] == 'x')) {
			try {
				mRectangle = imgArray[i].xmlContent.parent;
				mBounds = mRectangle.geometricBounds;
				mRectangle.geometricBounds =[mBounds[0], mBounds[1], mBounds[2], mBounds[1] + bounds[0]];
				var testImgBounds = mRectangle.graphics[0].geometricBounds;
				var testWidth = testImgBounds[3] - testImgBounds[1];
				if(testWidth < (mRectangle.geometricBounds[3] - mRectangle.geometricBounds[1])){
					mRectangle.graphics[0].fit(FitOptions.FILL_PROPORTIONALLY);
				}
				else{
					mRectangle.graphics[0].fit(FitOptions.PROPORTIONALLY);
				}
				mRectangle.fit(FitOptions.FRAME_TO_CONTENT);
				mRectangle.graphics[0].fit(FitOptions.centerContent);
			}
			catch (e) {
				$.writeln('ERROR in image: ' + myImgName);
			}
		}
		//rotateImages(imgArray[i])
		fitImages(imgArray[i]);
	}
}

function rotateImages() {
	var allImages = myRoot.evaluateXPathExpression("//" + imgString + "[@bounds]");
	for (var imgC = 0; imgC < allImages.length; imgC++) {
		var myBounds = getBounds(allImages[imgC].xmlAttributes.itemByName('bounds').value.toLowerCase());
		var r = 90;
		
		if (myBounds[2] == 'full') {
			var test1 = allImages[imgC].xmlContent.geometricBounds;
			var testW, testH;
			testW = test1[3] - test1[1];
			testH = test1[2] - test1[0];
			if (testH >= testW) {
				r = 0;
			}
			var imgSide = allImages[imgC].xmlContent.parentPage.side;
			if ((imgSide == PageSideOptions.LEFT_HAND) && (r != 0)) {
				r = -90;
			}
			allImages[imgC].xmlContent.absoluteRotationAngle = r;
			allImages[imgC].xmlContent.fit(FitOptions.FILL_PROPORTIONALLY);
			allImages[imgC].xmlContent.parent.fit(FitOptions.FRAME_TO_CONTENT);
		}
	}
}
//get value of bounds with rotation angle  --  takes string 'w_12p3,h_10,r_45'
function getBounds(attrValue) {
	var w = 'x';
	var h = 'x';
	var r = 0;
	var temp1, temp2, temp3, temp4, temp5;
	if (attrValue.indexOf(',') != -1) {
		temp1 = attrValue.split(',');
		for (var i = 0; i < temp1.length; i++) {
			temp2 = temp1[i].split('_');
			temp3 = temp2[0];
			switch (temp3) {
				case 'w':
				w = getPicas(temp2[1]);
				break;
				case 'h':
				h = getPicas(temp2[1]);
				break;
				case 'r':
				r = temp2[1];
				break;
			}
		}
	} else {
		temp2 = attrValue.split('_');
		switch (temp2[0]) {
			case 'w':
			w = getPicas(temp2[1]);
			break;
			case 'h':
			h = getPicas(temp2[1]);
			break;
			case 'r':
			r = temp2[1];
			break;
		}
	}
	return ([w, h, r]);
}

//convert pica measurements to decimal value  --  takes string
function getPicas(v) {
	var temp, vPica, dec;
	if (v.indexOf('p') == -1) {
		return v;
	} else {
		temp = v.split('p');
		dec = temp[1] / 12;
		vPica = (temp[0] * 1) + dec;
		return vPica;
	}
}

//fit images to frame size  --  takes xmlElement
function fitImages(image) {
	var mRectangle, bounds, myObjStyle;
	mRectangle = image.xmlContent.parent;
	try {
		mRectangle.graphics[0].fit(FitOptions.PROPORTIONALLY);
		mRectangle.graphics[0].fit(FitOptions.centerContent);
	}
	catch (e) {
	}
	//apply object style
	if (image.xmlAttributes.itemByName('objStyle').isValid) {
		myObjStyle = getObjStyle(image);
		try {
			mRectangle = image.xmlContent.parent.appliedObjectStyle = myObjStyle;
		}
		catch (e) {
		}
	}
}

//get object style based on @objStyle  --  takes xmlElement
function getObjStyle(obj) {
	var n, myObjectStyle;
	n = obj.xmlAttributes.itemByName('objStyle').value;
	myObjectStyle = myDoc.objectStyles.item(n);
	try {
		var myName = myObjectStyle.name;
	}
	catch (e) {
		//The object style did not exist, so create it.
		myObjectStyle = myDoc.objectStyles.add({
			name: n
		});
	}
	return myObjectStyle;
}

//anchor text frames  --  takes xmlElement
function anchorText(s) {
	var cred = s.evaluateXPathExpression("//*[@bounds][not(self::" + imgString + ")]");
	var anchoredFrame, bounds, myObjStyle, dim, r;
	for (var i = 0; i < cred.length; i++) {
		bounds = getBounds(cred[i].xmlAttributes.itemByName('bounds').value.toLowerCase());
		if (bounds[1] == 'x') {
			dim =[bounds[0], 5];
		} else {
			dim =[bounds[0], bounds[1]];
		}
		r = bounds[2] * 1;
		anchoredFrame = cred[i].placeIntoInlineFrame(dim);
		anchoredFrame.absoluteRotationAngle = r;
		if (cred[i].xmlAttributes.itemByName('objStyle').isValid) {
			myObjStyle = getObjStyle(cred[i]);
			anchoredFrame.appliedObjectStyle = myObjStyle;
		}
		anchoredFrame.fit(FitOptions.FRAME_TO_CONTENT);
	}
}

//converts image frams based on @convertShape  --  takes xmlElement
function convertToOval(s) {
	var shapes = s.evaluateXPathExpression("//*[@convertShape]");
	
	var myGraphic, targetShape;
	for (var i = 0; i < shapes.length; i++) {
		myGraphic = shapes[i].graphics[0];
		targetShape = shapes[i].xmlAttributes.itemByName('convertShape').value;
		if (targetShape == 'OVAL') {
			try {
				myGraphic.parent.convertShape(ConvertShapeOptions.CONVERT_TO_OVAL);
			}
			catch (e) {
				try {
					myGraphic.convertShape(ConvertShapeOptions.CONVERT_TO_OVAL);
				}
				catch (e) {
					$.writeln('error in image ' + shapes[i].xmlAttributes.itemByName('href').value);
				}
			}
		}
	}
}

//assigns master pages based on @newStory  --  takes xmlElement
function assignMasterPage(x, f) {
	var xMaster = x.xmlAttributes.item('newStory').value;
	var p = f.parentPage;
	var myMaster = getMasterSpread(xMaster);
	var myMPs = myMaster.pages;
	p.appliedMaster = myMaster;
	var m1 = p.marginPreferences;
	if (p.side == PageSideOptions.RIGHT_HAND) {
		f.geometricBounds =[p.bounds[0] + m1.top, p.bounds[1] + m1.left, p.bounds[2] - m1.bottom, p.bounds[3] - m1.right];
	} else if (p.side == PageSideOptions.LEFT_HAND) {
		f.geometricBounds =[p.bounds[0] + m1.top, p.bounds[1] + m1.right, p.bounds[2] - m1.bottom, p.bounds[3] - m1.left];
	} else if (p.side == PageSideOptions.SINGLE_SIDED) {
		f.geometricBounds =[p.bounds[0] + m1.top, p.bounds[1] + m1.right, p.bounds[2] - m1.bottom, p.bounds[3] - m1.left];
	}
}

//gets master page that matches value in @newStory  --  takes string
function getMasterSpread(v) {
	var masterPages = myDoc.masterSpreads;
	myMaster = myDoc.masterSpreads.itemByName('B-Body');
	var mp;
	for (var i = 0; i < masterPages.length; i++) {
		mp = masterPages[i].namePrefix;
		if (mp == v) {
			myMaster = masterPages[i];
		}
	}
	return myMaster;
}


//maps xml tags to styles
function remapStyles() {
	myDoc.mapXMLTagsToStyles();
}

//set all paragraphs to start anywhere  --  takes textFrame
function paragraphStarts(t) {
	var myStory = t.parentStory;
	var p = myStory.paragraphs;
	for (var i = 0; i < p.length; i++) {
		p[i].startParagraph = StartParagraph.ANYWHERE;
	}
}

//clear all style overrides  --  takes textFrame
function clear(f) {
	var myStory = f.parentStory;
	var p = myStory.paragraphs;
	for (var i = 0; i < p.length; i++) {
		p[i].clearOverrides();
	}
}

//Set table width to match text frame  --  takes xmlElement
function adjustTableWidths(s, f) {
	var tables1 = s.evaluateXPathExpression("//table[not(@skip)]");
	var tables2 = s.evaluateXPathExpression("//table.fig[not(@skip)]");
	var myTables = tables1.concat(tables2);
	var t, c, w, inset;
	for (var i = 0; i < myTables.length; i++) {
		//c = myTables[i].xmlAttributes.itemByName('ORIG_tcol').value;
		t = myTables[i].tables[0];
		c = t.columnCount;
		w = f.textFramePreferences.textColumnFixedWidth;
		//w = 40.25;
		for (var j = 0; j < t.columns.length; j++) {
			t.columns[j].width = (w / c) * 1;
		}
	}
}

//Add soft returns based on @minlines  --  takes xmlElement
function setMinLines(s) {
	var bQuoteArrays = s.evaluateXPathExpression("//*[@minLines]");
	var mLines;
	var mValue;
	for (var i = 0; i < bQuoteArrays.length; i++) {
		mLines = bQuoteArrays[i].xmlContent.lines.length -1;
		mValue = bQuoteArrays[i].xmlAttributes.itemByName('minLines').value;
		multiplier = mValue - mLines;
		
		if (mLines < mValue) {
			var lastEl = bQuoteArrays[i].evaluateXPathExpression("/*[last()]");
			var lastItem = lastEl[lastEl.length -1];
			lastItem.insertionPoints[-3].contents = Array(multiplier + 1).join("\n");
		}
	}
}

//Autosize anchored text frames  --  takes text frame
function resizeAnchoredFrames(f) {
	var s = f.parentStory
	var pItems = s.pageItems.everyItem().getElements();
	for (var i = 0; i < pItems.length; i++) {
		if (pItems[i] == "[object TextFrame]") {
			pItems[i].recompose();
			pItems[i].textFramePreferences.autoSizingType = AutoSizingTypeEnum.HEIGHT_ONLY;
			//pItems[i].textFramePreferences.autoSizingType = AutoSizingTypeEnum.OFF;
		}
	}
}

//Applies master page based on @master
function setMasterPages() {
	var x = myRoot.evaluateXPathExpression("//*[@master]");
	
	for (var j = 0; j < x.length; j++) {
		var aMaster = x[j].xmlAttributes.item('master').value;
		var storyID;
		try {
			storyID = x[j].xmlAttributes.item(idString).value;
		}
		catch (e) {
			storyID = 'No ID Available';
		}
		var masterPages = myDoc.masterSpreads;
		var myMaster;
		for (var i = 0; i < masterPages.length; i++) {
			mp = masterPages[i].namePrefix;
			if (mp == aMaster) {
				myMaster = masterPages[i];
			}
		}
		try {
			x[j].xmlContent.insertionPoints[0].parentTextFrames[0].parentPage.appliedMaster = myMaster;
			var p = x[j].xmlContent.insertionPoints[0].parentTextFrames[0].parentPage;
			var m1 = p.marginPreferences;
			var f = x[j].xmlContent.insertionPoints[0].parentTextFrames[0];
			if (p.side == PageSideOptions.RIGHT_HAND) {
				f.geometricBounds =[p.bounds[0] + m1.top, p.bounds[1] + m1.left, p.bounds[2] - m1.bottom, p.bounds[3] - m1.right];
			} else if (p.side == PageSideOptions.LEFT_HAND) {
				f.geometricBounds =[p.bounds[0] + m1.top, p.bounds[1] + m1.right, p.bounds[2] - m1.bottom, p.bounds[3] - m1.left];
			} else if (p.side == PageSideOptions.SINGLE_SIDED) {
				f.geometricBounds =[p.bounds[0] + m1.top, p.bounds[1] + m1.right, p.bounds[2] - m1.bottom, p.bounds[3] - m1.left];
			}
		}
		catch (e) {
			$.writeln('Error with story ' + storyID + ': story not placed into layout');
		}
	}
}

//Applies paragraph shading to anchored images  --  takes xmlElement
function imgShading() {
	var anchors = myRoot.evaluateXPathExpression("//*[@shade]");
	var p, img, h;
	for (var i = 0; i < anchors.length; i++) {
		if (anchors[i].xmlContent == "[object Text]") {
			p = anchors[i].paragraphs[0];
			if (p.pageItems.length > 0) {
				img = p.pageItems[0];
				h = img.geometricBounds[2] - img.geometricBounds[0] + .6;
				p.paragraphShadingTopOffset = h;
			}
		}
	}
}

//Renoves pages that become blank after
function removeBlankPages() {
	var tfs1;
	var empty =[];
	for (var bp = 0; bp < stories.length; bp++) {
		tfs1 = stories[bp].parentStory.textContainers;
		if (tfs1[tfs1.length -1].contents == '') {
			empty.push(tfs1[tfs1.length -1].parentPage);
		}
	}
	for (var ep = empty.length -1; ep >= 0; ep--) {
		var pp = empty[ep];
		pp.remove();
	}
}

function setStartTime() {
	var d = new Date();
	var time = d.getTime();
	return time;
}

function getTime(t) {
	var d = new Date();
	var time = d.getTime();
	return time - t;
}

function msToTime(duration) {
	var milliseconds = parseInt((duration % 1000) / 100);
	var seconds = parseInt((duration / 1000) % 60);
	var minutes = parseInt((duration /(1000 * 60)) % 60);
	var hours = parseInt((duration /(1000 * 60 * 60)) % 24);
	
	hours = (hours < 10) ? "0" + hours: hours;
	minutes = (minutes < 10) ? "0" + minutes: minutes;
	seconds = (seconds < 10) ? "0" + seconds: seconds;
	
	return hours + ":" + minutes + ":" + seconds + "." + milliseconds;
}

//Relinks the xml
function relinkXML() {
	var myLinks = myDoc.links;
	var myXMLLink;
	for (var i = 0; i < myLinks.length; i++) {
		temp = myLinks[i].linkType;
		//$.writeln(temp);
		
		if (temp == 'XML') {
			myXMLLink = myLinks[i];
		}
	}
	var linkPath;
	var testL
	try {
		if (myXMLLink.status == (LinkStatus.NORMAL || LinkStatus.LINK_OUT_OF_DATE)) {
			linkPath = myXMLLink.filePath;
		}
	}
	catch (e) {
		alert('Missing XML');
	}
	
	myXMLLink.relink(new File(linkPath));
}

function setSaveFreq() {
	//items in drop down
	var freq =[ 'none', 1, 2, 3, 4, 5, 10, 15, 20, 25];
	var startScript = false;
	
	//set up dialog window
	myDlg = new Window('dialog', 'Auto-save');
	//choose rows or columns
	myDlg.orientation = 'column';
	myDlg.alignment = 'right';
	//drop-down setup
	myDlg.DDgroup1 = myDlg.add('group');
	myDlg.DDgroup1.orientation = 'row';
	
	//add static text
	myDlg.DDgroup1.add('statictext', undefined, "Save frequency:");
	//add drop down list
	myDlg.DDgroup1.DD = myDlg.DDgroup1.add('dropdownlist', undefined, undefined, {
		items: freq
	})
	
	//add default selection
	myDlg.DDgroup1.DD.selection = 0;
	
	myDlg.DDgroup2 = myDlg.add('group');
	myDlg.DDgroup2.orientation = 'row';
	//add button
	myDlg.okBtn = myDlg.DDgroup2.add('button', undefined, 'OK');
	myDlg.cancelBtn = myDlg.DDgroup2.add('button', undefined, 'Cancel');
	
	// add button functions
	myDlg.okBtn.onClick = function () {
		//close window
		startScript = true;
		myDlg.DDgroup2.parent.close();
	}
	myDlg.cancelBtn.onClick = function () {
		//close window
		startScript = false;
		myDlg.DDgroup2.parent.close();
	}
	//show window and store drop down value
	result = myDlg.show();
	if (myDlg.DDgroup1.DD.selection.toString() == 'none') {
		return[ 'none', startScript];
	}
	
	return[myDlg.DDgroup1.DD.selection.toString() * 1, startScript];
}