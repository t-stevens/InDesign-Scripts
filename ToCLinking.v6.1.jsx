//turn off Adobe alerts
uia = app.scriptPreferences.userInteractionLevel;
app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;
//END turn off Adobe alerts

var myDoc = app.activeDocument;
var myRoot = myDoc.xmlElements.item(0);
var myErrors = [];

main();

//turn on Adobe alerts
app.scriptPreferences.userInteractionLevel = UserInteractionLevels.INTERACT_WITH_ALL;
//END turn on Adobe alerts

alert('DONE');

function main() {
	legacyTest();
	var XRefFormat = createCrossRefFormat();
	var tocHrefTags = myRoot.evaluateXPathExpression("//table-of-contents//*[not(contains(name(),'img'))][@href]");
	if(tocHrefTags.length == 0){
		tocHrefTags = myRoot.evaluateXPathExpression("//chapter[contains(@uri,'/contents')]//*[@href]");
	}
	var myTag, mySource, myRef, matchRef, myTitle, myDest, myXRef;
	var myBookmark = []
	for(var tagCount = 0; tagCount < tocHrefTags.length; tagCount++){
		myTag = tocHrefTags[tagCount];
		mySource = getSource(myTag);
		myRef = getReference(myTag);
		matchRef = myRef[0];
		myTitle = myRef[1];
		myDest = getDestination(matchRef, myTitle);
		myXRef = createHyperlinks(mySource, myDest, XRefFormat);
		myBookmark.push(createBookmark(myTag, myDest));
	}
	renameBookmark(myBookmark);
}

/* FUNCTIONS */
function legacyTest() {
	
	var lTest = myRoot.evaluateXPathExpression("//a[@data-legacy-href]");
	for (var i = 0; i < lTest.length; i++) {
		var myElement = lTest[i];
		var myRefTag = myElement.xmlAttributes.itemByName("data-legacy-href");
		try {
			myRefTag.name = "href";
		}
		catch (e) {
			myRefTag.remove();
		}
	}

	var crTest = myRoot.evaluateXPathExpression("//table-of-contents//*[@cross-ref]");
	if(crTest.length == 0){
		crTest = myRoot.evaluateXPathExpression("//chapter[contains(@uri,'/contents')]//*[@cross-ref]");
	}
	for (var i = 0; i < crTest.length; i++) {
		var myElement = crTest[i];
		if(myElement.parent.markupTag.name != 'a'){
			var myRefTag = myElement.xmlAttributes.itemByName("cross-ref");
			try {
				myRefTag.name = "href";
			}
			catch (e) {
				myRefTag.remove();
			}
		}
	}
}

function createCrossRefFormat() {
	var myXRefFormat, temp;
	var xRefs = myDoc.crossReferenceFormats;
	try {
		myXRefFormat = myDoc.crossReferenceFormats.itemByName('tocNum');
		temp = myXRefFormat.name;
		return myXRefFormat;
	}
	catch (e) {
		myXRefFormat = myDoc.crossReferenceFormats.add('tocNum');
		var bbs = myXRefFormat.buildingBlocks;
		bbs.add(BuildingBlockTypes.PAGE_NUMBER_BUILDING_BLOCK);
		return myXRefFormat;
	}
}

function getSource(refTag){
	var sourceTag = refTag;
	try{
		while(sourceTag.parent.markupTag.name != 'ul'){
			sourceTag = sourceTag.parent;
		}
	}
	catch(e){
		sourceTag = refTag;
		if(sourceTag.parent.xmlAttributes.itemByName('id').isValid){
			sourceTag = sourceTag.parent;
		}
	}
	var sourcePara = sourceTag.paragraphs[0];
	var sourceIP = sourcePara.insertionPoints;

	return sourceIP[sourceIP.length-2];
}

function getReference(aTag){
	var uriHref = '';
	var titleNum = '';
	var fullRef = aTag.xmlAttributes.itemByName("href").value;

	//get title piece
	var hashSplit = fullRef.split('#');
	if(hashSplit.length > 1){
		titleNum = hashSplit[hashSplit.length-1];
	}
	//separate by slashes
	var dashSplit = fullRef.split('/');
	//check for partial hrefs
	if(dashSplit.length == 1){
		var tempSplit = fullRef.split('.');
		uriHref = getDocPath(tempSplit[0]);
		return [uriHref, titleNum];
	}
	//remove relative path dots
	while(dashSplit[0] == '..'){
		dashSplit.shift();
	}
	//remove .html
	var chapPiece = dashSplit[dashSplit.length-1];
	var dotSplit = chapPiece.split('.');
	dashSplit[dashSplit.length-1] = dotSplit[0];
	//reassemble href
	uriHref = '/' + dashSplit.join('/');

	return [uriHref, titleNum];
}

function getDocPath(myPiece){
	var allUriElements = myRoot.evaluateXPathExpression("//*[@uri]");
	var pathRef = allUriElements[0];
	var myUri = pathRef.xmlAttributes.itemByName('uri').value;
	var uriSplit = myUri.split('/');
	var newPath = '/' + uriSplit[1] + '/' + uriSplit[2] + '/' + myPiece;
	return newPath;
}

function getDestination(myChap, myTitle){
	var myUriElement = myRoot.evaluateXPathExpression("//*[@uri='" + myChap + "']");
	if(myUriElement.length == 0){
		var testPiece = myChap.split('/');
		myChap = getDocPath(testPiece[testPiece.length-1]);
		myUriElement = myRoot.evaluateXPathExpression("//*[@uri='" + myChap + "']");
	}
	var chapElement, destElement;
	if(myTitle == ''){
		return myUriElement[0].paragraphs[0];
	}
	chapElement = myUriElement[0];
	destElement = chapElement.evaluateXPathExpression("//*[@id='" + myTitle + "']");
	if(destElement.length != 0){
		return destElement[0].paragraphs[0];
	}
	else{
		myErrors.push(myChap + '#' + myTitle);
		return myUriElement[0].paragraphs[0];
	}
}

function createHyperlinks(myS, myD, myF) {
	var destination = myDoc.hyperlinkTextDestinations.add(myD);
	var source = myDoc.crossReferenceSources.add(myS, myF);
	var myLink = myDoc.hyperlinks.add(source, destination);
	return myLink;
}

function createBookmark(bkTag, bkDest){
	try{
		var bkPage = bkDest.parentTextFrames[0].parentPage;
	}
	catch(e){
		var bkPage = myDoc.pages[0];
	}
	var fullText = bkTag.paragraphs[0].contents;
	var tabSplit = fullText.split('\u0009');
	if(tabSplit[0] == ''){
		tabSplit.shift();
	}
	if(tabSplit.length > 1){
		tabSplit.pop();
	}
	var bkText = tabSplit.join(' ');
	var myBkMk = myDoc.bookmarks.add(bkPage);
	return [myBkMk, bkText];
}

function renameBookmark(bkmrks){
	for(var bCount = 0; bCount < bkmrks.length; bCount++){
		var bkmk = bkmrks[bCount][0];
		var bkmkName = bkmrks[bCount][1];
		bkmk.name = bkmkName;
	}
}