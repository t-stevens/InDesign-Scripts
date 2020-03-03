/* SET_UP */
var myBook = app.activeBook;
var bookFiles = myBook.bookContents;
var myIndex = app.activeDocument;
var myRoot = myIndex.xmlElements.item(0);
var partFiles = openBookFiles();

//turn off Adobe alerts
uia = app.scriptPreferences.userInteractionLevel;
app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;
//END turn off Adobe alerts
/* end SET_UP */


main();

//turn on Adobe alerts
app.scriptPreferences.userInteractionLevel = UserInteractionLevels.INTERACT_WITH_ALL;
//END turn on Adobe alerts
alert('DONE');

function main(){
    updateLocators();
    closeDocs();
}

/* FUNCTIONS */
function openBookFiles(){
    var temp;
    var bookArray = [];
    for (var bookCount = 0; bookCount < bookFiles.length; bookCount++) {
        var f = bookFiles[bookCount].fullName;
        var fName = bookFiles[bookCount].name;
        if ((fName.toLowerCase().match(/part/)) &&(!fName.toLowerCase().match(/intro/))) {
            temp = app.open(f, false);
		    bookArray.push(temp);
        }
    }
    return bookArray;
}

function updateLocators(){
    var locators = myRoot.evaluateXPathExpression("//index.LOCATOR/a[@href]");
    var hRefData, locFile, locChap, locStart, locEnd, pageString;
    //for(var locCount = 0; locCount < 100; locCount++){
    for(var locCount = 0; locCount < locators.length; locCount++){
        hRefData = getHRefData(locators[locCount].xmlAttributes.itemByName('href').value);
        locFile = getRefFile(hRefData[0]);
        locChap = hRefData[1];
        locStart = hRefData[2];
        locEnd = hRefData[3];
        pageString = getPageNums(locFile, locChap, locStart, locEnd);
        locators[locCount].contents = pageString;
    }
}

function getHRefData(fullHref){
    var t1, t2, t3, t4, t5, t6, t7;
    var refSplit1, refSplit2, refPart, refFile, refChap, refParas, refStartPara, refEndPara;
    //get part
    refSplit1 = fullHref.split('/');
    while(refSplit1[0] == '..'){
        refSplit1.shift();
    }
    refPart = refSplit1[2];
    t1 = refPart.split('-');
    refFile = t1.join('');
    //get chapter uri
    t2 = refSplit1.join('/');
    refSplit2 = t2.split('.');
    refChap = '/' + refSplit2[0];
    //get paragraph numbers
    t3 = refSplit2[1];
    t4 = t3.split('=');
    t5 = t4[1].split('#');
    refParas = t5[0];
    //get start and end paragraphs
    t6 = refParas.split('-');
    if(t6.length > 1){
        refStartPara = t6[0];
        refEndPara = t6[t6.length-1];
    }
    else{
        refStartPara = refParas;
        refEndPara = refParas;
    }
    if(refStartPara.indexOf(',') != -1){
        t7 = refStartPara.split(',');
        refStartPara = t7[0];
    }
    if(refEndPara.indexOf(',') != -1){
        t7 = refEndPara.split(',');
        refEndPara = t7[t7.length-1];
    }
    //return values in array
    return [refFile, refChap, refStartPara, refEndPara];
}

function getRefFile(myPart){
    var partTest = myPart.toLowerCase();
    //$.writeln(partTest);
    for (var docCount = 0; docCount < partFiles.length; docCount++) {
        var fName = partFiles[docCount].name;
        if (fName.toLowerCase().match(partTest)) {
            return partFiles[docCount];
        }
    }
    //return 0;
}

function getPageNums(myDoc, myChap, myStart, myEnd){
    var docRoot = myDoc.xmlElements.item(0);
    var matchString = "//chapter[@uri='" + myChap + "']//*[@id='" + myStart + "']";
    var startParaArray = docRoot.evaluateXPathExpression(matchString);
    var startPara = startParaArray[0];
    var endPara = startParaArray[0];
    if(myStart != myEnd){
        matchString = "//chapter[@uri='" + myChap + "']//*[@id='" + myEnd + "']";
        var endParaArray = docRoot.evaluateXPathExpression(matchString);
        endPara = endParaArray[0];
    }
    try{
        var startIP = startPara.insertionPoints[0];
        var endIP = endPara.insertionPoints[endPara.insertionPoints.length-4];
        var startFrame = startIP.parentTextFrames[0];
        var endFrame = endIP.parentTextFrames[0];
        var startPage = startFrame.parentPage.name;
        var endPage = endFrame.parentPage.name;
        var pageRange = 'ZZZ';
        if(startPage != endPage){
            if((startPage.length == 3) && (endPage.length == 3)){
                if(startPage.charAt(0) == endPage.charAt(0)){
                    if(endPage.charAt(1) != 0){
                        endPage = endPage.substr(1);
                    }
                    else{
                        endPage = endPage.substr(2);
                    }
                }
            }
            pageRange = startPage + '–' + endPage;
        }
        else{
            pageRange = startPage;
        }
        return pageRange;
    }
    catch(e){
        return 'ZZZ';
    }
}

function closeDocs(){
    for (var docCount = 0; docCount < partFiles.length; docCount++) {
        var fName = partFiles[docCount].name;
        if (!fName.toLowerCase().match(/index/)) {
            partFiles[docCount].close(SaveOptions.NO);
        }
        else{
            $.writeln(fName + ' did not close');
        }
    }
}
/* end FUNCTIONS */