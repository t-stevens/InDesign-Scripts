var myBook = app.activeBook;
var bookFiles = myBook.bookContents;

var myNotes;
var myDocs =[];

for (var i = 0; i < bookFiles.length; i++) {
	var f = bookFiles[i].fullName;
	var fName = bookFiles[i].name;
	if (fName.toLowerCase().match(/notes/)) {
		myNotes = app.open(f, false);
	} else {
		temp = app.open(f, false);
		myDocs.push(temp);
	}
}
var notesRoot = myNotes.xmlElements.item(0);
var noteElements = notesRoot.evaluateXPathExpression("//chapNotes//*[@noteRef]");
var markerElements = getBodyItems();


var pairs = [];
var n, m;
for(var i = 0; i < markerElements.length; i++){
    m = markerElements[i].xmlAttributes.itemByName('noteChap').value;
    for(var j = 0; j < noteElements.length; j++){
        n = noteElements[j].xmlAttributes.itemByName('noteRef').value;
        if(n == m){
            pairs.push([markerElements[i], noteElements[j]]);
            }
        }
    }


for (var i = 0; i < pairs.length; i++) {
        var eNote = pairs[i][1];
        var nMark = pairs[i][0];
        var eNoteIP = eNote.insertionPoints[-2];
        //sIP.contents = '$%$';
       var destination = myNotes.hyperlinkTextDestinations.add(nMark.xmlContent);
       var xRefForm = myNotes.crossReferenceFormats.item("Page Number");
       var source = myNotes.crossReferenceSources.add(eNoteIP, xRefForm);
       var myLink = myNotes.hyperlinks.add(source, destination);
	}

close(SaveOptions.YES);
alert('DONE');

/*FUNCTIONS*/

function getBodyItems() {
	var allChaps =[];
	var tempRoot, temp;
	for (var docCount = 0; docCount < myDocs.length; docCount++) {
		tempRoot = myDocs[docCount].xmlElements.item(0);
		temp = tempRoot.evaluateXPathExpression("//sup");
		allChaps = allChaps.concat(temp);
	}
	return allChaps;
}

function close(save) {
	
	while (app.documents.length) {
		app.documents[app.documents.length -1].close(save);
	}
}