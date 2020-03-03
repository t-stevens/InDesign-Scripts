/*SET-UP*/
var myDoc = app.activeDocument;

/*END SETUP*/

//FUTURE TESTS
//check images that need to bleed
//check text is 1/8 " from page edge
//presets

//Run corrections
UnassignEmptyFrames ();
RemoveUnusedSwatches();
SetUnitsToPicas ();
//ChangeBaseStylesToBasic ();

//Get warnings
var warnings = [];
var wAlert = []
var wFlag1 = NoPhotoshopEPS ();
var eps1 = wFlag1.join("\n    ");
wAlert.push("This document contains the following Photoshop EPS files:\n    " + eps1);
if (wFlag1.length > 0){
    warnings.push(false);
    }
else{
    warnings.push(true);
    }

wAlert.push("The Church logo is not an EPS version");
warnings.push(IsLogoEPS ());

wAlert.push("This document contains no paragraph styles");
warnings.push(HasParaStyles ());

wAlert.push("This document contains no character styles");
warnings.push(HasCharStyles ());

wAlert.push("This document contains non CMYK swatches");
warnings.push(AreSwatchesCMYK ());

//Get errors
var errors = [];
var eAlert = [];
var eFlag1 = FindBadFonts ();
var bf = eFlag1.join("\n    ");
eAlert.push("The following non-LDS fonts are used:\n    " + bf);
if (eFlag1.length > 0){
    errors.push(false);
    }
else{
    errors.push(true);
    }

var eFlag2 = AreLinksValid ();
var l1 = eFlag2.join("\n    ")
eAlert.push("The following files are not linked to the image servers:\n    " + l1);
if (eFlag2.length > 0){
    errors.push(false);
    }
else{
    errors.push(true);
    }

//Display Findings
var show = [];
show.push("ERRORS:");
for(var i = 0; i < errors.length; i++){
    if(!errors[i]){
        show.push(eAlert[i]);
        }
    }
if(show.length == 1){
    show[0] = "No errors found";
    }
var w = "WARNINGS:";
show.push(w);
for(var i = 0; i < warnings.length; i++){
    if(!warnings[i]){
        show.push(wAlert[i]);
        }
    }
if(show[show.length-1] == w){
    show[show.length-1] = "No warnings found";
    }

var title = "Validation";
alert_scroll (title, show);


/*WARNINGS*/

//check for Photoshop EPS links
function NoPhotoshopEPS(){
    var myLinks = myDoc.links;
    var myEPS = [];
    var temp;
    for(var i = 0; i < myLinks.length; i++){
        temp = myLinks[i].linkType;
        if(temp == 'EPS'){
            myEPS.push(myLinks[i]);
            }
        }
    var psEPS = [];
    for(var i = 0; i < myEPS.length; i++){
        temp = new File(myEPS[i].filePath).creator;
        if(temp == "8BIM"){
            psEPS.push(myEPS[i].name);
        }
    }
    return psEPS;
}
//check Church logo is eps, not tif
function IsLogoEPS (){
    var myLinks = myDoc.links;
    var temp;
    var logos = [];
    for(var i = 0; i < myLinks.length; i++){
        temp = myLinks[i].name.toLowerCase();
        if(temp.search('chlogo') != -1){
            logos.push(myLinks[i]);
            }
        }
    for(var i = 0; i < logos.length; i++){
        temp = logos[i].linkType;
        if(temp != 'EPS'){
            return false;
            }
        }        
    return true;
    }
//check for paragraph styles
function HasParaStyles(){
    var pStyles = myDoc.allParagraphStyles;
    if(pStyles.length > 2){
        return true;
        }
    else{
        return false;
        }
    }
//check for character styles
function HasCharStyles(){
    var cStyles = myDoc.allCharacterStyles;
    if(cStyles.length > 1){
        return true;
        }
    else{
        return false;
        }
    }

//Check for non-CMYK swatches
function AreSwatchesCMYK(){
    var s = myDoc.swatches;
    var temp;
    for(var i = 4; i < s.length; i++){
            temp = s[i].space;
            if(temp != 'CMYK'){
                return false;
                }
        }
    return true;
    }
/*END WARNINGS*/

/*ERRORS*/
//check for non-LDS fonts
function FindBadFonts(){
    var myDocFonts = myDoc.fonts;
    var fName;
    var badFonts = [];
    for(var f = 0; f < myDocFonts.length; f++){
        fName = myDocFonts[f].fontFamily;
        if(fName.search('lds') == -1){
            badFonts.push(fName);
        }
    }
    return badFonts;
}
function AreLinksValid (){
    var myLinks = myDoc.links;
    var temp, temp2, temp3;
    var badLinks = [];
    for(var i = 0; i < myLinks.length; i++){
        temp = myLinks[i].filePath;
        temp2 = temp.split(':');
        temp3 = temp2[0];
        if((temp3 != 'CURMMD') && (temp3 != 'Image')){
            badLinks.push(temp);
            }
        }
    return badLinks;
    }
/*END ERRORS*/

/*CORRECTIVE FUNCTIONS*/
//update base styles
function ChangeBaseStylesToBasic(){
    var pStyles = myDoc.allParagraphStyles;
    var basicStyle = pStyles[1];

    var base;
    for(var i = 2; i < pStyles.length; i++){
        base = pStyles[i].basedOn;
        if(!base.isValid){
            pStyles[i].basedOn = basicStyle;
            }
        }
    }
//fix empty frames
function UnassignEmptyFrames(){
    var s = myDoc.stories;
    for(var i = 0; i < s.length; i++){
        var sText = s[i].textContainers;
        for(var j = 0; j < sText.length; j++){
            if(sText[j].contents === ""){
                try{
                sText[j].contentType = ContentType.UNASSIGNED;
                }
                catch(e){}
                }
            }
        }
}
//remove unused colors
function RemoveUnusedSwatches(){
    app.menuActions.item("$ID/Add All Unnamed Colors").invoke();
    var myUnusedSwatches = myDoc.unusedSwatches;  
    for (var s = myUnusedSwatches.length-1; s >= 0; s--) {  
        var mySwatch = myUnusedSwatches[s];  
        var name = mySwatch.name;  
        if (name != ""){  
            mySwatch.remove();  
        }  
    }  
}
//set units to picas
function SetUnitsToPicas(){
    myDoc.viewPreferences.horizontalMeasurementUnits = MeasurementUnits.PICAS;
    myDoc.viewPreferences.verticalMeasurementUnits = MeasurementUnits.PICAS;
    }
/*END CORRECTIVE FUNCTIONS*/

/*ALERT*/
function alert_scroll (title, input){
   if (input instanceof Array)
       input = input.join ("\r");
   var w = new Window ("dialog", title);
   var list = w.add ("edittext", undefined, input, {multiline: true, scrolling: true});
   list.maximumSize.height = w.maximumSize.height-400;
   list.minimumSize.width = 550;
   w.add ("button", undefined, "Close", {name: "ok"});
   w.show ();
}
/*END ALERT*/





