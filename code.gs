// 
let spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
let ui = SpreadsheetApp.getUi();
let bold = {};
bold[DocumentApp.Attribute.BOLD] = true;
//______________________________Just adding the Build Resume to the UI and binding it with 
function onOpen(){                                      
   ui.createMenu('Build').addItem('Build Resume','startBuild').addToUi();
  startBuild();
}

//______________________________Initiateing build
function startBuild(){
  
  let doc = DocumentApp.create(ui.prompt('Creating a new Doc File for your resume','Name Your Resume File',ui.ButtonSet.OK_CANCEL).getResponseText());
  
  // let doc = DocumentApp.openByUrl('https://docs.google.com/document/d/1_BLMBJhVdrvQglYphMoNbpFMPTKh7PmmHTpVb2RPGJI/edit');
  
  buildResume(doc);
}

function buildResume(doc){
  //building the header of the resume
  let body = doc.getBody();
  let fullData = spreadSheet.getActiveSheet().getDataRange().getValues();
  let data = [];
  for (i=1 ; i < fullData.length ; i++){
    if(fullData[i][0] == true){
      data.push(fullData[i]);
    }
  }
  let categories = [];
  for( i=1 ; i < data.length ; i++) {
    if (categories.includes(data[i][1]))
      continue;
    else
      categories.push(data[i][1]);
  }
  categories.sort();
  for ( i=0 ; i < categories.length ; i++){
    categoryHeading = body.appendParagraph(categories[i]);
    categoryHeading.setHeading(DocumentApp.ParagraphHeading.HEADING4).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    for(j=0 ; j <= data.length-1 ; j++){
      if (data[j][1] == categories[i]){
        element = data[j]
        body.appendParagraph(element[2]);
        body.appendParagraph(element[3]);
        body.appendListItem(element[4]).setGlyphType(DocumentApp.GlyphType.BULLET);
        body.appendListItem(element[5]).setGlyphType(DocumentApp.GlyphType.BULLET);
        body.appendListItem(element[6]).setGlyphType(DocumentApp.GlyphType.BULLET);
      }
    }
  }
}
