function onInstall() {
  onOpen();
}

/* What should the add-on do when a document is opened */
function onOpen() {
  SlidesApp.getUi()
  .createAddonMenu() // Add a new option in the Google Docs Add-ons Menu
  .addItem("Add page numbers", "insertPageNumber")
  .addItem("Remove page numbers", "removeAllPageNumber")
  .addSeparator()	
  .addItem("Advanced", "showSidebar")
  .addToUi();  // Run the showSidebar function when someone clicks the menu
}

/* Show a 300px sidebar with the HTML from googlemaps.html */
function showSidebar() {
  var html = HtmlService.createTemplateFromFile("addPN")
    .evaluate()
    .setTitle("Page Number Manager"); // The title shows in the sidebar
  SlidesApp.getUi().showSidebar(html);
}

/* This Google Script function does all the magic. */
function insertPageNumber() {  
  var boxWidth = 80
  var boxHeight = 40
  
  
  var active_pres = SlidesApp.getActivePresentation()
  var pageHeight = active_pres.getPageHeight()
  var pageWidth = active_pres.getPageWidth()
  
  var slides = active_pres.getSlides()
  
  for(var i=0; i<slides.length;i++)
  {
    var shape = slides[i].insertTextBox((i+1).toString()+" / "+slides.length, pageWidth-boxWidth, pageHeight-boxHeight, boxWidth, boxHeight).setTitle("pageNumber")
    shape.getText().getTextStyle().setFontSize(15)
    shape.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.END)
  }
}
function insertPageNumberBasedOnTemplate(){
  var template = getTemplateFromPage();
  if(!template){
    var ui = SlidesApp.getUi();
    var result = ui.alert(
      'Please confirm',
      "There's no template in this slide. Do you want to create one?",
      ui.ButtonSet.YES_NO);
    
    // Process the user's response.
    if (result == ui.Button.YES) {
      insertPageNumberTemplate()
      ui.alert("Warning", "Please edit the template first!");
      return;
    } else {
      ui.alert("Warning", "There's no template existed in this slide! Please add one.");
      return;
    }
  }
  
  var active_pres = SlidesApp.getActivePresentation()

  var slides = active_pres.getSlides()
  var text;
  for(var i=0; i<slides.length;i++)
  {
    var newPNEle = slides[i].insertPageElement(template)
    newPNEle.setTitle("pageNumber")
    text = newPNEle.asShape().getText()
    text.replaceAllText("#T#", slides.length)
    text.replaceAllText("#N#", i+1);    
  }
}
function insertPageNumberTemplate() {  
  
  
  var template = getTemplateFromPage()
  
  if(template){
      if(showAlertDupTemplate()){
        template.remove()
      }else{
        return;
      }
    }
  var active_pres = SlidesApp.getActivePresentation()
  var pageHeight = active_pres.getPageHeight()
  var pageWidth = active_pres.getPageWidth()
  var boxWidth = 80
  var boxHeight = 40
  
  var slide = active_pres.getSelection().getCurrentPage();
  slide.insertTextBox("#N#/#T#", pageWidth-boxWidth, pageHeight-boxHeight, boxWidth, boxHeight).setTitle("pageNumberTemplate")
  shape.getText().getTextStyle().setFontSize(15)
  shape.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.END)
}
function showAlertDupTemplate() {
  var ui = SlidesApp.getUi();

  var result = ui.alert(
     'Found a existed template',
     'Do you want to remove it?',
      ui.ButtonSet.YES_NO);

  // Process the user's response.
  if (result == ui.Button.YES) {
    // User clicked "Yes".
    return true
  } else {
    // User clicked "No" or X in the title bar.
    ui.alert("Warning", "There's a template existed in this slide! Please edit that one.");
    return false;
  }
}
function removeAllPageNumber(){
  var active_pres = SlidesApp.getActivePresentation()
  var slides = active_pres.getSlides()
  for(var i=0; i<slides.length;i++)
  {
    var pgElements = slides[i].getPageElements()
    Logger.log(pgElements.length)
    for(var j=0; j< pgElements.length; j++){
      if(pgElements[j].getTitle()=="pageNumber"){
        pgElements[j].remove()
      }
    }
  }
}
function removePageNumberTemplate(){
  var template = getTemplateFromPage();
  if(template){
    template.remove()
  }
}
function getTemplateFromPage(){
  var active_pres = SlidesApp.getActivePresentation()
  var pageHeight = active_pres.getPageHeight()
  var pageWidth = active_pres.getPageWidth()
  
  var slide = active_pres.getSelection().getCurrentPage();
  var pgElements = slide.getPageElements()
  for(var j=0; j< pgElements.length; j++){
    if(pgElements[j].getTitle()=="pageNumberTemplate"){
      return pgElements[j];
    }
  }
  return false
}