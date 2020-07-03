function onInstall(e) {
  onOpen(e);
}

/* What should the add-on do when a document is opened */
function onOpen(e) {
  SlidesApp.getUi()
  .createAddonMenu() 
  .addItem("Add page numbers", "insertPageNumber")
  .addItem("Remove page numbers", "removeAllPageNumber")
  .addSeparator()	
  .addItem("Add template", "insertPageNumberTemplate")
  .addItem("Add custom page numbers", "insertPageNumberBasedOnTemplate")
  .addSeparator()	
  .addItem("Advanced", "showSidebar")
  .addToUi();  
}

/* Show a 300px sidebar with the HTML from googlemaps.html */
function showSidebar() {
  var html = HtmlService.createTemplateFromFile("addPN")
    .evaluate()
    .setTitle("Page Number"); // The title shows in the sidebar
  SlidesApp.getUi().showSidebar(html);
}
function onEdit(e){
  

}
function updatePageNumber(){
  
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
      "There's no template in this slide. Do you want to create one and add page number with default setting?",
      ui.ButtonSet.YES_NO);
    
    // Process the user's response.
    if (result == ui.Button.YES) {
      insertPageNumberTemplate()
      template = getTemplateFromPage()
    } else {
      SlidesApp.getUi().alert("Warning", "There's no template existed in this slide! Please add one.", ui.ButtonSet.OK);
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
    text.replaceAllText("#C#", i+1);    
  }
  template.setDescription(template.getTop()+"/"+template.getLeft())
  template.setTop(active_pres.getPageHeight()).setLeft(active_pres.getPageWidth())
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
  var boxWidth = 100
  var boxHeight = 40
  
  var slide = active_pres.getSelection().getCurrentPage();
  var shape = slide.insertTextBox("#C# / #T#", pageWidth-boxWidth, pageHeight-boxHeight, boxWidth, boxHeight).setTitle("pageNumberTemplate")
  shape.getText().getTextStyle().setFontSize(15)
  shape.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.END)
  return shape
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
    for(var j=0; j< pgElements.length; j++){
      if(pgElements[j].getTitle()=="pageNumberTemplate"&&pgElements[j].getDescription()!=''){
        var pos = pgElements[j].getDescription().split('/')
        if(pos.length==2){
            pgElements[j].setTop(pos[0]).setLeft(pos[1])
            pgElements[j].setDescription("")
            SlidesApp.getUi().alert("Restore a template on page "+(i+1)+".");
        }
      }
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