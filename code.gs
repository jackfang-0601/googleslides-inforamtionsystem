
function getResponse() {
  const presentation = SlidesApp.openByUrl("https://docs.google.com/presentation/d/1TS87SCW_42sKNSVg6C7ycKUvwLvf9CV9T4RtVeDg5BU/edit#slide=id.SLIDES_API805723590_0");
  var slides = presentation.getSlides();
  for (var i = slides.length - 1; i >= 0; i--) { 
    var slide = slides[i];
    slide.remove();
  }
  const sheets = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1FnMiVE9y4zJsKXv5D6pGZVSUZmGTOE1vKcIi2EEKiFs/edit?resourcekey#gid=2088115445").getSheetByName("sheet1");
   var last_row = sheets.getLastRow();
   var previous_num = 0;
   const numberList = [];
   var a = 0;

   for (let i = 1;i<=last_row;i++){

    var type_of_fixture = sheets.getRange(i, 2).getValue();
    var degreed_Barrels = sheets.getRange(i, 3).getValue();
    var lense_type = sheets.getRange(i, 4).getValue();
    var identification_num = sheets.getRange(i, 5).getValue();
    var working = sheets.getRange(i, 6).getValue();
    var if_no_how = sheets.getRange(i, 7).getValue();
    var location = sheets.getRange(i, 8).getValue();
    var additional_notes = sheets.getRange(i, 9).getValue();
    
    const layout = presentation.getLayouts();
    var curSlide = presentation.appendSlide(layout[0]);
    var elements = curSlide.getPageElements();
    var titleArea = elements[0].asShape().getText();
    var titleTwoArea = elements[6].asShape().getText();
    var topLeftArea = elements[1].asShape().getText();
    var bottomLeftArea = elements[2].asShape().getText();
    var topRightArea = elements[3].asShape().getText();
    var bottomRightArea = elements[4].asShape().getText();
    var bigArea = elements[5].asShape().getText();
    var photo_url = sheets.getRange(i, 10).getValues().toString();

    
    if (degreed_Barrels == ""){
      topLeftArea.setText("Lense type:\n"+lense_type);
    }
    if (lense_type == ""){
      topLeftArea.setText("Degree:\n"+degreed_Barrels);
    }
    if (lense_type == ""&&degreed_Barrels == ""){

      topLeftArea.setText("n/a")
    
    }


    titleArea.setText("Type of fixture:\n"+type_of_fixture);
    titleTwoArea.setText("Identification Number: " +identification_num);
    bottomLeftArea.setText("Location:\n"+location);
    topRightArea.setText("Is the fixture working properly?\n"+working);
    bottomRightArea.setText("State the issue:\n "+if_no_how);
    bigArea.setText("Notes: \n "+additional_notes);
    const slide1 = presentation.getSlides()[i-1];
    if (photo_url == ""){

    }
    else{
    photo1 = getIdSheetFromUrl_(photo_url);
    slide1.insertImage(photo1,450,200,200,200);
    }
    if (previous_num == identification_num){
      numberList[a]= i-2;
      a+= 1;
    }


    previous_num = identification_num;
   }
  
  var slides = presentation.getSlides();
  for (var i = numberList.length - 1; i >= 0; i--) { 
    var slide = slides[numberList[i]];
    slide.remove();
  }
  


  

 
  
   
}
function getIdSheetFromUrl_(url)
{
    var id = url.split('id=')[1];
    if(!id)
    {
        id = url.split('/d/')[1];
        id = id.split('/edit')[0]; 
    }
    return DriveApp.getFileById(id);
}


  
