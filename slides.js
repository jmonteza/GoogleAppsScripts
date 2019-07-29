// TODO:
// Long names, school logo, menu

function createPresentation() {
  
    var year, school, name, fileName, colour, foreground, font, nameSize, titleSize, gradSize, schoolLogo, deck;
    
    year = new Date().getFullYear().toString(); // Current Year
    
    school = 'Crestwood'; // School Name
  
    name = "Class of " + year;
    
    fileName = "Class of " + year + " - " + school; // Class Year
    
    colour = "#C53023"; // School Colour
    
    foreground = "#000000"; // Foreground Colour (Black)
    
    font = "Roboto"; // Font 
    
    nameSize = 70; // Font size for names
    
    titleSize = 80; // Font size for title
    
    gradSize = 33; // Font size for Grad 2020
    
    schoolLogo = 'https://upload.wikimedia.org/wikipedia/en/thumb/3/31/JasperPlace_Logo.svg/1156px-JasperPlace_Logo.svg.png'; // School Logo
    
    deck = SlidesApp.create(fileName); // Creates the presentation
  
    title(deck, font, name, colour, school, foreground, titleSize );
    
    createSlides(deck, schoolLogo, year, gradSize, font);
    
    createStudents(deck);
  }
  
  function title(deck, font, name, colour, school, fcolour, fsize ){
    
    var [title, subtitle] = deck.getSlides()[0].getPageElements();
    
    // Class of ${Current Year}
    title.asShape().getText().setText(name);
    title.asShape().getText().getTextStyle().setBold(true).setForegroundColor(colour).setFontSize(fsize).setFontFamily(font);
    
    // School Name 
    subtitle.asShape().getText().setText(school);
    subtitle.asShape().getText().getTextStyle().setBold(true).setFontFamily(font).setForegroundColor(fcolour).setFontSize(33);
    
  }
  
  
  function countFiles() {
    // Function to count the number of files in the folder
    var files = DriveApp.getFoldersByName("Test").next().getFiles();
    
    var count = 0;
    
    while (files.hasNext()){
      count++;
      file = files.next();
    }
  
  return count;
  }
  
  function measureImage(){
    
  }
  
  
  function createSlides(deck, logoURL, year, fontSize, font) {
    // Function to create slides
    var numSlides = parseInt(countFiles()); // Number of files in Graduation folder
  
    for (var i = 0; i < numSlides; i++) {
    slide = deck.appendSlide(SlidesApp.PredefinedLayout.BLANK);
      
    var logo = slide.insertImage(logoURL).setTop(322).setLeft(50).setHeight(60).setWidth(80);
    var caption = 'GRAD' + " " + year;
    var shape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, 149, 324, 260, 90);  
    var textRange = shape.getText();
    textRange.setText(caption).getTextStyle().setFontSize(fontSize).setFontFamily(font).setBold(true);
  }
  }
  
  function createStudents(deck){
    // Get names and pictures from file
    
    var count = 1;
    
    // Student Names
    var folder = "Test" // Folder Name
    var files = DriveApp.getFoldersByName(folder).next().getFiles(); // Get files from folder
  
    while (files.hasNext()){
      var file = files.next();
      
      var slide = deck.getSlides()[count]
      
      // Get and format students' name from file name.
      var studentName = file.getName();
      var studentf = studentName.split('.').slice(0, -1).join('.')
      var fields = studentf.split('_');
      var firstname = fields[1] ;
      var lastname = fields[0];
      
      // Place the name on the slide
      var shape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, 50, 80, 320, 120);
      var textRange = shape.getText();
      textRange.setText(firstname + '\n' + lastname).getTextStyle().setFontSize(60).setFontFamily('Roboto').setBold(false).setForegroundColor('#C53023');
      var paragraphs = textRange.getParagraphs();
      for (var i = 0; i < 2; i++) {
        var paragraphStyle = paragraphs[i].getRange().getParagraphStyle();
        paragraphStyle.setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);  
      }
  
       
      // Get the image 
       
        var image = slide.insertImage(file.getAs('image/png'));
        var imgWidth = image.getWidth();
        var imgHeight = image.getHeight();
        var pageWidth = deck.getPageWidth();
        var pageHeight = deck.getPageHeight();
        var newX = pageWidth/2. - imgWidth/2.;
        var newY = pageHeight/2. - imgHeight/2.;
        image.setLeft(newX+200).setTop(newY);
        
      // Format the image's location on the slide.
       
  
      count++;
  
  }
  }
  