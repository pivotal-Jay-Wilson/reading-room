
interface bullet {
    title: string; 
    url?: string; 
    type: string
}

function onOpen() {
    var ui = SlidesApp.getUi();
    ui.createMenu('Reading Room')
        .addItem('Create TAS KBS', 'openDialog')
        .addToUi();
}

function openDialog() {
    var html = HtmlService.createHtmlOutputFromFile('index.html');
    SlidesApp.getUi().showModalDialog(html, 'Pick a Run Date');
}


function openDialog2() {
    var html = HtmlService.createHtmlOutputFromFile('index2.html');
    SlidesApp.getUi().showModalDialog(html, 'Pick a Run Date');
} 

function testKbs(){
    addKbs("2022-05-02", 5);
}

function addKbs(msg, lines) {
    const pres = SlidesApp.getActivePresentation();
    const allSlides = pres.getSlides();
    var kbtemp = allSlides[1];
    var slide = kbtemp.duplicate();
    slide.getShapes().forEach(shape => {
        if (shape.getShapeType() == SlidesApp.ShapeType.TEXT_BOX) {
            switch (shape.getText().asRenderedString().trim()) {
                case "TEMPLATE":
                    shape.remove();
                    break;
                case "Week of xx/xx/2020":
                    var date = Utilities.formatDate(new Date(), "GMT+1", "MM/dd/yyyy")
                    shape.getText().setText("Week of " + date)
                    break;
            }
        }
    }); 
    const leftside = slide.getShapes().find(shape => {
        return shape.getText().asRenderedString().trim().startsWith("{{leftside}}");
    });

    const rightside = slide.getShapes().find(shape => {
        return shape.getText().asRenderedString().trim().startsWith("{{rightside}}");
    });
    const textRangeLeft = leftside.getText();
    const textRangeRight = rightside.getText();

    textRangeLeft.clear();
    textRangeRight.clear();
    
    const ss = SpreadsheetApp.openById("1vPou7PjHA-VLkKpQe18uMIGIHzw7IqIYS319j4sjf4w");
    const sheet = ss.getSheetByName('TASKBFiltered');
    const range = sheet.getDataRange();
    const rows = range.getNumRows();
    var dateArray: string[] = msg.split("-");
    const today = new Date(parseInt(dateArray[0]), parseInt(dateArray[1]) -1, parseInt(dateArray[2]));
    let links:Array<bullet> = [];
    let linksO = [];
    let linksS = [];  
    for (let i = 1; i < rows; i++) {
        const update = new Date(range.getCell(i, 10).getValue());
        if (today.getFullYear() == update.getFullYear() && today.getMonth() == update.getMonth() && today.getDate() == update.getDate()) {
            const type = range.getCell(i, 11).getValue();
            const link = "https://community.pivotal.io/s/article/" + range.getCell(i, 8).getValue();
            const message = range.getCell(i, 9).getValue();
            if (type.toString().startsWith("O")) {
                linksO.push({ title: `${message}`, url: link, type: "B"});
            } else {
                linksS.push({ title: `${message}`, url: link, type: "B"})
            }
        }
    }

    links = links.concat({ title: "Opsmanager/TAS components", url: "", type: "H" });
    links = links.concat(linksO);
    if (linksS.length > 0){
        links = links.concat({ title: "Services/Other Tiles", url: "", type: "H" });
        links = links.concat(linksS);    
    }
    let headers: Array<GoogleAppsScript.Slides.TextRange> = [];
    let bullets: Array<GoogleAppsScript.Slides.TextRange> = [];
    let fit = fillSlide(links, textRangeLeft, 0, headers, bullets, leftside, lines);
    Logger.log(`length  = ${links.length} and fit = ${fit}, `);
    if (fit < links.length ) {
        Logger.log("here");
        fillSlide(links, textRangeRight, fit, headers, bullets, rightside, lines + fit );
    }
    
    for (const bullet of bullets) {
        formatShape(bullet);
    }

    for (const header of headers) {
        header.getParagraphStyle().setIndentStart(0).setSpaceAbove(2).setSpaceBelow(3);
        header.getListStyle().removeFromList();
        header.getTextStyle().setForegroundColor("#3f3f3f").setBold(true);
    }    
}

function formatShape (shape: GoogleAppsScript.Slides.TextRange){
    const ps = shape.getParagraphStyle();
    ps.setIndentEnd(0);
    ps.setLineSpacing(115);    
    ps.setSpaceAbove(0);
    ps.setSpaceBelow(5);
    ps.setIndentFirstLine(15);
    ps.setIndentStart(21);  
}

function fillSlide(links: Array<bullet> , 
                    textRange: GoogleAppsScript.Slides.TextRange, 
                    maxIndex: number, 
                    headers: Array<GoogleAppsScript.Slides.TextRange>, 
                    bullets: Array<GoogleAppsScript.Slides.TextRange>, 
                    shape: GoogleAppsScript.Slides.Shape,
                    lines) {
    let rowHeight = 0; 
    let fits = true;
    let index = 0;
    for (const link of links) {
      index++;  
      if ( index <= maxIndex) continue;  
      let h: GoogleAppsScript.Slides.TextRange;
      switch (link.type) {
          case "B":
              h = textRange.appendText(`${link.title}`);
              h.getTextStyle().setLinkUrl(link.url).setForegroundColor("#0091da").setBold(false);
              bullets.push(h);              
              break;
          case "H":
              h = textRange.appendText(`${link.title}`);
              h.getTextStyle().setForegroundColor("#3f3f3f").setBold(true);
              headers.push(h);
              break;
      }
      if (index >= lines){
        fits = false;
        break;
      } else {
        h = textRange.appendText("\n");
      }
    }
    textRange.getListStyle().applyListPreset(SlidesApp.ListPreset.DISC_CIRCLE_SQUARE);
    return index;
}


function getHeight(tr: GoogleAppsScript.Slides.TextRange) {
    let rowHeight = (tr.asString().split("\n").length - 1) * (tr.getTextStyle().getFontSize() / 72 * 96);
    const ps = tr.getParagraphStyle();
    rowHeight = rowHeight + ps.getSpaceAbove();
    rowHeight = rowHeight + ps.getSpaceBelow();
    rowHeight = rowHeight + (ps.getLineSpacing() / 100);
    return rowHeight;
}

function testBullets(){
    const pres = SlidesApp.getActivePresentation();
    const allSlides = pres.getSlides();
    const slide = allSlides[1];    
    const leftside = slide.getShapes().find(shape => {
        return shape.getText().asRenderedString().trim().startsWith("{{leftside}}");
    });
    const textRange = leftside.getText();
    textRange.clear();
    let ar = []
    ar.push(textRange.appendText(`Heading\n`));
    for (let index = 1; index < 4; index++) {
        ar.push(textRange.appendText(`This is a very long bullet to test the paragraph.  This is a very long bullet to test. ${index}\n`));
    }
    textRange.getListStyle().applyListPreset(SlidesApp.ListPreset.DISC_CIRCLE_SQUARE);

    let ps:GoogleAppsScript.Slides.ParagraphStyle = ar[0].getParagraphStyle();
    ps.setIndentEnd(0);
    //ps.setIndentStart(5);
    ps.setLineSpacing(115);    
    ps.setSpaceAbove(0);
    ps.setSpaceBelow(5);
    ar[0].getListStyle().removeFromList();

    ps = ar[1].getParagraphStyle();
    ps.setIndentEnd(0);

    ps.setIndentFirstLine(10);
    ps.setIndentStart(16);
    
    ps.setLineSpacing(115);    
    ps.setSpaceAbove(0);
    ps.setSpaceBelow(5);     
    
    ps = ar[2].getParagraphStyle();
    ps.setIndentEnd(0);

    ps.setIndentFirstLine(12);
    ps.setIndentStart(18);
    
    ps.setLineSpacing(115);    
    ps.setSpaceAbove(0);
    ps.setSpaceBelow(5);  

    ps = ar[3].getParagraphStyle();
    ps.setIndentEnd(0);
    
    ps.setIndentFirstLine(15);
    ps.setIndentStart(21);
    
    ps.setLineSpacing(115);    
    ps.setSpaceAbove(0);
    ps.setSpaceBelow(5);          
}

//later
// function addKbsTKG(msg) {
//     const pres = SlidesApp.getActivePresentation();
//     const allSlides = pres.getSlides();
//     var kbtemp = allSlides[1];
//     var slide = kbtemp.duplicate();
//     slide.getShapes().forEach(shape => {
//         if (shape.getShapeType() == SlidesApp.ShapeType.TEXT_BOX) {
//             switch (shape.getText().asRenderedString().trim()) {
//                 case "TEMPLATE":
//                     shape.remove();
//                     break;
//                 case "Week of xx/xx/2020":
//                     var date = Utilities.formatDate(new Date(), "GMT+1", "MM/dd/yyyy")
//                     shape.getText().setText("Week of " + date)
//                     break;
//             }
//         }
//     }); 
//     const leftside = slide.getShapes().find(shape => {
//         return shape.getText().asRenderedString().trim().startsWith("{{leftside}}");
//     });

//     const rightside = slide.getShapes().find(shape => {
//         return shape.getText().asRenderedString().trim().startsWith("{{rightside}}");
//     });
//     const textRangeLeft = leftside.getText();
//     const textRangeRight = rightside.getText();

//     textRangeLeft.clear();
//     textRangeRight.clear();
    
//     const ss = SpreadsheetApp.openById("1vPou7PjHA-VLkKpQe18uMIGIHzw7IqIYS319j4sjf4w");
//     const sheet = ss.getSheetByName('TKGKBFiltered');
//     const range = sheet.getDataRange();
//     const rows = range.getNumRows();
//     var dateArray: string[] = msg.split("-");
//     const today = new Date(parseInt(dateArray[0]), parseInt(dateArray[1]) -1, parseInt(dateArray[2]));
//     let links:Array<bullet> = [];
//     let linksO = [];
//     let linksS = [];  
//     for (let i = 1; i < rows; i++) {
//         const update = new Date(range.getCell(i, 10).getValue());
//         if (today.getFullYear() == update.getFullYear() && today.getMonth() == update.getMonth() && today.getDate() == update.getDate()) {
//             Logger.log(`${today.getFullYear()} == ${update.getFullYear()} && ${today.getMonth()} == ${update.getMonth()} && ${today.getDate()} == ${update.getDate()}`);
//             const type = range.getCell(i, 11).getValue();
//             const link = "https://community.pivotal.io/s/article/" + range.getCell(i, 8).getValue();
//             const message = range.getCell(i, 9).getValue();
//             if (type.toString().startsWith("O")) {
//                 linksO.push({ title: `${message}`, url: link, type: "B"});
//             } else {
//                 linksS.push({ title: `${message}`, url: link, type: "B"})
//             }
//         }
//     }

//     links = links.concat({ title: "Opsmanager/TAS components", url: "", type: "H" });
//     links = links.concat(linksO);
//     links = links.concat({ title: "Services/Other Tiles", url: "", type: "H" });
//     links = links.concat(linksS);
//     let headers: Array<GoogleAppsScript.Slides.TextRange> = [];
//     let bullets: Array<GoogleAppsScript.Slides.TextRange> = [];
//     let fit = fillSlide(links, textRangeLeft, 0, headers, bullets, leftside);
//     if (fit != 0) {
//         fillSlide(links, textRangeRight, fit, headers, bullets, rightside);
//     }
//     for (const bullet of bullets) {
//         formatShape(bullet);
//     }

//     for (const header of headers) {
//         header.getParagraphStyle().setIndentStart(0).setSpaceAbove(2).setSpaceBelow(3);
//         header.getListStyle().removeFromList();
//         header.getTextStyle().setForegroundColor("#3f3f3f").setBold(true);
//     }    
// }
