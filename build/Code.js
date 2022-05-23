var __values = (this && this.__values) || function(o) {
    var s = typeof Symbol === "function" && Symbol.iterator, m = s && o[s], i = 0;
    if (m) return m.call(o);
    if (o && typeof o.length === "number") return {
        next: function () {
            if (o && i >= o.length) o = void 0;
            return { value: o && o[i++], done: !o };
        }
    };
    throw new TypeError(s ? "Object is not iterable." : "Symbol.iterator is not defined.");
};
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
function testKbs() {
    addKbs("2022-05-02", 5);
}
function addKbs(msg, lines) {
    var e_1, _a, e_2, _b;
    var pres = SlidesApp.getActivePresentation();
    var allSlides = pres.getSlides();
    var kbtemp = allSlides[1];
    var slide = kbtemp.duplicate();
    slide.getShapes().forEach(function (shape) {
        if (shape.getShapeType() == SlidesApp.ShapeType.TEXT_BOX) {
            switch (shape.getText().asRenderedString().trim()) {
                case "TEMPLATE":
                    shape.remove();
                    break;
                case "Week of xx/xx/2020":
                    var date = Utilities.formatDate(new Date(), "GMT+1", "MM/dd/yyyy");
                    shape.getText().setText("Week of " + date);
                    break;
            }
        }
    });
    var leftside = slide.getShapes().find(function (shape) {
        return shape.getText().asRenderedString().trim().startsWith("{{leftside}}");
    });
    var rightside = slide.getShapes().find(function (shape) {
        return shape.getText().asRenderedString().trim().startsWith("{{rightside}}");
    });
    var textRangeLeft = leftside.getText();
    var textRangeRight = rightside.getText();
    textRangeLeft.clear();
    textRangeRight.clear();
    var ss = SpreadsheetApp.openById("1vPou7PjHA-VLkKpQe18uMIGIHzw7IqIYS319j4sjf4w");
    var sheet = ss.getSheetByName('TASKBFiltered');
    var range = sheet.getDataRange();
    var rows = range.getNumRows();
    var dateArray = msg.split("-");
    var today = new Date(parseInt(dateArray[0]), parseInt(dateArray[1]) - 1, parseInt(dateArray[2]));
    var links = [];
    var linksO = [];
    var linksS = [];
    for (var i = 1; i < rows; i++) {
        var update = new Date(range.getCell(i, 10).getValue());
        if (today.getFullYear() == update.getFullYear() && today.getMonth() == update.getMonth() && today.getDate() == update.getDate()) {
            var type = range.getCell(i, 11).getValue();
            var link = "https://community.pivotal.io/s/article/" + range.getCell(i, 8).getValue();
            var message = range.getCell(i, 9).getValue();
            if (type.toString().startsWith("O")) {
                linksO.push({ title: "".concat(message), url: link, type: "B" });
            }
            else {
                linksS.push({ title: "".concat(message), url: link, type: "B" });
            }
        }
    }
    links = links.concat({ title: "Opsmanager/TAS components", url: "", type: "H" });
    links = links.concat(linksO);
    if (linksS.length > 0) {
        links = links.concat({ title: "Services/Other Tiles", url: "", type: "H" });
        links = links.concat(linksS);
    }
    var headers = [];
    var bullets = [];
    var fit = fillSlide(links, textRangeLeft, 0, headers, bullets, leftside, lines);
    Logger.log("length  = ".concat(links.length, " and fit = ").concat(fit, ", "));
    if (fit < links.length) {
        Logger.log("here");
        fillSlide(links, textRangeRight, fit, headers, bullets, rightside, lines + fit);
    }
    try {
        for (var bullets_1 = __values(bullets), bullets_1_1 = bullets_1.next(); !bullets_1_1.done; bullets_1_1 = bullets_1.next()) {
            var bullet = bullets_1_1.value;
            formatShape(bullet);
        }
    }
    catch (e_1_1) { e_1 = { error: e_1_1 }; }
    finally {
        try {
            if (bullets_1_1 && !bullets_1_1.done && (_a = bullets_1["return"])) _a.call(bullets_1);
        }
        finally { if (e_1) throw e_1.error; }
    }
    try {
        for (var headers_1 = __values(headers), headers_1_1 = headers_1.next(); !headers_1_1.done; headers_1_1 = headers_1.next()) {
            var header = headers_1_1.value;
            header.getParagraphStyle().setIndentStart(0).setSpaceAbove(2).setSpaceBelow(3);
            header.getListStyle().removeFromList();
            header.getTextStyle().setForegroundColor("#3f3f3f").setBold(true);
        }
    }
    catch (e_2_1) { e_2 = { error: e_2_1 }; }
    finally {
        try {
            if (headers_1_1 && !headers_1_1.done && (_b = headers_1["return"])) _b.call(headers_1);
        }
        finally { if (e_2) throw e_2.error; }
    }
}
function formatShape(shape) {
    var ps = shape.getParagraphStyle();
    ps.setIndentEnd(0);
    ps.setLineSpacing(115);
    ps.setSpaceAbove(0);
    ps.setSpaceBelow(5);
    ps.setIndentFirstLine(15);
    ps.setIndentStart(21);
}
function fillSlide(links, textRange, maxIndex, headers, bullets, shape, lines) {
    var e_3, _a;
    var rowHeight = 0;
    var fits = true;
    var index = 0;
    try {
        for (var links_1 = __values(links), links_1_1 = links_1.next(); !links_1_1.done; links_1_1 = links_1.next()) {
            var link = links_1_1.value;
            index++;
            if (index <= maxIndex)
                continue;
            var h = void 0;
            switch (link.type) {
                case "B":
                    h = textRange.appendText("".concat(link.title));
                    h.getTextStyle().setLinkUrl(link.url).setForegroundColor("#0091da").setBold(false);
                    bullets.push(h);
                    break;
                case "H":
                    h = textRange.appendText("".concat(link.title));
                    h.getTextStyle().setForegroundColor("#3f3f3f").setBold(true);
                    headers.push(h);
                    break;
            }
            if (index >= lines) {
                fits = false;
                break;
            }
            else {
                h = textRange.appendText("\n");
            }
            //   rowHeight = rowHeight + getHeight(textRange);
            //   if (rowHeight > 190) {
            //     fits = index++;
            //     break;
            //   } else {
            //     h = textRange.appendText("\n");
            //  }
        }
    }
    catch (e_3_1) { e_3 = { error: e_3_1 }; }
    finally {
        try {
            if (links_1_1 && !links_1_1.done && (_a = links_1["return"])) _a.call(links_1);
        }
        finally { if (e_3) throw e_3.error; }
    }
    textRange.getListStyle().applyListPreset(SlidesApp.ListPreset.DISC_CIRCLE_SQUARE);
    return index;
}
function getHeight(tr) {
    var rowHeight = (tr.asString().split("\n").length - 1) * (tr.getTextStyle().getFontSize() / 72 * 96);
    var ps = tr.getParagraphStyle();
    rowHeight = rowHeight + ps.getSpaceAbove();
    rowHeight = rowHeight + ps.getSpaceBelow();
    rowHeight = rowHeight + (ps.getLineSpacing() / 100);
    return rowHeight;
}
function testBullets() {
    var pres = SlidesApp.getActivePresentation();
    var allSlides = pres.getSlides();
    var slide = allSlides[1];
    var leftside = slide.getShapes().find(function (shape) {
        return shape.getText().asRenderedString().trim().startsWith("{{leftside}}");
    });
    var textRange = leftside.getText();
    textRange.clear();
    var ar = [];
    ar.push(textRange.appendText("Heading\n"));
    for (var index = 1; index < 4; index++) {
        ar.push(textRange.appendText("This is a very long bullet to test the paragraph.  This is a very long bullet to test. ".concat(index, "\n")));
    }
    textRange.getListStyle().applyListPreset(SlidesApp.ListPreset.DISC_CIRCLE_SQUARE);
    var ps = ar[0].getParagraphStyle();
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
