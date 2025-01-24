#include "lib.js"

yell("Hello World");
var hex = [
    "#FFC600",
    "#914E72",
    "#00A95C",
    "#F15060",
    "#765BA7",
    "#FFE800",
    "#FF48B0",
    "#5EC8E5",
    "#FF4C65",
];

// Replace with your image path
var imagePath = "/Users/alvinkwabena/Documents/Alvin Ashiatey/Class/Code and Interfaces - ART 754/Workshops/Adobe scripting/InDesign/final/img.png";


function main() {
    if (app.activeDocument) {
        var document = app.activeDocument;
        var page = document.pages.item(0);
        var layer;
        var layerName = "Code & Interfaces";

        if (document.layers.item(layerName) == null) {
            layer = document.layers.add({ name: layerName});
        } else {
            layer = document.layers.item(layerName);
        }

        rectangle("#fefefe")
        placeImageIntoRectangle(imagePath);

        var randomFontIndex = Math.floor(Math.random() * app.fonts.length);
        var font = app.fonts[randomFontIndex];

        var topTextFrame = createConfiguredTextFrame(page, layer, "Code & Interfaces - ART 754", font);

        // https://www.indesignjs.de/extendscriptAPI/indesign-latest/#AutoSizingReferenceEnum.html
        topTextFrame.textFramePreferences.autoSizingReferencePoint = AutoSizingReferenceEnum.TOP_LEFT_POINT;
        topTextFrame.textFramePreferences.autoSizingType = AutoSizingTypeEnum.HEIGHT_ONLY;
        // https://www.indesignjs.de/extendscriptAPI/indesign-latest/#Justification.html
        topTextFrame.paragraphs.everyItem().justification = Justification.FULLY_JUSTIFIED;

        var bottomTextFrame = createConfiguredTextFrame(page, layer, "02.01.2024", font);

        // https://www.indesignjs.de/extendscriptAPI/indesign-latest/#AutoSizingReferenceEnum.html
        bottomTextFrame.textFramePreferences.autoSizingReferencePoint = AutoSizingReferenceEnum.BOTTOM_LEFT_POINT;
        bottomTextFrame.textFramePreferences.autoSizingType = AutoSizingTypeEnum.HEIGHT_ONLY;
        // https://www.indesignjs.de/extendscriptAPI/indesign-latest/#Justification.html
        bottomTextFrame.paragraphs.everyItem().justification = Justification.CENTER_ALIGN;

    }
}

function createConfiguredTextFrame(page, layer, textContent, font) {
    var textFrame = page.textFrames.add(layer);
    // https://www.indesignjs.de/extendscriptAPI/indesign-latest/#VerticalJustification.html
    textFrame.textFramePreferences.verticalJustification =
        VerticalJustification.TOP_ALIGN;
    textFrame.contents = textContent.toUpperCase();

    // Retrieve margin preferences
    var margins = page.marginPreferences;

    textFrame.visibleBounds = [
        margins.top,
        margins.left,
        page.bounds[2] - margins.bottom,
        page.bounds[3] - margins.right
    ];

    textFrame.paragraphs.everyItem().appliedFont = font;
    textFrame.paragraphs.everyItem().hyphenation = false;

    var fontSize = 120;
    textFrame.texts[0].pointSize = fontSize;
    textFrame.paragraphs.everyItem().fillColor = colors("#fefefe");
    // https://www.indesignjs.de/extendscriptAPI/indesign-latest/#Justification.html
    textFrame.paragraphs.everyItem().justification = Justification.FULLY_JUSTIFIED;
    textFrame.paragraphs.everyItem().leading = (fontSize / 120) * 100;
    return textFrame;
}

function colors(clr) {
    var doc = app.activeDocument;
    var color
    try {
        color = doc.colors.item(clr);
        color.properties = {
            model: ColorModel.PROCESS,
            space: ColorSpace.RGB,
            colorValue: hexToRGB(clr)
        };
    } catch (e) {
        color = doc.colors.add({
            name: clr,
            model: ColorModel.PROCESS,
            space: ColorSpace.RGB,
            colorValue: hexToRGB(clr)
        });
    }
    return color;
}

function hexToRGB(hex) {
    var r = parseInt(hex.substring(1, 3), 16);
    var g = parseInt(hex.substring(3, 5), 16);
    var b = parseInt(hex.substring(5, 7), 16);
    return [r, g, b];
}

function rectangle(hex){
    var doc = app.activeDocument;
    var page = doc.pages.item(0);
    var rect = page.rectangles.add();
    var margins = page.marginPreferences;
    rect.visibleBounds = [
        margins.top,
        margins.left,
        page.bounds[2] - margins.bottom,
        page.bounds[3] - margins.right
    ];
    rect.fillColor = colors(hex);
}

function placeImageIntoRectangle(imagePath) {
    if (app.documents.length === 0) {
        alert("Please open a document first.");
        return;
    }

    var document = app.activeDocument;
    var page = document.pages.item(0);

    // Create a rectangle
    var rect = page.rectangles.add();
    rect.geometricBounds = page.bounds;

    // Place the image inside the rectangle
    var image = rect.place(File(imagePath))[0];

    // https://https://www.indesignjs.de/extendscriptAPI/indesign-latest/#FitOptions.html
    rect.fit(FitOptions.FILL_PROPORTIONALLY);

    // https://www.indesignjs.de/extendscriptAPI/indesign-latest/#BlendMode.html
    rect.transparencySettings.blendingSettings.blendMode = BlendMode.LUMINOSITY;

}

app.doScript(main, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, "Hello World");