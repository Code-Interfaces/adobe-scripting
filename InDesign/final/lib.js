// A series of functions to create shapes and text in InDesign by Alvin Ashiatey
function rect(x, y, width, height){
    var doc = app.activeDocument;
    var page = doc.pages.item(0);
    var rect = page.rectangles.add();
    rect.geometricBounds = [y, x, width, height];
    return rect;
}
function layer(name){
    var doc = app.activeDocument;
    if (doc.layers.itemByName(name).isValid) {
        return doc.layers.itemByName(name);
    }
    var layer = doc.layers.add();
    layer.name = name;
    return layer;
}

function image(src, x, y, width, height){
    var image = rect(x, y, width, height);
    image.place(File(src));
    return image;
}

function yell (msg) {
    alert(msg);
}

function line(x1, y1, x2, y2){
    var doc = app.activeDocument;
    var page = doc.pages.item(0);
    var line = page.graphicLines.add();
    line.paths[0].pathPoints[0].anchor = [x1, y1];
    line.paths[0].pathPoints[1].anchor = [x2, y2];
    return line;
}

function circle(x, y, radius){
    var doc = app.activeDocument;
    var page = doc.pages.item(0);
    var circle = page.ovals.add();
    circle.geometricBounds = [y - radius, x - radius, y + radius, x + radius];
    return circle;
}

function ellipse(x, y, width, height){
    var doc = app.activeDocument;
    var page = doc.pages.item(0);
    var ellipse = page.ovals.add();
    ellipse.geometricBounds = [y - height/2, x - width/2, y + height/2, x + width/2];
    return ellipse;
}

function polygon(x, y, radius, sides){
    var doc = app.activeDocument;
    var page = doc.pages.item(0);
    var polygon = page.polygons.add();
    polygon.properties = {
        geometricBounds: [y - radius, x - radius, y + radius, x + radius],
        numberOfSides: sides
    };
    return polygon;
}

function triangle(x1, y1, x2, y2, x3, y3){
    return polygon([x1, y1, x2, y2, x3, y3], 3);
}

function roundedRectangle(x, y, width, height, cornerRadius){
    var doc = app.activeDocument;
    var page = doc.pages.item(0);
    var rect = page.rectangles.add();
    rect.geometricBounds = [y, x, y + height, x + width];
    rect.cornerRadius = cornerRadius;
    return rect;
}

function setStroke(shape, weight, color){
    shape.strokeWeight = weight;
    shape.strokeColor = color;
}

function hexToRGB(hex) {
    var r = parseInt(hex.substring(1, 3), 16);
    var g = parseInt(hex.substring(3, 5), 16);
    var b = parseInt(hex.substring(5, 7), 16);
    return [r, g, b];
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

function map(val, min, max, newMin, newMax){
    return (val - min) * (newMax - newMin) / (max - min) + newMin;
}

function textFrame(x, y, width, height, text){
    var doc = app.activeDocument;
    var page = doc.pages.item(0);
    var textFrame = page.textFrames.add();
    textFrame.geometricBounds = [y, x, height, width];
    textFrame.contents = text;
    return textFrame;
}



