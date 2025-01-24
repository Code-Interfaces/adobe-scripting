// Adobe Illustrator Script to place an image in a grid on a Tabloid page, scaled proportionally to fit the grid cell, and positioned flush with the top-left corner

// Replace with your own image path
var imagePath = "/Users/alvinkwabena/Documents/Alvin Ashiatey/Class/Code and Interfaces - ART 754/Workshops/Adobe scripting/Illustrator/final/img.png";

function placeImagesInGrid(rows, cols) {
    var doc = app.activeDocument;
    var artBoard = doc.artboards[0];
    var artBoardRect = artBoard.artboardRect; // Format [left, top, right, bottom]

    var width = artBoardRect[2] - artBoardRect[0]; // right - left
    var height = artBoardRect[1] - artBoardRect[3]; // top - bottom

    var gridWidth = width / cols;
    var gridHeight = width / cols;

    var placedItems = [];
    var scaleIncrement = 0;

    for (var row = 0; row < rows; row++) {
        for (var col = 0; col < cols; col++) {

            var group = app.activeDocument.groupItems.add();

            // Place the image
            var placedItem = doc.placedItems.add();
            placedItem.file = new File(imagePath);

            // Calculate proportional scale to fit the grid cell
            placedItem.width = gridWidth;
            placedItem.height = gridHeight;

            var horizontalScale = gridWidth / placedItem.width * 100;
            var verticalScale = gridHeight / placedItem.height * 100;

            var scale = Math.min(horizontalScale, verticalScale) + scaleIncrement;
            placedItem.resize(scale, scale, true); // 'true' maintains the proportions

            var posX = artBoardRect[0] + col * gridWidth;
            var posY = artBoardRect[1] - row * gridHeight;
            var adjustedPosY = posY - (gridHeight - placedItem.height * (scale - scaleIncrement) / 100);

            var maskX = posX;
            var maskY = posY - gridHeight;

            // Calculate new position to center the image within the mask
            var centeredPosX = maskX + (gridWidth - placedItem.width) / 2;
            var centeredPosY = maskY + (gridHeight + placedItem.height) / 2;

            // Apply the new position
            placedItem.position = [centeredPosX, centeredPosY];

            // Masking the image
            var mask = doc.pathItems.rectangle(posY, posX, gridWidth, gridHeight);
            mask.stroked = false;
            mask.filled = false;

            mask.move(group, ElementPlacement.PLACEATBEGINNING);
            placedItem.move(group, ElementPlacement.PLACEATEND);
            group.clipped = true;

            placedItems.push(group);
            scaleIncrement += 10;
        }
    }

    // Grouping the images
    var group = doc.groupItems.add();
    for (var i = 0; i < placedItems.length; i++) {
        placedItems[i].move(group, ElementPlacement.PLACEATEND);
    }
}

function addCenteredTextAtBottom(textContent) {
    var doc = app.activeDocument;
    var artBoard = doc.artboards[0];
    var artBoardRect = artBoard.artboardRect;

    // Create text frame
    var textFrame = doc.textFrames.add();
    textFrame.contents = textContent;

    // Set text properties
    textFrame.textRange.paragraphAttributes.justification = Justification.CENTER;
    textFrame.textRange.characterAttributes.size = 52; // Set text size (example: 12pt)

    // Calculate position
    var textWidth = textFrame.width;
    var textHeight = textFrame.height;
    var bottomMargin = 55; // You can adjust this value for the margin
    var posX = (artBoardRect[0] + artBoardRect[2] - textWidth) / 2; // Center horizontally
    var posY = artBoardRect[3] + textHeight + bottomMargin; // Position at the bottom

    // Apply position
    textFrame.position = [posX, posY];
}


var rows = 8;
var cols = 7;

if (app.documents.length > 0) {
    placeImagesInGrid(rows, cols); // Adjust rows and columns as needed
    addCenteredTextAtBottom("Code & Interfaces - ART 754");
} else {
    var doc = app.documents.add(DocumentColorSpace.CMYK, 792, 1224); // Tabloid size in points
    placeImagesInGrid(rows, cols); // Adjust rows and columns as needed
    addCenteredTextAtBottom("Code & Interfaces - ART 754");
}
