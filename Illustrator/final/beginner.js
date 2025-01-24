// alert("Hello, World!");

if (app.documents.length > 0) {
    var doc = app.activeDocument;
    var artBoard = doc.artboards[0];
    var artBoardRect = artBoard.artboardRect; // Format [left, top, right, bottom]

    var width = artBoardRect[2] - artBoardRect[0]; // right - left
    var height = artBoardRect[1] - artBoardRect[3]; // top - bottom

    var rows = 5;
    var cols = 5;

    var scaleDown = 1.2;

    var gridWidth = (width / cols) / scaleDown;
    var gridHeight = (height / rows) / scaleDown;


    var placedItems = [];

    for (var row = 0; row < rows; row++) {
        for (var col = 0; col < cols; col++) {

            var group = doc.groupItems.add();


            var rect_1 = doc.pathItems.rectangle(0, 0, gridWidth, gridHeight);
            var white = new RGBColor();
            white.red = 245;
            white.green = 245;
            white.blue = 245;
            rect_1.fillColor = white;
            rect_1.position = [col * gridWidth + (gridWidth/2), row * gridHeight - (gridHeight/2)];
            rect_1.move(group, ElementPlacement.PLACEATEND);

            var random = Math.random()

            var rect_2;
            var rect_2_centerX;
            var rect_2_centerY;

            var rect_size = 20;

            if (random > 0.5) {
                rect_2 = doc.pathItems.rectangle(0, 0, rect_size, gridHeight);
                rect_2_centerX = rect_1.position[0] + rect_1.width/2 - rect_2.width/2;
                rect_2_centerY = rect_1.position[1];
            } else {
                rect_2 = doc.pathItems.rectangle(0, 0, gridWidth, rect_size);
                rect_2_centerX = rect_1.position[0];
                rect_2_centerY = rect_1.position[1] - rect_1.height/2 + rect_2.height/2;
            }

            var red = new RGBColor();
            red.red = 255;
            red.green = 0;
            red.blue = 0;
            rect_2.fillColor = red;
            rect_2.position = [rect_2_centerX, rect_2_centerY];

            rect_2.move(group, ElementPlacement.PLACEATBEGINNING);
            placedItems.push(group);


        }
    }

    var finalGroup = doc.groupItems.add();
    for (var i = 0; i < placedItems.length; i++) {
        placedItems[i].move(finalGroup, ElementPlacement.PLACEATEND);
    }

    var artBoardCenterX = artBoardRect[0] + width / 2;
    var artBoardCenterY = artBoardRect[1] - height / 2;

    var groupCenterX = finalGroup.left + finalGroup.width / 2;
    var groupCenterY = finalGroup.top - finalGroup.height / 2;

    var deltaX = artBoardCenterX - groupCenterX;
    var deltaY = artBoardCenterY - groupCenterY;

    finalGroup.translate(deltaX, deltaY);

}