// This is a simple script that alerts "Hello, World!" in InDesign.

//  We can use the alert() function to display a message in a dialog box.
// alert("Hello, World!");

if (app.documents.length > 0) {
    var doc = app.activeDocument;
    var page = doc.pages.item(0);

    var margin = page.marginPreferences;

    var width = doc.documentPreferences.pageWidth - margin.left - margin.right;
    var height = doc.documentPreferences.pageHeight - margin.top - margin.bottom;

    var rows = 5;
    var cols = 5;

    var gridWidth = (width / cols);
    var gridHeight = (height / rows);

    for (var row = 0; row < rows; row++) {
        for (var col = 0; col < cols; col++) {

            var x = col * gridWidth  + margin.left;
            var y = row * gridHeight + margin.top;

            var smokeWhite = doc.colors.item("Smoke White");
            if (!smokeWhite.isValid) {
                smokeWhite = doc.colors.add();
                smokeWhite.properties = {
                    name: "Smoke White",
                    model: ColorModel.PROCESS,
                    space: ColorSpace.RGB,
                    colorValue: [245, 245, 245]
                };
            }

            var rect_1 = page.rectangles.add({
                geometricBounds: [y, x, y + gridHeight, x + gridWidth],
                fillColor: smokeWhite,
                strokeColor: doc.swatches.item("None")
            });

            var random = Math.random();
            var rect_2;
            var rect_size = 2;

            var red = doc.colors.item("Red");
            if (!red.isValid) {
                red = doc.colors.add();
                red.properties = {
                    name: "Red",
                    model: ColorModel.PROCESS,
                    space: ColorSpace.RGB,
                    colorValue: [255, 0, 0]
                };
            }

            if (random > 0.5) {
                rect_2 = page.rectangles.add({
                    geometricBounds: [y, x + gridWidth / 2 - rect_size / 2, y + gridHeight, x + gridWidth / 2 + rect_size / 2],
                    fillColor: red,
                    strokeColor: doc.swatches.item("None")
                });
            } else {
                rect_2 = page.rectangles.add({
                    geometricBounds: [y + gridHeight / 2 - rect_size / 2, x, y + gridHeight / 2 + rect_size / 2, x + gridWidth],
                    fillColor: red,
                    strokeColor: doc.swatches.item("None")
                });
            }

        }
    }
} else {
    alert("No open document found");
}
