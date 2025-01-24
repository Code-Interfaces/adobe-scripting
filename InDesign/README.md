# INDESIGN SCRIPTING

Most adobe applications have a scripting interface. This allows you to automate tasks and create custom workflows.
Each application exposes an API that you can use to interact with the application. 

## Hello World

Let's start with a simple script that alerts "Hello World" when run.

```javascript
alert("Hello World");
```

Now let's try a more complex script that creates a new document and adds a text frame to it.

```javascript
var doc = app.documents.add();
var textFrame = doc.pages[0].textFrames.add();
textFrame.geometricBounds = [0, 0, 100, 100];
textFrame.contents = "Hello World";
```

## Vera Molnar

Lets try to recreate a piece by Vera Molnar using InDesign scripting.

![Int/Cont (1996)](https://i0.wp.com/dam.org/museum/wp-content/uploads/2020/08/Molnar-IntCont-4-scaled.jpg?resize=1024%2C1024&ssl=1)

Lets start by checking if there are any open documents if not we alert the user.

```javascript
if (app.documents.length > 0) {
     // code goes here
} else {
    alert("No open document found");
}
```

If there is an open document, we start by assigning the active document to a variable and then selecting the first page.

```javascript
var doc = app.activeDocument;
var page = doc.pages.item(0);
```

Now lets get the margins of the page so we can calculate our work area.

```javascript
var margin = page.marginPreferences;
var width = doc.documentPreferences.pageWidth - margin.left - margin.right;
var height = doc.documentPreferences.pageHeight - margin.top - margin.bottom;
```

For this piece, there is a 2 x 2 grid in which a red rectangle is centered. First, we will define the variables for the rows and columns, and then calculate the grid's width and height.

```javascript
var rows = 2;
var cols = 2;

var gridWidth = (width / cols);
var gridHeight = (height / rows);
```

Now we can create the rectangles and position them in the grid.

```javascript
for (var row = 0; row < rows; row++) {
        for (var col = 0; col < cols; col++) {

            var x = col * gridWidth  + margin.left;
            var y = row * gridHeight + margin.top;

            // Here we create a color swatch for the background of the first rectangle
            var smokeWhite = doc.colors.item("Smoke White");
            // We check if the color swatch exists, if not we create it
            if (!smokeWhite.isValid) {
                smokeWhite = doc.colors.add();
                smokeWhite.properties = {
                    name: "Smoke White",
                    model: ColorModel.PROCESS,
                    space: ColorSpace.RGB,
                    colorValue: [245, 245, 245]
                };
            }

            // Here we create the first rectangle and set its properties
            // We use the geometricBounds property to set the position and size of the rectangle
            // We use the fillColor property to set the color of the rectangle
            // We use the strokeColor property to set the color of the stroke
            var rect_1 = page.rectangles.add({
                geometricBounds: [y, x, y + gridHeight, x + gridWidth],
                fillColor: smokeWhite,
                strokeColor: doc.swatches.item("None")
            });

            // Here is where we introduce the random element
            // We use the Math.random() function to generate a random number between 0 and 1
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
            
            // Here we use the random number to decide if the second rectangle will be horizontal or vertical
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
```
