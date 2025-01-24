# ILLUSTRATOR SCRIPTING

A lot of what we learned in the InDesign scripting workshop can be applied to Illustrator scripting. The main difference is the object model and the methods available to us. Let's take a look at some basic scripting examples for Illustrator.

## Hello World

Let's start with a simple script that alerts "Hello World" when run.

```javascript
alert("Hello World");

```

Now let's try a more complex script that creates a new document and adds a rectangle to it.

```javascript
var doc = app.documents.add();
var rect = doc.pathItems.rectangle(100, 100, 100, 100);
```

## Vera Molnar

Let's try to recreate a piece by Vera Molnar using Illustrator scripting.

![Int/Cont (1996)](https://i0.wp.com/dam.org/museum/wp-content/uploads/2020/08/Molnar-IntCont-4-scaled.jpg?resize=1024%2C1024&ssl=1)

Let's start by checking if there are any open documents if not we alert the user.

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
var artboard = doc.artboards[0];
var artboardRect = artboard.artboardRect;
```

Now, let's get the width and height of the artboard so that we can calculate our work area. The `artboardRect` property returns an array with the coordinates of the artboard's bounding box in the format `[y1, x1, y2, x2]`. So to get the width and height we can do the following:

```javascript
var width = artboardRect[2] - artboardRect[0];
var height = artboardRect[1] - artboardRect[3];
```

For this piece, there is a 2 x 2 grid in which a red rectangle is centered. First, we will define the variables for the rows and columns, and then calculate the grid's width and height. For this example, we will introduce a margin to the artboard.

```javascript
var rows = 2;
var cols = 2;

var margin = 1.2;

var gridWidth = (width / cols) / margin;
var gridHeight = (height / rows) / margin;
```

Now we can create the rectangles and position them in the grid.

```javascript
var allRects = [];

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
```

Lets break down the code:

1. We create a group for each row and column.
2. We create the first rectangle and set its properties.
3. We introduce the random element and create the second rectangle.
4. We set the color of the second rectangle and position it in the grid.
5. We add the rectangles to the group.
6. We add the group to the `placedItems` array.
7. We create a final group and add all the groups to it.
8. We calculate the center of the artboard and the center of the final group.
9. We calculate the difference between the two centers and translate the final group to the center of the artboard.



