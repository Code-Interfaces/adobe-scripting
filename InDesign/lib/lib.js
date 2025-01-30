// A series of functions to create shapes and text in InDesign by Alvin Ashiatey

// Global variables to store the width and height of the page
var _width, _height, _marginWidth, _marginHeight;

/**
 * Gets the active document or creates a new one if none exists
 * @returns {Document} The active InDesign document
 * @private
 */
function _doc() {
  return app.activeDocument;
}

/**
 * Sets the page size of the document
 * @param {Document} doc - The InDesign document
 * @param {string} size - The desired page size ("letter", "tabloid", etc.)
 * @returns {boolean} - Success status
 * @example
 * // Set the page size to letter
 * _setPageSize(_doc(), "letter");
 */
function _setPageSize(doc, size) {
  const pageSizes = {
    letter: [612, 792], // 8.5 x 11 inches in points
    tabloid: [792, 1224], // 11 x 17 inches in points
    a3: [841.89, 1190.551], // 297 x 420 mm in points
    a4: [595.276, 841.89], // 210 x 297 mm in points
    a5: [419.528, 595.276], // 148 x 210 mm in points
    // Add more sizes as needed
  };

  if (!pageSizes[size.toLowerCase()]) {
    alert("Invalid page size: " + size);
    return false;
  }

  try {
    var page = doc.pages.item(0);
    page.resize(
      CoordinateSpaces.INNER_COORDINATES,
      AnchorPoint.TOP_LEFT_ANCHOR,
      ResizeMethods.REPLACING_CURRENT_DIMENSIONS_WITH,
      pageSizes[size]
    );
    _width = pageSizes[size][0];
    _height = pageSizes[size][1];
    _marginWidth = _width - _getBounds(doc).left - _getBounds(doc).right;
    _marginHeight = _height - _getBounds(doc).top - _getBounds(doc).bottom;
    return true;
  } catch (e) {
    alert("Error setting page size: " + e.message);
    return false;
  }
}

/**
 * Sets the page size of the document by specific dimensions
 * @param {Document} doc - The InDesign document
 * @param {number} width - The width of the page (in points)
 * @param {number} height - The height of the page (in points)
 * @returns {boolean} - Success status
 * @example
 * // Set the page size to 8.5 x 11 inches (612 x 792 points)
 * _setPageSizeByDimensions(_doc(), 612, 792);
 */
function _setPageSizeByDimensions(doc, width, height) {
  try {
    var page = doc.pages.item(0);
    page.resize(
      CoordinateSpaces.INNER_COORDINATES,
      AnchorPoint.TOP_LEFT_ANCHOR,
      ResizeMethods.REPLACING_CURRENT_DIMENSIONS_WITH,
      [width, height]
    );
    _width = width;
    _height = height;
    _marginWidth = _width - _getBounds(doc).left - _getBounds(doc).right;
    _marginHeight = _height - _getBounds(doc).top - _getBounds(doc).bottom;
    return true;
  } catch (e) {
    alert("Error setting page size: " + e.message);
    return false;
  }
}

/**
 * Gets the margin settings of the document
 * @param {Document} doc - The InDesign document
 * @returns {Object} - An object containing the margin settings (top, left, bottom, right)
 * @example
 * var margins = _margin(_doc());
 * alert("Top margin: " + margins.top);
 */
function _getBounds(doc) {
  var page = doc.pages.item(0);
  var marginPreferences = page.marginPreferences;
  return {
    top: marginPreferences.top,
    left: marginPreferences.left,
    bottom: marginPreferences.bottom,
    right: marginPreferences.right,
  };
}

/**
 * Clears the document by removing all pages except the first one and clearing its contents
 * @param {Document} doc - The InDesign document to clear
 * @returns {boolean} - Success status
 */
function _clear(doc) {
  try {
    // Helper function to remove all items from a collection
    function removeAllItems(collection) {
      while (collection.length > 0) {
        collection.item(0).remove();
      }
    }

    // Remove all pages except the first one
    while (doc.pages.length > 1) {
      doc.pages.item(1).remove();
    }

    // Clear the contents of the first page
    removeAllItems(doc.pages.item(0).textFrames);

    // Remove all layers except the default layer
    while (doc.layers.length > 1) {
      doc.layers.item(1).remove();
    }

    // Clear the default layer
    removeAllItems(doc.layers.item(0).pageItems);

    return true;
  } catch (e) {
    alert("Error clearing document: " + e.message);
    return false;
  }
}

/**
 * Creates a rectangle on the first page of the active document
 * @param {number} x - The x-coordinate of the rectangle's left edge (in current document units)
 * @param {number} y - The y-coordinate of the rectangle's top edge (in current document units)
 * @param {number} width - The width of the rectangle (in current document units)
 * @param {number} height - The height of the rectangle (in current document units)
 * @returns {Object} The created rectangle object with methods for styling
 * @example
 * // Create a rectangle at (10, 10) with width 100 and height 50
 * var myRect = _rect(10, 10, 110, 60);
 */
function _rect(x, y, width, height) {
  var doc = _doc();
  var page = doc.pages.item(0);
  var rect = page.rectangles.add();
  rect.geometricBounds = [y, x, height, width];
  return {
    rect: rect,
    fillColor: function (color) {
      _setFillColor(rect, color);
    },
    stroke: function (weight, color) {
      _setStroke(rect, weight, color);
    },
    noStroke: function () {
      _noStroke(rect);
    },
    centerTo: function (keyObject) {
      _centerTo(rect, keyObject);
    },
    flip: function (direction) {
      _flip(rect, direction);
    },
    rotate: function (angle) {
      _rotate(rect, angle);
    },
  };
}

/**
 * Creates a rectangle with rounded corners
 * @param {number} x - The x-coordinate of the rectangle's left edge
 * @param {number} y - The y-coordinate of the rectangle's top edge
 * @param {number} width - The width of the rectangle
 * @param {number} height - The height of the rectangle
 * @param {number} cornerRadius - The radius of the rounded corners
 * @returns {Rectangle} The created rectangle object
 * @example
 * // Create a rounded rectangle at (10,10) size 100x50 with 5pt corner radius
 * var myRect = _roundedRectangle(10, 10, 100, 50, 5);
 */
function _roundedRectangle(x, y, width, height, cornerRadius) {
  var rect = _rect(x, y, width, height);
  _cornerOption(rect, CornerOptions.ROUNDED_CORNER, cornerRadius);
  return rect;
}

/**
 * Sets the corner options for a shape
 * @param {PageItem} shape - The InDesign shape object
 * @param {CornerOptions} cornerType - The type of corner (e.g., CornerOptions.ROUNDED_CORNER)
 * @param {number} cornerSize - The size of the corner radius
 * @example
 * // Set rounded corners with a radius of 5 points
 * _cornerOption(myShape, CornerOptions.ROUNDED_CORNER, 5);
 */
function _cornerOption(shape, cornerType, cornerSize) {
  var corners = ["bottomLeft", "bottomRight", "topLeft", "topRight"];
  for (var i = 0; i < corners.length; i++) {
    shape[corners[i] + "CornerOption"] = cornerType;
    shape[corners[i] + "CornerRadius"] = cornerSize;
  }
}

/**
 * Creates a square with optional rounded corners
 * @param {number} x - The x-coordinate of the square's left edge (in current document units)
 * @param {number} y - The y-coordinate of the square's top edge (in current document units)
 * @param {number} sideLength - The length of each side of the square (in current document units)
 * @param {number} [cornerRadius=0] - The radius of the rounded corners (in current document units)
 * @returns {Rectangle} The created square object
 * @example
 * // Create a square at (10, 10) with sides of 100 and corner radius of 5
 * var mySquare = _square(10, 10, 100, 5);
 */
function _square(x, y, sideLength) {
  return _rect(x, y, sideLength, sideLength);
}

/**
 * Creates or retrieves a layer by name
 * @param {string} name - The name of the layer to create or retrieve
 * @returns {Layer} The layer object (either existing or newly created)
 * @example
 * // Create or get a layer named "Background"
 * var bgLayer = _layer("Background");
 */
function _layer(name) {
  var doc = _doc();
  if (doc.layers.itemByName(name).isValid) {
    return doc.layers.itemByName(name);
  }
  var layer = doc.layers.add();
  layer.name = name;
  return layer;
}

/**
 * Displays an alert message
 * @param {string} msg - The message to display in the alert dialog
 * @example
 * _yell("Hello World!");
 */
function _yell(msg) {
  alert(msg);
}

/**
 * Creates a straight line between two points
 * @param {number} x1 - Starting x-coordinate
 * @param {number} y1 - Starting y-coordinate
 * @param {number} x2 - Ending x-coordinate
 * @param {number} y2 - Ending y-coordinate
 * @returns {object} The created line object with methods for styling
 * @example
 * // Create a diagonal line from (0,0) to (100,100)
 * var myLine = _line(0, 0, 100, 100);
 */
function _line(x1, y1, x2, y2) {
  var doc = _doc();
  var page = doc.pages.item(0);
  var line = page.graphicLines.add();
  line.paths[0].pathPoints[0].anchor = [x1, y1];
  line.paths[0].pathPoints[1].anchor = [x2, y2];
  return {
    line: line,
    stroke: function (weight, color) {
      _setStroke(line, weight, color);
    },
    noStroke: function () {
      _noStroke(line);
    },
    centerTo: function (keyObject) {
      _centerTo(line, keyObject);
    },
    flip: function (direction) {
      _flip(line, direction);
    },
    rotate: function (angle) {
      _rotate(line, angle);
    },
  };
}

/**
 * Creates a circle with specified center point and radius
 * @param {number} x - X-coordinate of the circle's center
 * @param {number} y - Y-coordinate of the circle's center
 * @param {number} radius - Radius of the circle
 * @returns {object} The created circle object with methods for styling
 * @example
 * // Create a circle centered at (100,100) with radius 50
 * var myCircle = _circle(100, 100, 50);
 */
function _circle(x, y, radius) {
  var doc = _doc();
  var page = doc.pages.item(0);
  var circle = page.ovals.add();
  circle.geometricBounds = [y, x, y + radius, x + radius];
  return {
    circle: circle,
    noStroke: function () {
      _noStroke(circle);
    },
    fillColor: function (color) {
      _setFillColor(circle, color);
    },
    centerTo: function (keyObject) {
      _centerTo(circle, keyObject);
    },
  };
}

/**
 * Creates an ellipse with specified center point and dimensions
 * @param {number} x - X-coordinate of the ellipse's center
 * @param {number} y - Y-coordinate of the ellipse's center
 * @param {number} width - Width of the ellipse
 * @param {number} height - Height of the ellipse
 * @returns {object} The created ellipse object with methods for styling
 * @example
 * // Create an ellipse centered at (100,100) with width 200 and height 100
 * var myEllipse = _ellipse(100, 100, 200, 100);
 */
function _ellipse(x, y, width, height) {
  var doc = _doc();
  var page = doc.pages.item(0);
  var ellipse = page.ovals.add();
  ellipse.geometricBounds = [
    y - height / 2,
    x - width / 2,
    y + height / 2,
    x + width / 2,
  ];
  return {
    ellipse: ellipse,
    noStroke: function () {
      _noStroke(ellipse);
    },
    fillColor: function (color) {
      _setFillColor(ellipse, color);
    },
    centerTo: function (keyObject) {
      _centerTo(ellipse, keyObject);
    },
    flip: function (direction) {
      _flip(ellipse, direction);
    },
    rotate: function (angle) {
      _rotate(ellipse, angle);
    },
  };
}

/**
 * Creates a polygon on the active page
 * @param {Number} x - X coordinate for top-left corner
 * @param {Number} y - Y coordinate for top-left corner
 * @param {Number} width - Width of the polygon
 * @param {Number} height - Height of the polygon
 * @param {Number} sides - Number of sides (3 or greater)
 * @param {Number} [cornerRadius] - Radius for rounded corners (optional)
 * @param {Number} [starInset] - Creates a star shape when > 0 (0-100%) (optional)
 * @param {Boolean} [reversed] - Whether to reverse the polygon orientation (optional)
 * @returns {Polygon} The created polygon object
 * @example
 * // Create a hexagon
 * var hex = _polygon(100, 100, 100, 100, 6);
 * // Create a star
 * var star = _polygon(100, 100, 100, 100, 5, 0, 50);
 */
function _polygon(
  x,
  y,
  width,
  height,
  sides,
  cornerRadius,
  starInset,
  reversed
) {
  try {
    // Set defaults for optional parameters
    cornerRadius = cornerRadius || 0;
    starInset = starInset || 0;
    reversed = reversed || false;

    // Validate inputs
    if (sides < 3) {
      throw new Error("Number of sides must be 3 or greater");
    }
    if (width <= 0 || height <= 0) {
      throw new Error("Width and height must be greater than 0");
    }
    if (cornerRadius < 0) {
      throw new Error("Corner radius cannot be negative");
    }
    if (starInset < 0 || starInset > 100) {
      throw new Error("Star inset must be between 0 and 100");
    }

    var doc = _doc();
    var page = doc.layoutWindows[0].activePage;

    // Create the polygon
    var polygon = page.polygons.add({
      geometricBounds: [y, x, y + height, x + width],
      numberOfSides: sides,
      insetPercentage: starInset,
      cornerRadius: cornerRadius,
      reversed: reversed,
    });

    return {
      polygon: polygon,
      noStroke: function () {
        _noStroke(polygon);
      },
      fillColor: function (color) {
        _setFillColor(polygon, color);
      },
      centerTo: function (keyObject) {
        _centerTo(polygon, keyObject);
      },
      flip: function (direction) {
        _flip(polygon, direction);
      },
      rotate: function (angle) {
        _rotate(polygon, angle);
      },
    };
  } catch (e) {
    alert("Error creating polygon: " + e.message);
    return null;
  }
}

/**
 * Creates a custom polygon on the active page using specified points
 * @param {Array<Array<number>>} points - An array of points where each point is an array of [x, y] coordinates
 * @returns {Polygon} The created polygon object
 * @example
 * // Create a custom polygon
 * var points = [
 *   [100, 100],
 *   [150, 100],
 *   [200, 200],
 *   [120, 150]
 * ];
 * var customPolygon = _polygonCustom(points);
 */
function _polygonCustom(points) {
  try {
    if (points.length < 3) {
      throw new Error("A polygon must have at least 3 points");
    }

    var doc = _doc();
    var page = doc.layoutWindows[0].activePage;

    // Create the polygon
    var polygon = page.polygons.add();
    polygon.paths[0].entirePath = points;

    return {
      polygon: polygon,
      noStroke: function () {
        _noStroke(polygon);
      },
      fillColor: function (color) {
        _setFillColor(polygon, color);
      },
      centerTo: function (keyObject) {
        _centerTo(polygon, keyObject);
      },
      flip: function (direction) {
        _flip(polygon, direction);
      },
      rotate: function (angle) {
        _rotate(polygon, angle);
      },
    };
  } catch (e) {
    alert("Error creating custom polygon: " + e.message);
    return null;
  }
}

/**
 * Creates an equilateral triangle on the active page
 * @param {Number} x - X coordinate for top-left corner
 * @param {Number} y - Y coordinate for top-left corner
 * @param {Number} size - Size of the triangle (width/height will be equal)
 * @param {Number} [cornerRadius] - Radius for rounded corners (optional)
 * @param {Boolean} [reversed] - Whether to reverse the triangle orientation (optional)
 * @returns {Polygon} The created triangle object
 * @example
 * // Create a basic triangle
 * var tri = _triangle(100, 100, 50);
 * // Create a rounded triangle
 * var roundTri = _triangle(100, 100, 50, 5);
 */
function _triangle(x, y, size, cornerRadius, reversed) {
  return _polygon(x, y, size, size, 3, cornerRadius || 0, 0, reversed || false);
}

/**
 * Creates a right-angle triangle on the active page
 * @param {number} x - X coordinate for the top-left corner
 * @param {number} y - Y coordinate for the top-left corner
 * @param {number} width - Width of the triangle's base
 * @param {number} height - Height of the triangle's height
 * @returns {Polygon} The created right-angle triangle object
 * @example
 * // Create a right-angle triangle
 * var rightTriangle = _rightAngleTriangle(100, 100, 50, 50);
 */
function _rightAngleTriangle(x, y, width, height) {
  var points = [
    [x, y],
    [x + width, y],
    [x + width, y + height],
  ];
  return _polygonCustom(points);
}

/**
 * Adds a guide to the document
 * @param {string} orientation - The orientation of the guide ("horizontal" or "vertical")
 * @param {number} position - The position of the guide (in points)
 * @returns {Guide} The created guide object
 * @example
 * // Add a horizontal guide at 100 points
 * var myGuide = _guide("horizontal", 100);
 */
function _guide(orientation, position) {
  try {
    var guideLayer = _layer("Guides");

    var guide = guideLayer.guides.add();
    guide.orientation =
      orientation.toLowerCase() === "horizontal"
        ? HorizontalOrVertical.HORIZONTAL
        : HorizontalOrVertical.VERTICAL;
    guide.location = position;

    return { guide: guide };
  } catch (e) {
    alert("Error adding guide: " + e.message);
    return null;
  }
}

/**
 * Sets the fill color of a shape
 * @param {PageItem} shape - The InDesign shape object
 * @param {Color|String} color - The fill color (InDesign color object or hex string)
 * @example
 * // Set red fill color
 * _setFillColor(myShape, "#FF0000");
 */
function _setFillColor(shape, color) {
  if (typeof color === "string") {
    color = _color(color);
  }
  shape.fillColor = color;
}

/**
 * Sets the stroke of a shape with specified weight and color.
 * @param {PageItem} shape - The InDesign shape object.
 * @param {number} weight - The stroke weight.
 * @param {Color|String} color - The stroke color (InDesign color object or hex string).
 * @example
 * // Set black 2pt stroke
 * setStroke(myShape, 2, "#000000");
 */
function _setStroke(shape, weight, color) {
  shape.strokeWeight = weight;
  shape.strokeColor = color;
}

/**
 * Removes the stroke from a shape.
 * @param {PageItem} shape - The InDesign shape object.
 * @example
 * // Remove stroke from the shape
 * noStroke(myShape);
 */
function _noStroke(shape) {
  shape.strokeWeight = 0;
}

/**
 * Converts a hex color string to RGB array
 * @param {string} hex - The hex color string (e.g., "#FF0000")
 * @returns {Array<number>} Array of [r,g,b] values (0-255)
 * @example
 * var rgbColor = _hexToRGB("#FF0000"); // Returns [255, 0, 0]
 */
function _hexToRGB(hex) {
  var r = parseInt(hex.substring(1, 3), 16);
  var g = parseInt(hex.substring(3, 5), 16);
  var b = parseInt(hex.substring(5, 7), 16);
  return [r, g, b];
}

/**
 * Creates or retrieves a color in the document
 * @param {string} clr - The hex color string (e.g., "#FF0000")
 * @returns {Color} The InDesign color object
 * @example
 * var redColor = _colors("#FF0000");
 */
function _color(clr) {
  var doc = _doc();
  var color;
  try {
    color = doc.colors.item(clr);
    color.properties = {
      model: ColorModel.PROCESS,
      space: ColorSpace.RGB,
      colorValue: _hexToRGB(clr),
    };
  } catch (e) {
    color = doc.colors.add({
      name: clr,
      model: ColorModel.PROCESS,
      space: ColorSpace.RGB,
      colorValue: _hexToRGB(clr),
    });
  }
  return color;
}

/**
 * Creates or retrieves an RGB color in the document
 * @param {number} r - Red value (0-255)
 * @param {number} g - Green value (0-255)
 * @param {number} b - Blue value (0-255)
 * @param {string} [name] - Optional name for the color. If not provided, generates one
 * @returns {Color} The InDesign color object
 * @example
 * var redColor = _colorRGB(255, 0, 0);
 * var namedBlue = _colorRGB(0, 0, 255, "MyBlue");
 */
function _colorRGB(r, g, b, name) {
  var doc = _doc();

  // Validate RGB values
  r = Math.min(255, Math.max(0, Math.round(r)));
  g = Math.min(255, Math.max(0, Math.round(g)));
  b = Math.min(255, Math.max(0, Math.round(b)));

  // Generate color name if not provided
  if (!name) {
    name = "RGB_" + r + "_" + g + "_" + b;
  }

  var color;
  try {
    // Try to get existing color
    color = doc.colors.item(name);
    color.properties = {
      model: ColorModel.PROCESS,
      space: ColorSpace.RGB,
      colorValue: [r, g, b],
    };
  } catch (e) {
    // Create new color if it doesn't exist
    color = doc.colors.add({
      name: name,
      model: ColorModel.PROCESS,
      space: ColorSpace.RGB,
      colorValue: [r, g, b],
    });
  }
  return color;
}

/**
 * Creates or retrieves a CMYK color in the document
 * @param {number} c - Cyan value (0-100)
 * @param {number} m - Magenta value (0-100)
 * @param {number} y - Yellow value (0-100)
 * @param {number} k - Black value (0-100)
 * @param {string} [name] - Optional name for the color. If not provided, generates one
 * @returns {Color} The InDesign color object
 * @example
 * var magentaColor = _colorCMYK(0, 100, 0, 0);
 * var namedGreen = _colorCMYK(100, 0, 100, 0, "MyGreen");
 */
function _colorCMYK(c, m, y, k, name) {
  var doc = _doc();

  // Validate CMYK values
  c = Math.min(100, Math.max(0, Math.round(c)));
  m = Math.min(100, Math.max(0, Math.round(m)));
  y = Math.min(100, Math.max(0, Math.round(y)));
  k = Math.min(100, Math.max(0, Math.round(k)));

  // Generate color name if not provided
  if (!name) {
    name = "CMYK_" + c + "_" + m + "_" + y + "_" + k;
  }

  var color;
  try {
    // Try to get existing color
    color = doc.colors.item(name);
    color.properties = {
      model: ColorModel.PROCESS,
      space: ColorSpace.CMYK,
      colorValue: [c, m, y, k],
    };
  } catch (e) {
    // Create new color if it doesn't exist
    color = doc.colors.add({
      name: name,
      model: ColorModel.PROCESS,
      space: ColorSpace.CMYK,
      colorValue: [c, m, y, k],
    });
  }
  return color;
}

/**
 * Returns all colors in the swatch
 * @returns {Array<Color>} An array of all color objects in the swatch
 * @example
 * var allColors = _colors();
 * allColors.forEach(function(color) {
 *   alert(color.name);
 * });
 */
function _colors() {
  var doc = _doc();
  return doc.colors.everyItem().getElements();
}

/**
 * Linearly interpolates between two colors
 * @param {Array<number>} c1 - First color as RGB array [r,g,b]
 * @param {Array<number>} c2 - Second color as RGB array [r,g,b]
 * @param {number} t - Interpolation value (0 to 1)
 * @returns {Array<number>} The interpolated RGB color
 * @example
 * // Get color halfway between red and blue
 * var purple = _lerpColor([255,0,0], [0,0,255], 0.5);
 */
function _lerpColor(c1, c2, t) {
  var r = Math.round(lerp(c1[0], c2[0], t));
  var g = Math.round(lerp(c1[1], c2[1], t));
  var b = Math.round(lerp(c1[2], c2[2], t));
  return [r, g, b];
}

/**
 * Linearly interpolates between two values
 * @param {number} val1 - First value
 * @param {number} val2 - Second value
 * @param {number} t - Interpolation value (0 to 1)
 * @returns {number} The interpolated value
 * @example
 * var half = _lerp(0, 100, 0.5); // Returns 50
 */
function _lerp(val1, val2, t) {
  return val1 * (1 - t) + val2 * t;
}

/**
 * Maps a value from one range to another
 * @param {number} val - The value to map
 * @param {number} min - The lower bound of the current range
 * @param {number} max - The upper bound of the current range
 * @param {number} newMin - The lower bound of the target range
 * @param {number} newMax - The upper bound of the target range
 * @returns {number} The mapped value in the new range
 * @example
 * // Map 50 from range 0-100 to range 0-1
 * var mapped = _map(50, 0, 100, 0, 1); // Returns 0.5
 *
 * // Map 0.5 from range 0-1 to range 0-100
 * var percent = _map(0.5, 0, 1, 0, 100); // Returns 50
 */
function _map(val, min, max, newMin, newMax) {
  return ((val - min) * (newMax - newMin)) / (max - min) + newMin;
}

// TEXT FUNCTIONS

/**
 * Creates a text frame on the page
 * @param {number} x - The x-coordinate of the text frame's left edge
 * @param {number} y - The y-coordinate of the text frame's top edge
 * @param {number} width - The width of the text frame
 * @param {number} height - The height of the text frame
 * @param {string} text - The text content to place in the frame
 * @returns {object} The created text frame object with methods for styling
 * @example
 * // Create a text frame at (10,10) size 200x50 with "Hello World"
 * var myText = _textFrame(10, 10, 200, 50, "Hello World");
 */
function _textFrame(x, y, width, height, text) {
  var doc = _doc();
  var page = doc.pages.item(0);
  var textFrame = page.textFrames.add();
  textFrame.geometricBounds = [y, x, height, width];
  textFrame.contents = text;
  return {
    frame: textFrame,
    fontSize: function (size) {
      _setFontSize(textFrame, size);
    },
    fontColor: function (color) {
      _setFontColor(textFrame, color);
    },
    fontName: function (name) {
      _setFontName(textFrame, name);
    },
    centerTo: function (keyObject) {
      _centerTo(textFrame, keyObject);
    },
    textJustification: function (justification) {
      _setTextJustification(textFrame, justification);
    },
    verticalJustification: function (justification) {
      _setVerticalJustification(textFrame, justification);
    },
    noHyphenation: function () {
      _noHyphenation(textFrame);
    },
    findReplace: function (grep, replaceWith) {
      _grepFindAndReplace(textFrame, grep, replaceWith);
    },
    characters: function () {
      return _characters(textFrame);
    },
    words: function () {
      return _words(textFrame);
    },
    flip: function (direction) {
      _flip(textFrame, direction);
    },
    rotate: function (angle) {
      _rotate(textFrame, angle);
    },
  };
}

/**
 * Sets the font size for a text frame
 * @param {TextFrame} textFrame - The InDesign text frame object
 * @param {number} fontSize - The desired font size
 * @example
 * // Set the font size to 12 points for a text frame
 * _setFontSize(myTextFrame, 12);
 */
function _setFontSize(textFrame, fontSize) {
  try {
    textFrame.texts.item(0).pointSize = fontSize;
    return true;
  } catch (e) {
    alert("Error setting font size: " + e.message);
    return false;
  }
}

/**
 * Sets the font color for a text frame
 * @param {TextFrame} textFrame - The InDesign text frame object
 * @param {Color|String} color - The font color (InDesign color object or hex string)
 * @example
 * // Set the font color to red
 * _setFontColor(myTextFrame, "#FF0000");
 */
function _setFontColor(textFrame, color) {
  try {
    if (typeof color === "string") {
      color = _color(color);
    }
    textFrame.texts.item(0).fillColor = color;
    return true;
  } catch (e) {
    alert("Error setting font color: " + e.message);
    return false;
  }
}

/**
 * Sets the font name for a text frame
 * @param {TextFrame} textFrame - The InDesign text frame object
 * @param {string} fontName - The desired font name
 * @example
 * // Set the font name to "Arial"
 * _setFontName(myTextFrame, "Arial");
 */
function _setFontName(textFrame, fontName) {
  try {
    textFrame.texts.item(0).appliedFont = fontName;
    return true;
  } catch (e) {
    alert("Error setting font name: " + e.message);
    return false;
  }
}

/**
 * Sets the vertical justification for a text frame
 * @param {TextFrame} textFrame - The InDesign text frame object
 * @param {VerticalJustification|string} justification - The desired vertical justification (e.g., VerticalJustification.CENTER_ALIGN or "center")
 * @returns {boolean} - Success status
 * @example
 * // Set the vertical justification to center
 * _setVerticalJustification(myTextFrame, "center");
 */
function _setVerticalJustification(textFrame, justification) {
  const justificationMap = {
    top: VerticalJustification.TOP_ALIGN,
    center: VerticalJustification.CENTER_ALIGN,
    bottom: VerticalJustification.BOTTOM_ALIGN,
    justify: VerticalJustification.JUSTIFY_ALIGN,
  };

  try {
    if (typeof justification === "string") {
      if (!justificationMap[justification.toLowerCase()]) {
        throw new Error(
          "Invalid vertical justification value: " + justification
        );
      }
      justification = justificationMap[justification.toLowerCase()];
    }
    textFrame.textFramePreferences.verticalJustification = justification;
    return true;
  } catch (e) {
    alert("Error setting vertical justification: " + e.message);
    return false;
  }
}

/**
 * Sets the text justification for a text frame
 * @param {TextFrame} textFrame - The InDesign text frame object
 * @param {Justification|string} justification - The desired text justification (e.g., Justification.CENTER_ALIGN or "center")
 * @returns {boolean} - Success status
 * @example
 * // Set the text justification to center
 * _setTextJustification(myTextFrame, "center");
 */
function _setTextJustification(textFrame, justification) {
  const justificationMap = {
    left: Justification.LEFT_ALIGN,
    center: Justification.CENTER_ALIGN,
    right: Justification.RIGHT_ALIGN,
    justify: Justification.FULLY_JUSTIFIED,
  };

  try {
    if (typeof justification === "string") {
      if (!justificationMap[justification.toLowerCase()]) {
        throw new Error("Invalid justification value: " + justification);
      }
      justification = justificationMap[justification.toLowerCase()];
    }
    textFrame.texts.item(0).justification = justification;
    return true;
  } catch (e) {
    alert("Error setting text justification: " + e.message);
    return false;
  }
}

/**
 * Returns all characters in a text frame
 * @param {TextFrame} textFrame - The InDesign text frame object
 * @returns {Array<Character>} An array of all character objects in the text frame
 * @example
 * // Get all characters in a text frame
 * var characters = _characters(myTextFrame);
 */
function _characters(textFrame) {
  try {
    return textFrame.characters.everyItem().getElements();
  } catch (e) {
    alert("Error getting characters: " + e.message);
    return [];
  }
}

/**
 * Returns all words in a text frame
 * @param {TextFrame} textFrame - The InDesign text frame object
 * @returns {Array<Word>} An array of all word objects in the text frame
 * @example
 * // Get all words in a text frame
 * var words = _words(myTextFrame);
 */
function _words(textFrame) {
  try {
    return textFrame.words.everyItem().getElements();
  } catch (e) {
    alert("Error getting words: " + e.message);
    return [];
  }
}

/**
 * Creates a character style in the document, removing any existing style with the same name
 * @param {string} styleName - The name of the character style
 * @param {Object} properties - The properties to apply to the character style
 * @returns {CharacterStyle} The created character style object
 * @example
 * // Create a character style named "BoldRed" with bold font and red color
 * var boldRedStyle = _createCharacterStyle("BoldRed", { fontStyle: "Bold", fillColor: "Red" });
 */
function _createCharacterStyle(styleName, properties) {
  try {
    var doc = _doc();
    var charStyle;

    // Check if the character style already exists and remove it
    if (doc.characterStyles.itemByName(styleName).isValid) {
      doc.characterStyles.itemByName(styleName).remove();
    }

    // Create a new character style
    charStyle = doc.characterStyles.add({ name: styleName });
    for (var prop in properties) {
      if (properties.hasOwnProperty(prop)) {
        charStyle[prop] = properties[prop];
      }
    }
    return charStyle;
  } catch (e) {
    alert("Error creating character style: " + e.message);
    return null;
  }
}

/**
 * Finds words in a text frame using GREP (regular expressions)
 * @param {TextFrame} textFrame - The InDesign text frame object
 * @param {string} grep - The GREP pattern to search for
 * @returns {Array<Text>} An array of found text objects
 * @example
 * // Find all occurrences of the word "Omelas" in the text frame
 * var foundWords = _grepFindText(myTextFrame, "Omelas");
 */
function _grepFindText(grep) {
  try {
    app.findGrepPreferences = app.changeGrepPreferences = null;
    app.findGrepPreferences.findWhat = grep;
    var foundItems = app.findGrepPreferences.getElements();
    return foundItems;
  } catch (e) {
    alert("Error finding words: " + e.message);
    return [];
  }
}

/**
 * Finds and replaces text in a text frame using GREP (regular expressions)
 * @param {TextFrame} textFrame - The InDesign text frame object
 * @param {string} grep - The GREP pattern to search for
 * @param {string} replaceWith - The text to replace the found text with
 * @returns {boolean} - Success status
 * @example
 * // Replace all occurrences of "Omelas" with "City" in the text frame
 * _grepFindAndReplace(myTextFrame, "Omelas", "City");
 */
function _grepFindAndReplace(textFrame, grep, replaceWith) {
  try {
    app.findGrepPreferences = app.changeGrepPreferences = null;
    app.findGrepPreferences.findWhat = grep;
    app.changeGrepPreferences.changeTo = replaceWith;
    textFrame.changeGrep();
    return true;
  } catch (e) {
    alert("Error finding and replacing text: " + e.message);
    return false;
  }
}

/**
 * Applies a character style to text found using a GREP pattern
 * @param {TextFrame} textFrame - The InDesign text frame object
 * @param {string} grep - The GREP pattern to search for
 * @param {CharacterStyle} charStyle - The character style to apply
 * @returns {boolean} - Success status
 * @example
 * // Apply the "BoldRed" character style to all occurrences of "Omelas" in the text frame
 * _applyGrepCharacterStyle(myTextFrame, "Omelas", boldRedStyle);
 */
function _applyGrepCharacterStyle(textFrame, grep, charStyle) {
  try {
    app.findGrepPreferences = app.changeGrepPreferences = null;
    app.findGrepPreferences.findWhat = grep;
    app.changeGrepPreferences.appliedCharacterStyle = charStyle;
    textFrame.changeGrep();
    return true;
  } catch (e) {
    alert("Error applying GREP character style: " + e.message);
    return false;
  }
}

/**
 * Disables hyphenation for a text frame
 * @param {TextFrame} textFrame - The InDesign text frame object
 * @returns {boolean} - Success status
 * @example
 * // Disable hyphenation for a text frame
 * _noHyphenation(myTextFrame);
 */
function _noHyphenation(textFrame) {
  try {
    textFrame.texts.item(0).hyphenation = false;
    return true;
  } catch (e) {
    alert("Error disabling hyphenation: " + e.message);
    return false;
  }
}

// IMAGE FUNCTIONS

/**
 * Places an image file in a rectangle on the page
 * @param {string} src - File path to the image
 * @param {number} x - The x-coordinate of the image container's left edge
 * @param {number} y - The y-coordinate of the image container's top edge
 * @param {number} width - The width of the image container
 * @param {number} height - The height of the image container
 * @returns {object} The created image object with methods for styling
 * @example
 * // Place an image at (10, 10) with width 200 and height 150
 * var myImage = _image("/path/to/image.jpg", 10, 10, 200, 150);
 */
function _image(src, x, y, width, height) {
  try {
    var doc = _doc();
    var page = doc.pages.item(0);
    var imageFrame = page.rectangles.add();
    imageFrame.geometricBounds = [y, x, y + height, x + width];
    imageFrame.place(File(src));
    imageFrame.strokeWeight = 0;
    imageFrame.fit(FitOptions.FILL_PROPORTIONALLY);
    return {
      frame: imageFrame,
      fit: function () {
        imageFrame.fit(FitOptions.FILL_PROPORTIONALLY);
      },
      blendMode: function (mode) {
        _setImageBlendMode(imageFrame, mode);
      },
      textWrap: function (offset) {
        _setTextWrap(imageFrame, offset);
      },
      centerTo: function (keyObject) {
        _centerTo(imageFrame, keyObject);
      },
      flip: function (direction) {
        _flip(imageFrame, direction);
      },
      rotate: function (angle) {
        _rotate(imageFrame, angle);
      },
      cornerRadius: function (radius) {
        _cornerOption(imageFrame, CornerOptions.ROUNDED_CORNER, radius);
      },
    };
  } catch (e) {
    alert("Error placing image: " + e.message);
    return null;
  }
}

/**
 * Sets the blend mode for an image
 * @param {Rectangle} imageFrame - The rectangle object containing the placed image
 * @param {BlendMode|string} blendMode - The desired blend mode (e.g., BlendMode.MULTIPLY or "multiply")
 * @returns {boolean} - Success status
 * @example
 * // Set the blend mode to multiply for an image
 * _setImageBlendMode(myImageFrame, "multiply");
 */
function _setImageBlendMode(imageFrame, blendMode) {
  const blendModeMap = {
    normal: BlendMode.NORMAL,
    multiply: BlendMode.MULTIPLY,
    screen: BlendMode.SCREEN,
    overlay: BlendMode.OVERLAY,
    softLight: BlendMode.SOFT_LIGHT,
    hardLight: BlendMode.HARD_LIGHT,
    colorDodge: BlendMode.COLOR_DODGE,
    colorBurn: BlendMode.COLOR_BURN,
    darken: BlendMode.DARKEN,
    lighten: BlendMode.LIGHTEN,
    difference: BlendMode.DIFFERENCE,
    exclusion: BlendMode.EXCLUSION,
    hue: BlendMode.HUE,
    saturation: BlendMode.SATURATION,
    color: BlendMode.COLOR,
    luminosity: BlendMode.LUMINOSITY,
  };

  try {
    if (typeof blendMode === "string") {
      if (!blendModeMap[blendMode.toLowerCase()]) {
        throw new Error("Invalid blend mode value: " + blendMode);
      }
      blendMode = blendModeMap[blendMode.toLowerCase()];
    }
    imageFrame.transparencySettings.blendingSettings.blendMode = blendMode;
    return true;
  } catch (e) {
    alert("Error setting blend mode: " + e.message);
    return false;
  }
}

/**
 * Sets text wrap options for an image
 * @param {Rectangle} imageFrame - The rectangle object containing the placed image
 * @param {number} offset - The offset distance for the text wrap (in points)
 * @returns {boolean} - Success status
 * @example
 * // Set text wrap with a 10-point offset for an image
 * _setTextWrap(myImageFrame, 10);
 */
function _setTextWrap(imageFrame, offset) {
  try {
    imageFrame.textWrapPreferences.textWrapMode =
      TextWrapModes.BOUNDING_BOX_TEXT_WRAP;
    imageFrame.textWrapPreferences.textWrapOffset = [
      offset,
      offset,
      offset,
      offset,
    ];
    return true;
  } catch (e) {
    alert("Error setting text wrap: " + e.message);
    return false;
  }
}

// UTILITY FUNCTIONS

/**
 * Aligns a frame to the specified alignment target
 * @param {PageItem} frame - The InDesign frame object (e.g., Rectangle, TextFrame, etc.)
 * @param {string} target - The alignment target ("page", "margin", "spread", "selection")
 * @param {AlignOptions|string} alignment - The desired alignment (e.g., AlignOptions.HORIZONTAL_CENTERS or "center")
 * @returns {boolean} - Success status
 * @example
 * // Align a rectangle to the horizontal center of the page
 * _alignFrame(myRect, "page", AlignOptions.HORIZONTAL_CENTERS);
 * // Align a rectangle to the center of the page using a string
 * _alignFrame(myRect, "page", "center");
 */
function _alignFrame(frame, target, alignment) {
  const alignmentMap = {
    top: AlignOptions.TOP_EDGES,
    bottom: AlignOptions.BOTTOM_EDGES,
    left: AlignOptions.LEFT_EDGES,
    right: AlignOptions.RIGHT_EDGES,
    center: AlignOptions.HORIZONTAL_CENTERS,
    "vertical center": AlignOptions.VERTICAL_CENTERS,
  };

  try {
    var doc = _doc();
    var alignTo;

    switch (target.toLowerCase()) {
      case "page":
        alignTo = AlignDistributeBounds.PAGE_BOUNDS;
        break;
      case "margin":
        alignTo = AlignDistributeBounds.MARGIN_BOUNDS;
        break;
      case "spread":
        alignTo = AlignDistributeBounds.SPREAD_BOUNDS;
        break;
      case "selection":
        alignTo = AlignDistributeBounds.SELECTION_BOUNDS;
        break;
      default:
        throw new Error("Invalid alignment target: " + target);
    }

    if (typeof alignment === "string") {
      if (!alignmentMap[alignment.toLowerCase()]) {
        throw new Error("Invalid alignment value: " + alignment);
      }
      alignment = alignmentMap[alignment.toLowerCase()];
    }

    doc.align(frame, alignment, alignTo);
    return true;
  } catch (e) {
    alert("Error aligning frame: " + e.message);
    return false;
  }
}

/**
 * Centers a centeringObject to a keyObject both vertically and horizontally
 * @param {PageItem} centeringObject - The InDesign frame object to be centered (e.g., Rectangle, TextFrame, etc.)
 * @param {PageItem} [keyObject=null] - The InDesign frame object to center to (e.g., Rectangle, TextFrame, etc.). If null, centers to the page.
 * @returns {boolean} - Success status
 * @example
 * // Center a rectangle to another rectangle
 * _centerTo(myRect, keyRect);
 * // Center a rectangle to the page
 * _centerTo(myRect);
 */
function _centerTo(centeringObject, keyObject) {
  try {
    var doc = _doc();
    var alignTo = keyObject
      ? AlignDistributeBounds.KEY_OBJECT
      : AlignDistributeBounds.PAGE_BOUNDS;

    // Align horizontally
    doc.align(
      centeringObject,
      AlignOptions.HORIZONTAL_CENTERS,
      alignTo,
      keyObject
    );
    // Align vertically
    doc.align(
      centeringObject,
      AlignOptions.VERTICAL_CENTERS,
      alignTo,
      keyObject
    );

    return true;
  } catch (e) {
    alert("Error centering object: " + e.message);
    return false;
  }
}

/**
 * Centers a frame on the page
 * @param {PageItem} frame - The InDesign frame object (e.g., Rectangle, TextFrame, etc.)
 * @returns {boolean} - Success status
 * @example
 * // Center a rectangle on the page
 * _centerFrame(myRect);
 */
function _centerFrame(frame) {
  try {
    var doc = _doc();
    var page = doc.pages.item(0);
    var pageWidth = page.bounds[3] - page.bounds[1];
    var pageHeight = page.bounds[2] - page.bounds[0];
    var frameWidth = frame.geometricBounds[3] - frame.geometricBounds[1];
    var frameHeight = frame.geometricBounds[2] - frame.geometricBounds[0];

    var newX = (pageWidth - frameWidth) / 2;
    var newY = (pageHeight - frameHeight) / 2;

    frame.geometricBounds = [newY, newX, newY + frameHeight, newX + frameWidth];
    return true;
  } catch (e) {
    alert("Error centering frame: " + e.message);
    return false;
  }
}

/**
 * Flips an object either horizontally or vertically
 * @param {PageItem} object - The InDesign frame object to be flipped (e.g., Rectangle, TextFrame, etc.)
 * @param {string} direction - The direction to flip ("horizontal" or "vertical")
 * @param {AnchorPoint} [referencePoint=AnchorPoint.CENTER_ANCHOR] - The reference point for the flip (e.g., AnchorPoint.CENTER_ANCHOR)
 * @returns {boolean} - Success status
 * @example
 * // Flip a rectangle horizontally with center anchor point
 * _flip(myRect, "horizontal", AnchorPoint.CENTER_ANCHOR);
 * // Flip a rectangle vertically with default center anchor point
 * _flip(myRect, "vertical");
 */
function _flip(object, direction, referencePoint) {
  try {
    var doc = _doc();
    referencePoint = referencePoint || AnchorPoint.CENTER_ANCHOR;
    doc.layoutWindows[0].transformReferencePoint = referencePoint;

    switch (direction.toLowerCase()) {
      case "horizontal":
        object.flipItem(Flip.HORIZONTAL, referencePoint);
        break;
      case "vertical":
        object.flipItem(Flip.VERTICAL, referencePoint);
        break;
      default:
        throw new Error("Invalid flip direction: " + direction);
    }

    return true;
  } catch (e) {
    alert("Error flipping object: " + e.message);
    return false;
  }
}

/**
 * Rotates an object by a specified angle
 * @param {PageItem} object - The InDesign frame object to be rotated (e.g., Rectangle, TextFrame, etc.)
 * @param {number} angle - The angle to rotate the object (in degrees)
 * @param {AnchorPoint} [referencePoint=AnchorPoint.CENTER_ANCHOR] - The reference point for the rotation (e.g., AnchorPoint.CENTER_ANCHOR)
 * @returns {boolean} - Success status
 * @example
 * // Rotate a rectangle by 45 degrees with center anchor point
 * _rotate(myRect, 45, AnchorPoint.CENTER_ANCHOR);
 * // Rotate a rectangle by 90 degrees with default center anchor point
 * _rotate(myRect, 90);
 */
function _rotate(object, angle, referencePoint) {
  try {
    var doc = _doc();
    referencePoint = referencePoint || AnchorPoint.CENTER_ANCHOR;
    doc.layoutWindows[0].transformReferencePoint = referencePoint;

    object.rotationAngle = angle;
    return true;
  } catch (e) {
    alert("Error rotating object: " + e.message);
    return false;
  }
}

/**
 * Positions a frame at a random position within given width and height bounds
 * @param {PageItem} frame - The InDesign frame object (Rectangle, TextFrame, Oval, etc.)
 * @param {number} maxWidth - The maximum width boundary in current document units
 * @param {number} maxHeight - The maximum height boundary in current document units
 * @param {number} [padding=0] - Optional padding from boundaries in current document units
 * @returns {boolean} - Success status
 * @example
 * // Position a rectangle randomly within 500x700 bounds with 10pt padding
 * var rect = _rect(0, 0, 100, 50);  // Create a 100x50 rectangle
 * _randomPosFrame(rect, 500, 700, 10);
 */
function _randomPosFrame(frame, maxWidth, maxHeight, padding) {
  if (padding == undefined) padding = 0;

  try {
    // Get frame dimensions from geometric bounds [y1, x1, y2, x2]
    var bounds = frame.geometricBounds;
    var frameWidth = bounds[3] - bounds[1];
    var frameHeight = bounds[2] - bounds[0];

    // Make sure frame fits within bounds
    if (
      frameWidth + padding * 2 > maxWidth ||
      frameHeight + padding * 2 > maxHeight
    ) {
      alert("Frame is too large for the specified bounds");
      return false;
    }

    // Calculate random position within bounds, accounting for padding
    var randomX =
      padding + Math.random() * (maxWidth - frameWidth - padding * 2);
    var randomY =
      padding + Math.random() * (maxHeight - frameHeight - padding * 2);

    // Set new position while maintaining frame dimensions
    frame.geometricBounds = [
      randomY, // y1 (top)
      randomX, // x1 (left)
      randomY + frameHeight, // y2 (bottom)
      randomX + frameWidth, // x2 (right)
    ];

    return true;
  } catch (e) {
    alert("Error positioning frame: " + e.message);
    return false;
  }
}

/**
 * Generates a random number between two values
 * @param {number} min - The minimum value (inclusive)
 * @param {number} max - The maximum value (exclusive)
 * @returns {number} A random number between min and max
 * @example
 * // Get random number between 0 and 10
 * var rand = _random(0, 10);
 *
 * // Get random integer between 1 and 6 (dice roll)
 * var dice = Math.floor(_random(1, 7));
 */
function _random(min, max) {
  return Math.floor(Math.random() * (max - min) + min);
}

/**
 * Sets the measurement units for a document
 * @param {Document} doc - The InDesign document
 * @param {string} unit - The desired unit ('pt', 'mm', 'inch', 'px', 'cm', 'p', 'ag', 'c')
 * @returns {boolean} - Success status
 */
function _setMeasurementUnit(doc, unit) {
  // Map of unit shortcuts to MeasurementUnits enum
  const unitMap = {
    pt: MeasurementUnits.POINTS,
    mm: MeasurementUnits.MILLIMETERS,
    inch: MeasurementUnits.INCHES, // Changed from 'in' to 'inch'
    px: MeasurementUnits.PIXELS,
    cm: MeasurementUnits.CENTIMETERS,
    p: MeasurementUnits.PICAS,
    ag: MeasurementUnits.AGATES,
    c: MeasurementUnits.CICEROS,
  };

  try {
    // Check if unit is valid
    if (!unitMap[unit.toLowerCase()]) {
      throw new Error(
        "Invalid measurement unit. Valid units are: " +
          Object.keys(unitMap).join(", ")
      );
    }

    // Set both horizontal and vertical units
    doc.viewPreferences.horizontalMeasurementUnits =
      unitMap[unit.toLowerCase()];
    doc.viewPreferences.verticalMeasurementUnits = unitMap[unit.toLowerCase()];

    // Set ruler units
    doc.viewPreferences.rulerOrigin = RulerOrigin.PAGE_ORIGIN;

    // Set default units for new objects
    app.scriptPreferences.measurementUnit = unitMap[unit.toLowerCase()];

    return true;
  } catch (e) {
    alert("Error setting measurement unit: " + e.message);
    return false;
  }
}

/**
 * Converts a value from one unit to another
 * @param {number} value - The value to convert
 * @param {string} fromUnit - The current unit
 * @param {string} toUnit - The target unit
 * @returns {number} - The converted value
 */
function _convertUnits(value, fromUnit, toUnit) {
  // Conversion factors to points
  const toPoints = {
    pt: 1,
    mm: 2.834645669,
    inch: 72, // Changed from 'in' to 'inch'
    px: 1,
    cm: 28.34645669,
    p: 12, // picas
    ag: 14.4, // agates
    c: 12.7878, // ciceros
  };

  try {
    // Convert to points first
    const points = value * toPoints[fromUnit.toLowerCase()];
    // Then convert to target unit
    return points / toPoints[toUnit.toLowerCase()];
  } catch (e) {
    alert("Error converting units: " + e.message);
    return value;
  }
}
