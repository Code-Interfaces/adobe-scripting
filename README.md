# Adobe Scripting

This is a brief introduction to scripting in Adobe applications. We will be using ExtendScript, which is a JavaScript-based scripting language. ExtendScript is used to automate tasks in Adobe applications. It is a dialect of JavaScript, so if you are familiar with JavaScript, you will be able to quickly pick up ExtendScript. If you are not familiar with JavaScript, don't worry, you will still be able to follow along.

## Running Scripts

### InDesign

To run a script in InDesign, open the Scripts panel by going to `Window > Utilities > Scripts`. Then, navigate to the folder that contains the script that you want to run. Double click on the script to run it.

### Illustrator 2024

To run a script in Illustrator, open the Scripts panel by going to `File > Scripts > Other Script...`. Then, navigate to the folder that contains the script that you want to run. Double click on the script to run it.

## A brief intro to ExtendScript (JavaScript)

### Variables

Variables are declared using the `var` keyword. Variables can be assigned a value using the `=` operator. You can reassign a variable to a new value by using the `=` operator again.

```javascript
var myVariable = 1;
myVariable = 2;
```

### Data Types

There are several different data types in JavaScript. The most common ones you would come across are:

- `Number` - A number, such as `1` or `3.14`
- `String` - A string of characters, such as `"Hello World"`
- `Boolean` - A boolean value, either `true` or `false`
- `Object` - An object, such as `{}`
- `Array` - An array, such as `[]`
- `Function` - A function, such as `function() {}`

### Conditionals

Conditionals are used to execute code based on a condition. The most common conditional is the `if` statement. The `if` statement takes a condition, and if the condition is true, it executes the code inside the curly braces.

```javascript
if (true) {
  // This code will run
}

if (false) {
  // This code will not run
}
```

You can also use the `else` keyword to execute code if the condition is false.

```javascript
if (true) {
  // This code will run
} else {
  // This code will not run
}

if (false) {
  // This code will not run
} else {
  // This code will run
}
```

You can also use the `else if` keyword to execute code if the condition is false, but another condition is true.

```javascript
if (true) {
  // This code will run
} else if (true) {
  // This code will not run
} else {
  // This code will not run
}

if (false) {
  // This code will not run
} else if (true) {
  // This code will run
} else {
  // This code will not run
}

if (false) {
  // This code will not run
} else if (false) {
  // This code will not run
} else {
  // This code will run
}
```

### Loops

Loops are used to execute code multiple times. The most common loop is the `for` loop. The `for` loop takes a variable, a condition, and an incrementer. The variable is initialized to a value, and the condition is checked. If the condition is true, the code inside the curly braces is executed. Then, the incrementer is executed, and the condition is checked again. This process repeats until the condition is false.

```javascript
for (var i = 0; i < 10; i++) {
  // This code will run 10 times
  // i will be 0, 1, 2, 3, 4, 5, 6, 7, 8, 9
}
```

### Functions

Functions are used to execute code. Functions can take parameters, which are variables that are passed into the function. Functions can also return values, which are variables that are returned from the function.
One thing to note is extendScript is not fully compatible with ES6, so arrow functions are not supported and also you can not have default parameters.

```javascript
function myFunction() {
  // This code will run
}

myFunction(); // This will call the function

function myFunctionWithParameters(parameter1, parameter2) {
  // This code will run
  // parameter1 and parameter2 are variables that are passed into the function
}

myFunctionWithParameters(1, 2); // This will call the function and pass in 1 and 2 as the parameters

function myFunctionWithReturnValue() {
  // This code will run
  return 1;
}

var returnValue = myFunctionWithReturnValue(); // This will call the function and assign the return value to the variable returnValue
```



---
### References and Resources

- [Adobe InDesign CS6 Scripting Guide: Javascript](https://usermanual.wiki/adobe/InDesigncs6ScriptingJSEN.3768967468/view)
- [InDesign Automatisieren 2. Auflage 2015](https://github.com/grefel/indesignjs)
- [Adobe Extensibility API Docs](https://docsforadobe.dev/)
- [InDesign ExtendScript API](https://www.indesignjs.de/extendscriptAPI/indesign-latest/#about.html)
- [JavaScript for InDesign, by Peter Kahrel](https://creativepro.com/now-available-javascript-for-indesign-2nd-edition/)
- [inDesignjs-resources](https://grefel.github.io/indesignjs-resources/)
- [Adobe Illustrator Scripting Guide](https://ai-scripting.docsforadobe.dev/)
