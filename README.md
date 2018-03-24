# pptx-generator

#### pptx-generator npm module used to generate pptx along with the features of text add, image add, read existing template.


#### Using open source pptxgenjs plugin with some customization

## Installation

```bash
npm i pptx-generator --save
```

## TEXT OPTIONS

```js
 x ( x position)
 y ( y position)
 align
 font_size
 color
 w (width)
 h (height)
 autoFit
```

## IMAGE OPTIONS

```js
 path : Image Path
 x (x position)
 y (y position)
 w (width)
 h (height)
```


## Sample

```js

var PPTX_GENERATOR = require('pptx-generator');

PPTX_GENERATOR.createPresentation("/tmp/test-template.pptx").then(function(presentation){
    var slide = PPTX_GENERATOR.addNewSlide();
    PPTX_GENERATOR.addText(slide, "Hello World !!", { x:1, y:3, align:'c', font_size:40, color:'ffffff',w: 9.0,h:0.5,autoFit:true});

    var slide1 = PPTX_GENERATOR.addNewSlide();
    PPTX_GENERATOR.addText(slide1, "Hello World1 !!", { x:0.5, y:0.5, align:'c', font_size:20, color:'ffffff',w: 9.0,h:0.5,autoFit:true});
    PPTX_GENERATOR.addImage(slide1, { path: './media/image4.jpeg', x:3.0, y:1.5, w:3, h:3});

    var slide2 = PPTX_GENERATOR.addNewSlide();
    PPTX_GENERATOR.addText(slide2, "Hello World3 !!", { x:1, y:3, align:'c', font_size:40, color:'ffffff',w: 9.0,h:0.5,autoFit:true});

    PPTX_GENERATOR.generate("/tmp/test.pptx");
    res.send('Done !!')
});

```