var PPT_Template 	= require('./ppt-index'),
    DEFERRED        = require('deferred'),
    FS 				= require('fs'),
    XML2JSON 		= require('xml2json');

var PPTXGEN 		= require(__dirname + '/pptxgenjs/dist/pptxgen.js');

var Presentation    = PPT_Template.Presentation;

var currentPresentation, PPTX, newSlides;

var PPTX_GENERATOR =  {

    createPresentation : function(template) {
        var deferred = new DEFERRED, _self = this;
        currentPresentation = new Presentation();

        currentPresentation.loadFile(template).then(function(){

            PPTX = new PPTXGEN();
            PPTX.setLayout('LAYOUT_4x3');
            PPTX.getPPTXT().imageCounter = _self.getImageCount();
            newSlides = [];
            deferred.resolve(currentPresentation);
        }).catch(function(err){
            console.error("error in reading file from : " + template);
            deferred.reject(err);
        });
        return deferred.promise;
    },

    getPresentation : function(){
        return currentPresentation;
    },


    addNewSlide : function(){
        return PPTX.addNewSlide();
    },

    addText : function(slide, text, options){
        slide.addText(text, options);
        return slide;
    },

    addImage : function(slide, options){
        slide.addImage(options);
        this.addMedia(options.path);
        return slide;
    },

    updateMedia: function(slide,placeHolder, path){
        var id = this.getImageId(slide.content,placeHolder);
        var bitmap  = Array.prototype.slice.call(FS.readFileSync(path), 0);
        if(currentPresentation.contents["ppt/media/image" + (id)+'.jpeg']){
            currentPresentation.contents["ppt/media/image" + (id)+'.jpeg'] = bitmap;
        }else if(currentPresentation.contents["ppt/media/image" + (id)+'.png']){
            currentPresentation.contents["ppt/media/image" + (id)+'.png'] = bitmap;
        }
    },

    getImageId : function(str, pat){
        var headerIndex = str.indexOf(pat);
        var subStr = str.substr(headerIndex-50, headerIndex);
        var idIndex  = subStr.search('id');
        var idSubStr = subStr.substr(idIndex,10);
        return idSubStr.split(`"`)[1];
    },

    replaceText : function(slide, placeHolder, value){
        slide.content = slide.content.replace(placeHolder, value);
        newSlides.push(slide);
        return slide;
    },

    /**
     * this function is for get 
     * the index of string to remove image
     */
    nthIndex : function(str, pat){
        var indexOfWord = str.indexOf(pat);
        var subStr      = str.substr(indexOfWord -100, indexOfWord + 900);
        var firstPicIndex = subStr.indexOf('<p:pic>');
        var lastPicindex  = subStr.indexOf('</p:pic>');
        var indexOfHeader = subStr.indexOf(pat);
        lastPicindex  =  (indexOfWord + lastPicindex) - ( indexOfHeader - 8) ;
        firstPicIndex = (indexOfWord - (firstPicIndex + 10 ));
        return  {firstPicIndex:firstPicIndex, lastPicindex : lastPicindex};
    },

    removeImage : function(slide,placeHolder,path){
        var content     = (JSON.parse(XML2JSON.toJson(slide.content)));
        var indexObj    = this.nthIndex(slide.content,placeHolder);
        var xml = (slide.content).substr(0, indexObj.firstPicIndex) + (slide.content).substr(indexObj.lastPicindex);
        slide.content = xml;
        return slide;
    },

    addMedia: function(path){
        try{
            path = this.validateMediaExtension(path);
            var extension      = 'png';
            if (path.indexOf('.') > -1 ) extension = path.split('.').pop();
            var bitmap = Array.prototype.slice.call(FS.readFileSync(path), 0);
            currentPresentation.contents["ppt/media/image" + PPTX.getPPTXT().imageCounter  + "."  + extension] = bitmap;
        }catch(err){
            console.error("error in adding media : " + path);
        }
    },

    validateMediaExtension: function(path){
        if(path.indexOf('.png') == -1 && path.indexOf('.jpeg') == -1){
            path = path.replace(path.split('.').pop(), "jpeg");
        }
        return path;
    },

    getImageCount: function(){
        return Object.keys(currentPresentation.contents).filter(function(key){
            return key.indexOf("ppt/media/image") > -1;
        }).length;
    },

    getSlide : function(index){
        return currentPresentation.getSlide(index);
    },

    getSlideCount : function(){
        return currentPresentation.getSlideCount();
    },

    getSlideClone : function(index){
        return currentPresentation.getSlide(index).clone();
    },

    generate : function(outputPath){
        var _self = this;
        var deferred = new DEFERRED;
        PPTX.getPPTXT().slides.forEach(function(slide){
            var cloneSlide = _self.getSlideClone(_self.getSlideCount());
            cloneSlide.content = PPTX.makeXmlSlide(PPTX.getPPTXT().slides[slide.numb - 1]);
            cloneSlide.rel = PPTX.makeXmlSlideRel(slide.numb, _self.getSlideCount());
            newSlides.push(cloneSlide);
        });
        currentPresentation.generate(newSlides).then(function(newPresentation){
            console.log("output file - " + outputPath);
            newPresentation.saveAs(outputPath);
            deferred.resolve(outputPath);
        }).catch(function(err){
            console.error("Error in generating file !!");
            deferred.reject(err);
        });
        return deferred.promise;
    }

};

module.exports = PPTX_GENERATOR;