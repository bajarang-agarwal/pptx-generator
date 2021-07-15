"use strict";

Object.defineProperty(exports, "__esModule", {
  value: true
});
exports["default"] = void 0;

var _entities = _interopRequireDefault(require("entities"));

function _interopRequireDefault(obj) { return obj && obj.__esModule ? obj : { "default": obj }; }

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }

function _defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } }

function _createClass(Constructor, protoProps, staticProps) { if (protoProps) _defineProperties(Constructor.prototype, protoProps); if (staticProps) _defineProperties(Constructor, staticProps); return Constructor; }

var Slide = /*#__PURE__*/function () {
  function Slide(rel, content) {
    _classCallCheck(this, Slide);

    // ppt/slides/_rels/slide(i).xml.rels
    this.rel = rel; // ppt/slides/slide(i).xml

    this.content = content;
  }
  /**
   * 
   */


  _createClass(Slide, [{
    key: "clone",
    value: function clone() {
      return new Slide(this.rel, this.content);
    }
    /**
     * 
     */

  }, {
    key: "fill",
    value: function fill(pair) {
      // 檢查key 和value是否存在
      // 處理 XML Entities
      var value = _entities["default"].encodeXML(pair.value);

      var key = pair.key; // offset: 避免遞迴置換

      var offset = 0;
      var temp = 0; // Replace All

      while ((temp = this.content.indexOf(key, offset)) > -1) {
        this.content = Slide.replace(this.content, offset, key, value);
        offset = temp + value.length;
      }
    }
    /**
     * 
     */

  }, {
    key: "fillAll",
    value: function fillAll(pairs) {
      var _this = this;

      pairs.forEach(function (pair) {
        _this.fill(pair);
      });
    }
    /**
     * 
     */

  }], [{
    key: "replace",
    value: function replace(str, offset, a, b) {
      var index = str.indexOf(a, offset);
      return index > -1 ? str.substring(0, index) + str.substring(index, str.length).replace(a, b) : str;
    }
    /**
     * 
     */

  }, {
    key: "pair",
    value: function pair(key, value) {
      return {
        key: key,
        value: value
      };
    }
  }]);

  return Slide;
}();

exports["default"] = Slide;