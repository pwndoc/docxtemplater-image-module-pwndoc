"use strict";

var _typeof = typeof Symbol === "function" && typeof Symbol.iterator === "symbol" ? function (obj) { return typeof obj; } : function (obj) { return obj && typeof Symbol === "function" && obj.constructor === Symbol && obj !== Symbol.prototype ? "symbol" : typeof obj; };

var _createClass = function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; }();

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }

var templates = require("./templates");
var DocUtils = require("docxtemplater").DocUtils;
var DOMParser = require("@xmldom/xmldom").DOMParser;

function isNaN(number) {
	return !(number === number);
}

var ImgManager = require("./imgManager");
var moduleName = "open-xml-templating/docxtemplater-image-module";

function getInnerDocx(_ref) {
	var part = _ref.part;

	return part;
}

function getInnerPptx(_ref2) {
	var part = _ref2.part,
	    left = _ref2.left,
	    right = _ref2.right,
	    postparsed = _ref2.postparsed;

	var xmlString = postparsed.slice(left + 1, right).reduce(function (concat, item) {
		return concat + item.value;
	}, "");
	var xmlDoc = new DOMParser().parseFromString("<xml>" + xmlString + "</xml>");
	part.offset = { x: 0, y: 0 };
	part.ext = { cx: 0, cy: 0 };
	var offset = xmlDoc.getElementsByTagName("a:off");
	var ext = xmlDoc.getElementsByTagName("a:ext");
	if (ext.length > 0) {
		part.ext.cx = parseInt(ext[ext.length - 1].getAttribute("cx"), 10);
		part.ext.cy = parseInt(ext[ext.length - 1].getAttribute("cy"), 10);
	}
	if (offset.length > 0) {
		part.offset.x = parseInt(offset[offset.length - 1].getAttribute("x"), 10);
		part.offset.y = parseInt(offset[offset.length - 1].getAttribute("y"), 10);
	}
	return part;
}

var ImageModule = function () {
	function ImageModule(options) {
		_classCallCheck(this, ImageModule);

		this.name = "ImageModule";
		this.options = options || {};
		this.imgManagers = {};
		if (this.options.centered == null) {
			this.options.centered = false;
		}
		if (this.options.getImage == null) {
			throw new Error("You should pass getImage");
		}
		if (this.options.getSize == null) {
			throw new Error("You should pass getSize");
		}
		if (this.options.border && !/^[a-fA-F0-9]{6}$/.test(this.options.border)) {
			throw new Error("border option is incorrect");
		}
		this.imageNumber = 1;
	}

	_createClass(ImageModule, [{
		key: "optionsTransformer",
		value: function optionsTransformer(options, docxtemplater) {
			var relsFiles = docxtemplater.zip.file(/\.xml\.rels/).concat(docxtemplater.zip.file(/\[Content_Types\].xml/)).map(function (file) {
				return file.name;
			});
			this.fileTypeConfig = docxtemplater.fileTypeConfig;
			this.fileType = docxtemplater.fileType;
			this.zip = docxtemplater.zip;
			options.xmlFileNames = options.xmlFileNames.concat(relsFiles);
			return options;
		}
	}, {
		key: "set",
		value: function set(options) {
			if (options.zip) {
				this.zip = options.zip;
			}
			if (options.xmlDocuments) {
				this.xmlDocuments = options.xmlDocuments;
			}
		}
	}, {
		key: "parse",
		value: function parse(placeHolderContent) {
			var module = moduleName;
			var type = "placeholder";
			if (this.options.setParser) {
				return this.options.setParser(placeHolderContent);
			}
			if (placeHolderContent.substring(0, 2) === "%%") {
				return { type: type, value: placeHolderContent.substr(2), module: module, centered: true };
			}
			if (placeHolderContent.substring(0, 1) === "%") {
				return { type: type, value: placeHolderContent.substr(1), module: module, centered: false };
			}
			return null;
		}
	}, {
		key: "postparse",
		value: function postparse(parsed) {
			var expandTo = void 0;
			var getInner = void 0;
			if (this.fileType === "pptx") {
				expandTo = "p:sp";
				getInner = getInnerPptx;
			} else {
				expandTo = this.options.centered ? "w:p" : "w:t";
				getInner = getInnerDocx;
			}
			return DocUtils.traits.expandToOne(parsed, { moduleName: moduleName, getInner: getInner, expandTo: expandTo });
		}
	}, {
		key: "render",
		value: function render(part, options) {
			if (!part.type === "placeholder" || part.module !== moduleName) {
				return null;
			}
			var tagValue = options.scopeManager.getValue(part.value, {
				part: part
			});
			if (!tagValue) {
				return { value: this.fileTypeConfig.tagTextXml };
			} else if ((typeof tagValue === "undefined" ? "undefined" : _typeof(tagValue)) === "object") {
				return this.getRenderedPart(part, tagValue.rId, tagValue.sizePixel);
			}
			var imgManager = new ImgManager(this.zip, options.filePath, this.xmlDocuments, this.fileType);
			var imgBuffer = this.options.getImage(tagValue, part.value);
			var rId = imgManager.addImageRels(this.getNextImageName(), imgBuffer);
			var sizePixel = this.options.getSize(imgBuffer, tagValue, part.value);
			return this.getRenderedPart(part, rId, sizePixel);
		}
	}, {
		key: "resolve",
		value: function resolve(part, options) {
			var _this = this;

			var imgManager = new ImgManager(this.zip, options.filePath, this.xmlDocuments, this.fileType);
			if (!part.type === "placeholder" || part.module !== moduleName) {
				return null;
			}
			var value = options.scopeManager.getValue(part.value, {
				part: part
			});
			if (!value) {
				return { value: this.fileTypeConfig.tagTextXml };
			}
			return new Promise(function (resolve) {
				var imgBuffer = _this.options.getImage(value, part.value);
				resolve(imgBuffer);
			}).then(function (imgBuffer) {
				var rId = imgManager.addImageRels(_this.getNextImageName(), imgBuffer);
				return new Promise(function (resolve) {
					var sizePixel = _this.options.getSize(imgBuffer, value, part.value);
					resolve(sizePixel);
				}).then(function (sizePixel) {
					return {
						rId: rId,
						sizePixel: sizePixel
					};
				});
			});
		}
	}, {
		key: "getRenderedPart",
		value: function getRenderedPart(part, rId, sizePixel) {
			if (isNaN(rId)) {
				throw new Error("rId is NaN, aborting");
			}
			var size = [DocUtils.convertPixelsToEmus(sizePixel[0]), DocUtils.convertPixelsToEmus(sizePixel[1])];
			var centered = this.options.centered || part.centered;
			var border = this.options.border || null;
			var newText = void 0;
			if (this.fileType === "pptx") {
				newText = this.getRenderedPartPptx(part, rId, size, centered);
			} else {
				newText = this.getRenderedPartDocx(rId, size, centered, border);
			}
			return { value: newText };
		}
	}, {
		key: "getRenderedPartPptx",
		value: function getRenderedPartPptx(part, rId, size, centered) {
			var offset = { x: parseInt(part.offset.x, 10), y: parseInt(part.offset.y, 10) };
			var cellCX = parseInt(part.ext.cx, 10) || 1;
			var cellCY = parseInt(part.ext.cy, 10) || 1;
			var imgW = parseInt(size[0], 10) || 1;
			var imgH = parseInt(size[1], 10) || 1;
			if (centered) {
				offset.x = Math.round(offset.x + cellCX / 2 - imgW / 2);
				offset.y = Math.round(offset.y + cellCY / 2 - imgH / 2);
			}
			return templates.getPptxImageXml(rId, [imgW, imgH], offset);
		}
	}, {
		key: "getRenderedPartDocx",
		value: function getRenderedPartDocx(rId, size, centered, border) {
			if (centered) return templates.getImageXmlCentered(rId, size);else if (border) return templates.getImageXmlBordered(rId, size, border);else return templates.getImageXml(rId, size);
		}
	}, {
		key: "getNextImageName",
		value: function getNextImageName() {
			var name = "image_generated_" + this.imageNumber + ".png";
			this.imageNumber++;
			return name;
		}
	}]);

	return ImageModule;
}();

module.exports = ImageModule;