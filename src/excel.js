var _ = require('lodash');

if (!String.prototype.endsWith) {
  String.prototype.endsWith = function(searchString, position) {
      var subjectString = this.toString();
      if (typeof position !== 'number' || !isFinite(position) || Math.floor(position) !== position || position > subjectString.length) {
        position = subjectString.length;
      }
      position -= searchString.length;
      var lastIndex = subjectString.indexOf(searchString, position);
      return lastIndex !== -1 && lastIndex === position;
  };
}

var HtmlToExcelBuilder = function(){
    function rgb2hex(rgb) {
            if (/^#[0-9A-F]{6}$/i.test(rgb)) return rgb;

            rgb = rgb.match(/^rgb\((\d+),\s*(\d+),\s*(\d+)\)$/);
            function hex(x) {
                return ("0" + parseInt(x).toString(16)).slice(-2);
            }
            if(rgb && rgb.length==4){
                return hex(rgb[1]) + hex(rgb[2]) + hex(rgb[3]);
            }else{
                return '';
            }
        }

        function b64toBlob(b64Data, contentType, sliceSize) {
          contentType = contentType || '';
          sliceSize = sliceSize || 512;

          var byteCharacters = atob(b64Data);
          var byteArrays = [];

          for (var offset = 0; offset < byteCharacters.length; offset += sliceSize) {
            var slice = byteCharacters.slice(offset, offset + sliceSize);

            var byteNumbers = new Array(slice.length);
            for (var i = 0; i < slice.length; i++) {
              byteNumbers[i] = slice.charCodeAt(i);
            }

            var byteArray = new Uint8Array(byteNumbers);

            byteArrays.push(byteArray);
          }

          var blob = new Blob(byteArrays, {type: contentType});
          return blob;
        }

        var save = function(filename, b64Data) {
            //var blob = new Blob([data], {type: 'text/csv'});
            var blob = b64toBlob(b64Data, 'application/excel');
            if(window.navigator.msSaveOrOpenBlob) {
                window.navigator.msSaveBlob(blob, filename);
            }else{
                var elem = window.document.createElement('a');
                elem.href = window.URL.createObjectURL(blob);
                elem.download = filename;
                document.body.appendChild(elem);
                elem.click();
                document.body.removeChild(elem);
            }
        };

        var getProperties = function(propName, arr){
            if(!arr) return null;
            var propertyValues = [];
            for(var key in arr){
                propertyValues.push(arr[key][propName]);
            }
            return propertyValues;
        };

        var convertColor = function(raw){
            if( raw && raw != 'transparent' ){
                return rgb2hex(raw);
            }
            return null;
        };

        var convertFontSize = function(raw){
            if( raw ){
                if(raw.endsWith("px")){
                    var pxValue = Number(raw.replace("px", ""));
                    return Math.round((pxValue/100)*75);
                }else{
                    return raw.replace("pt", "");
                }
            }
            return null;
        };

        //		var exportCssStyles = ['background-color', 'color', 'font-size'];
        var exportCssStyles = ['background-color', 'color'];

        var fillTemplate = { type: 'pattern',
                        patternType: 'solid'};

        var cssMappings = {'background-color' : {stylePart : 'fill', styleProperty : 'fgColor', converter: convertColor, template: fillTemplate},
            'color' : {stylePart : 'font', styleProperty : 'color', converter: convertColor},
            'font-size' : {stylePart : 'font', styleProperty : 'size', converter: convertFontSize},
        };

        var fontStyles = ['color', 'font-size'];

        var fontTemplate = {
                            color: '000000',
                            size: 11
                        };

        var extractTableData = function(el, sheet){
            var trElements = $(el).find("tr");
            var rows = [];
            $.each(trElements, function(index, trEl){

                var line = [];
                var tdElements = $(trEl).find("td");
                $.each(tdElements, function(index, tdEl){
                    var cellValue = $(tdEl).text();

                    var style = {};

                    for(var key in exportCssStyles){
                        var cssProperty = exportCssStyles[key];
                        var cssValue = $(tdEl).css(cssProperty);
                        var cssMapping = cssMappings[cssProperty];
                        var value = cssMapping.converter(cssValue);
                        if(value){
                            if(!style[cssMapping.stylePart]){
                                if(cssMapping.template){
                                    style[cssMapping.stylePart] = JSON.parse(JSON.stringify(cssMapping.template));
                                }else{
                                    style[cssMapping.stylePart] = {};
                                }
                            }
                            style[cssMapping.stylePart][cssMapping.styleProperty] = value;
                        }
                    }

                    if( !jQuery.isEmptyObject(style) ){
                        var sheetStyle = sheet.createFormat(style);
                        cellValue = {value: cellValue, metadata: {style: sheetStyle.id}};
                    }
                    line.push(cellValue);
                });
                rows.push(line);
            });

            return rows;
        };


        var creationContexts = [];
        var defaultCreationContext = {
                fileName: "excelJS-export.xlsx",
                rowSelector: 'tr',
                columnSelector: 'td',
                style: {}
        };

        this.addCreationContext = function(creationContext){
            creationContexts.push(creationContext);
        };

        var createExcelSheet = function(workbook, creationContext, index){
            var sheetName = creationContext.sheetName ?  creationContext.sheetName :  'Table '+index;
            var worksheet = workbook.createWorksheet({ name: sheetName });
            var stylesheet = workbook.getStyleSheet();
            var rootEl = creationContext.tableRoot ? creationContext.tableRoot : $(creationContext.rootSelector);
            var tableData = extractTableData(rootEl, stylesheet);
            worksheet.sheetView.showGridLines = true;
            worksheet.setData(tableData);
            workbook.addWorksheet(worksheet);

            return worksheet;
        };

        /**
        requires creationContexts being added.
        */
        this.createExcelFile = function(outputFileName){
            var workbook = ExcelBuilder.Builder.createWorkbook();

            $.each(creationContexts, function( index, creationContext ) {
                var mergedContext = $.extend(true, {}, defaultCreationContext, creationContext);
                createExcelSheet(workbook, mergedContext, index);
            });

            ExcelBuilder.Builder.createFile(workbook).then(function (data) {
                save(outputFileName, data);
            });
        };
};


var ExcelJS = {

    createHtmlTableExcelFile : function(tableRootSelector, outputFileName){
            var creationContext = {
                fileName: outputFileName,
                rootSelector: tableRootSelector,
                tableRoot: $(tableRootSelector)
            };
            creationContexts = [];
            var excelBuilder = new HtmlToExcelBuilder();
            excelBuilder.addCreationContext(creationContext);
            return excelBuilder.createExcelFile(outputFileName);
        }
};

try {
    if(typeof window !== 'undefined') {
        window.ExcelJS = ExcelJS;
    }
} catch (e) {
    //Silently ignore?
    console.info("Not attaching EB to window");
}