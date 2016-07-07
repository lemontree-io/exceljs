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
if (!String.prototype.hashCode) {
    String.prototype.hashCode = function() {
        var hash = 0, i, chr, len;
        if (this.length === 0) return hash;
        for (i = 0, len = this.length; i < len; i++) {
            chr   = this.charCodeAt(i);
            hash  = ((hash << 5) - hash) + chr;
            hash |= 0; // Convert to 32bit integer
        }
        return hash;
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

        var fillTemplate = { type: 'pattern',
                        patternType: 'solid'};

        var borderTemplate = {top : { style: 'thin' }, right : { style: 'thin' }, bottom : { style: 'thin' }, left : { style: 'thin' }};

        var cssMappings = {
            'background-color' : {stylePart : 'fill', styleProperty : 'fgColor', converter: convertColor, template: fillTemplate},
            'color' : {stylePart : 'font', styleProperty : 'color', converter: convertColor},
            'font-size' : {stylePart : 'font', styleProperty : 'size', converter: convertFontSize},
            'border-top-color' : {stylePart : 'border', styleProperty : 'top.color', converter: convertColor, template: borderTemplate},
            'border-right-color' : {stylePart : 'border', styleProperty : 'right.color', converter: convertColor, template: borderTemplate},
            'border-bottom-color' : {stylePart : 'border', styleProperty : 'bottom.color', converter: convertColor, template: borderTemplate},
            'border-left-color' : {stylePart : 'border', styleProperty : 'left.color', converter: convertColor, template: borderTemplate}
        };

        var creationContexts = [];
        var defaultCreationContext = {
                sheetName: null,
                rowSelector: 'tr',
                columnSelector: 'td',
                autoSizeColumns: true,
                maxColumnWidth: 40,
                minColumnWidth: 8,
                style: {},
                includeCssProperties: ['background-color', 'color',  'border-top-color', 'border-right-color', 'border-bottom-color', 'border-left-color']
        } ;

        var defaultCssStyle = {
                            'color': '#000000',
                            'font-size': 11
                        };

        var calculateExcelColumnChar = function( rowIndex, colIndex ){

            var nextValue = colIndex + 1;
            var excelColumnString = '';

            while(nextValue > 0){
                var nextOffset = (nextValue - 1) % 26;
                var nextCharCode = Number('A'.charCodeAt(0)) + Number(nextOffset);
                excelColumnString = String.fromCharCode(nextCharCode) + excelColumnString;
                nextValue = Math.floor((nextValue -  nextOffset) / 26);
            }
            var excelRowIndex = rowIndex + 1;

            return excelColumnString + '' + excelRowIndex;

        };
        var calculateMaxCharacterWidthPerColumn = function(cells){
            var maxCharactersPerColumn = [];
            for(var i = 0; i < cells.length; i++){
                for(var j = 0; j < cells[i].length; j++){
                    var cellValue = cells[i][j];
                    if(typeof cellValue !== 'string'){
                        cellValue = cellValue.value;
                    }
                    var cellLines = cellValue.split(/\r?\n/);
                    for(var key in  cellLines){
                        var lineValue = cellLines[key].trim();
                        if((typeof maxCharactersPerColumn[j] === 'undefined') || maxCharactersPerColumn[j] < lineValue.length){
                            maxCharactersPerColumn[j] = lineValue.length;
                        }
                    }
                }
            }
            return maxCharactersPerColumn;
        };

        var calculateStyleHashCode = function(style){

            var strVal = '';
            for(var stylePartKey in style){
                var stylePartArr = style[stylePartKey];
                for(var stylePropertyKey in stylePartArr){
                    strVal += stylePartKey+":"+stylePropertyKey+"="+stylePartArr[stylePropertyKey]+";";
                 }
            }
            return strVal.hashCode();
        };

        var convertCellStyle = function(tdEl, creationContext, styleSheet, styles){
            var style = {};

            for(var key in creationContext.includeCssProperties){
                var cssProperty = creationContext.includeCssProperties[key];
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
                    var stylePropertyChain = cssMapping.styleProperty.split('.');
                    var styleObjectParent = style[cssMapping.stylePart];
                    for(var i = 0; i < stylePropertyChain.length; i++){
                        if(i === (stylePropertyChain.length - 1)){
                            styleObjectParent[stylePropertyChain[i]] = value;
                        }else{
                            if(typeof style[cssMapping.stylePart] === 'undefined') {
                               styleObjectParent[stylePropertyChain[i]] = {};
                            }
                            styleObjectParent = styleObjectParent[stylePropertyChain[i]];
                        }
                    }
                }
            }
            var sheetStyle;
            if( !jQuery.isEmptyObject(style) ){
                var styleHash = calculateStyleHashCode(style);
                if(typeof styles[styleHash] !== 'undefined'){
                    sheetStyle = styles[styleHash];
                }else{
                    sheetStyle = styleSheet.createFormat(style);
                    styles[styleHash] = sheetStyle;
                }
            }
            return sheetStyle;
        };

        var createCell = function(cellValue, sheetStyle){
            return {value: cellValue, metadata: {style: sheetStyle.id}};
        };

        var extractTableData = function(el, workSheet, styleSheet, creationContext){
            var id = "excel-js-export-tmp-table-ikj3usd0bg3ukujd5s";
            $("body").append("<div id="+id+" style='visibility: hidden;'></div>");
            $('#'+id).append(el);

            var trElements = $(el).find(creationContext.rowSelector);
            var cells = [];
            var styles = [];
            $.each(trElements, function(rowIndex, trEl){

                var tdElements = $(trEl).find(creationContext.columnSelector);

                cells[rowIndex] = [];
                var offsetColIndex = -1;
                console.log("reading rowIndex "+rowIndex);
                $.each(tdElements, function(colIndex, tdEl){
                    offsetColIndex++;
                    while(typeof cells[rowIndex][offsetColIndex] !== 'undefined'){ //already set by previous colspan cells detection
                        offsetColIndex++;
                    }
                    console.log("reading colIndex "+rowIndex);
                    var cellValue = $(tdEl).text();

                    var colspan = $(tdEl).attr("colspan");
                    colspan = colspan ? colspan : 1;

                    var rowspan = $(tdEl).attr("rowspan");
                    rowspan = rowspan ? rowspan : 1;

                    var sheetStyle = convertCellStyle(tdEl, creationContext, styleSheet, styles);

                    if(colspan > 1 || rowspan > 1){

                        var startCellId = calculateExcelColumnChar(rowIndex, offsetColIndex);

                        var endColIndex = Number(offsetColIndex) + Number(colspan) - 1;
                        var endRowIndex = Number(rowIndex) + Number(rowspan) - 1;
                        var endCellId = calculateExcelColumnChar(endRowIndex, endColIndex);
                        workSheet.mergeCells(startCellId, endCellId);

                        //init empty cell values for rowspan/colspan cells

                        for(var i = rowIndex; i <= endRowIndex; i++){
                            for(var j = offsetColIndex; j <= endColIndex; j++){
                                if(!(i === rowIndex && j === offsetColIndex)){
                                    cells[i][j] = createCell('', sheetStyle);
                                }
                            }
                        }
                    }

                    cellValue = createCell(cellValue, sheetStyle);
                    cells[rowIndex][offsetColIndex] = cellValue;
                });

            });
            $('#'+id).remove();
            return cells;
        };

        var autosizeColumns = function(cells, worksheet, creationContext){
            if(creationContext.autoSizeColumns){
                var maxCharactersPerColumn = calculateMaxCharacterWidthPerColumn(cells);
                var columnWidthStyles = [];
                for(var key in maxCharactersPerColumn){

                    var columnWidth = maxCharactersPerColumn[key];
                    if(typeof creationContext.minColumnWidth !== 'undefined'){
                        columnWidth = Math.max(columnWidth, creationContext.minColumnWidth);
                    }
                    if(typeof creationContext.maxColumnWidth !== 'undefined'){
                        columnWidth = Math.min(columnWidth, creationContext.maxColumnWidth);
                    }
                    columnWidthStyles.push({ width: columnWidth });
                }
                worksheet.setColumns(columnWidthStyles);
            }
        };




        this.addCreationContext = function(creationContext){
            creationContexts.push(creationContext);
        };

        this.addCreationContexts = function(creationContextsParam){
            for(var key in creationContextsParam){
                creationContexts.push(creationContextsParam[key]);
            }
        };

        var createExcelSheet = function(workbook, creationContext, index){
            var sheetName = creationContext.sheetName ?  creationContext.sheetName :  'Table '+index;
            var worksheet = workbook.createWorksheet({ name: sheetName });
            var stylesheet = workbook.getStyleSheet();
            var tableCellData = extractTableData(creationContext.tableRoot, worksheet, stylesheet, creationContext);
            worksheet.sheetView.showGridLines = true;
            worksheet.setData(tableCellData);
            autosizeColumns(tableCellData, worksheet, creationContext);
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

    createExcelFile : function(creationContexts, outputFileName){
        var excelBuilder = new HtmlToExcelBuilder();
        excelBuilder.addCreationContexts(creationContexts);
        return excelBuilder.createExcelFile(outputFileName);
    },

    createHtmlTableExcelFile : function ( tableRootSelector, outputFileName ) {
        var creationContext;
        if(typeof tableRootSelector === 'string'){
            creationContext = {
                tableRoot: $(tableRootSelector)
            };
        }else{
            creationContext = {
                tableRoot: tableRootSelector
            };
        }
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