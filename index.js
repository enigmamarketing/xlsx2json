/*jslint node:true, unparam:true, nomen:true, regexp:true*/
'use strict';

var xlsx = require('node-xlsx');


module.exports = function (filePath) {
    // get POST data

    var startRow = 0,
        colId = 0,
        obj = xlsx.parse(filePath),
        line = 0,
        json = [];

    obj.forEach(function (worksheet) {
        var i = 0, col,
            colTrans = [],
            data = worksheet.data,
            sheetSplit = worksheet.name.split('.'),
            addition;

        addition = json;
        sheetSplit.forEach(function (split) {
            if(!addition[split.toLowerCase()]) {
                addition[split.toLowerCase()] = {};
            }
            addition = addition[split.toLowerCase()];
        });
        
        for (i = 1; i < data[0].length; i += 1) {
            col = data[0][i];
            
            if (!col) { break; }
            
            col = col + '';
            col = col.trim();
            
            colTrans.push({
                name: col,
                column: i
            });
        }

        colTrans.forEach(function (column) {
            var columnJson = {};

            data.forEach(function (row) {
                var array = '',
                    colName,
                    colAddition,
                    colFinal,
                    colLevel = 0;
                line = 0;
                if (line >= startRow && row[colId]) {
                    colName = row[colId].split('.');
                    array = row[colId].match(/\[(.*?)\]/);
                    if (array) {
                        if (!columnJson.hasOwnProperty(array.input.replace(array[0], ''))) {
                            columnJson[array.input.replace(array[0], '')] = {};
                        }
                        columnJson[array.input.replace(array[0], '')][array[1]] = row[column.column];
                    } else {

                        colAddition = columnJson;
                        colName.forEach(function (split) {

                            if(!colAddition[split]) {
                                colAddition[split] = {};
                            }
                            if (colLevel === (colName.length - 1)) {
                                colFinal = split;
                            } else {
                                colAddition = colAddition[split];
                            }
                            colLevel += 1;
                        });
                        colAddition[colFinal] = row[column.column];
                    }
                }
                line += 1;
            });

            addition[column.name] = columnJson;
        });
    });
    return json;
};
