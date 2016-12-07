/*jslint node:true, unparam:true, nomen:true, regexp:true*/
'use strict';

var xlsx = require('node-xlsx');


module.exports = function (filePath) {
    var startRow = null,
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

        for (i = 0; i < data.length; i += 1) {
            if (data[i][0] === '{build-doc}') {
                startRow = i + 1;
                break;
            }
        }
        if (startRow === null) {
            throw new Error('Unable to find start of build document!');
        }

        for (i = 1; i < data[startRow - 1].length; i += 1) {
            col = data[startRow - 1][i];

            if (!col) { break; }

            col = col + '';
            col = col.trim();

            colTrans.push({
                name: col,
                column: i
            });
        }

        colTrans.forEach(function (column) {
            var columnJson = {},
                line = 0;

            data.forEach(function (row) {
                var array = '',
                    colName,
                    colAddition,
                    colFinal;

                if (line >= startRow && row[colId]) {
                    colName = row[colId].split('.');
                    array = row[colId].match(/\[(.*?)\]/);
                    if (array) {
                        if (!columnJson.hasOwnProperty(array.input.replace(array[0], ''))) {
                            columnJson[array.input.replace(array[0], '')] = {};
                        }
                        columnJson[array.input.replace(array[0], '')][array[1]] = new Object(row[column.column]);
                    } else {
                        colAddition = columnJson;
                        colName.forEach(function (split, i) {
                            if(!colAddition.hasOwnProperty(split)) {
                                colAddition[split] = {};
                            }

                            if (i >= colName.length - 1) {
                                colFinal = split;
                            } else {
                                colAddition = colAddition[split];
                            }
                        });
                        colAddition[colFinal] = new Object(row[column.column]);
                    }
                }
                line += 1;
            });

            addition[column.name] = columnJson;
        });
    });
    return json;
};
