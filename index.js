/*jslint node:true, unparam:true, nomen:true, regexp:true*/
'use strict';

var xlsx = require('node-xlsx'),
    DataObjectParser = require('dataobject-parser');


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
        sheetSplit.forEach(function (sheetComponent) {
            if(!addition.hasOwnProperty(sheetComponent)) {
                addition[sheetComponent] = {};
            }
            addition = addition[sheetComponent];
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
            var columnJson = new DataObjectParser();

            data.forEach(function (row, i) {
                if (i >= startRow && row[colId]) {
                    columnJson.set(row[colId], new Object(row[column.column]));
                }
            });

            addition[column.name] = columnJson.data();
        });
    });
    return json;
};
