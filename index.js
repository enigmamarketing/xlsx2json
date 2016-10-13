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
        var i = 0,
            colTrans = [],
            data = worksheet.data,
            sheetSplit = worksheet.name.split('.'),
            addition;

        addition = json;
        sheetSplit.forEach(function (split) {
            addition[split.toLowerCase()] = {};

            addition = addition[split.toLowerCase()];
        });

        data[0].forEach(function (col) {
            if (i > 0) {
                colTrans.push({
                    name: col,
                    column: i
                });
            }
            i++;
        });

        colTrans.forEach(function (column) {
            var columnJson = {};

            data.forEach(function (row) {
                var array = '';
                line = 0;
                if (line >= startRow && row[colId]) {
                    array = row[colId].match(/\[(.*?)\]/);
                    if (array) {
                        if (!columnJson.hasOwnProperty(array.input.replace(array[0], ''))) {
                            columnJson[array.input.replace(array[0], '')] = {};
                        }
                        columnJson[array.input.replace(array[0], '')][array[1]] = row[column.column];
                    } else {
                        columnJson[row[colId]] = row[column.column];
                    }
                }
                line += 1;
            });

            addition[column.name.toLowerCase()] = columnJson;
        });
    });
    return json;
};
