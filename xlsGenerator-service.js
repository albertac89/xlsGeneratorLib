(function () {
    'use strict';
    angular.module('App.xlsGenerator')
    /**
     * @ngdoc service
     * @name App.xlsGenerator.service:xlsGeneratorFactory
     * @description
     * Tools to generate a xls
     * @requires $log
     * @requires $window
     * @requires $filter
     * @requires timeZoneutfOffset
     */
        .factory('xlsGeneratorFactory',
            function ($log, $window, $filter, timeZoneutfOffset) {

                $log.debug('xlsGeneratorFactory loading');

                var factory = {};
                var headerTableColor = '#CCCCFF';

                var mimeType = 'application/vnd.ms-excel;';
                var charset = 'charset=utf-8;';
                var uri = 'data:' + mimeType + charset + 'base64,';
                var template = '<html xmlns:html="http://www.w3.org/TR/REC-html40" ' +
                    // 'xmlns:xsl="http://www.w3.org/1999/XSL/Transform" ' +
                    'xmlns="urn:schemas-microsoft-com:excel:spreadsheet" ' +
                    // 'xmlns:o="urn:schemas-microsoft-com:office:office" ' +
                    // 'xmlns:x="urn:schemas-microsoft-com:office:excel" ' +
                    // 'xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"' +
                    '><head><!--[if gte mso 9]><xml><x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet><x:Name>{worksheet}' +
                    '</x:Name><x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions></x:ExcelWorksheet></x:ExcelWorksheets></x:ExcelWorkbook></xml><![endif]-->';
                var base64=function(s){return $window.btoa(s);};
                var format=function(s,c){return s.replace(/{(\w+)}/g,function(m,p){return c[p];});};

                /**
                 * @ngdoc method
                 * @methodOf App.xlsGenerator.service:xlsGeneratorFactory
                 * @name generateTable
                 * @description
                 * Creates and download the file
                 * @param {object} columnDefs The columns of the table.
                 * @param {object} data The data of the table.
                 */
                factory.generateTable = function (columnDefs, data) {
                    var table = document.createElement('table');
                    table.border = 1;

                    //Add headers
                    var thead = document.createElement('thead');
                    var tr = document.createElement('tr');
                    tr.style.background = headerTableColor;
                    angular.forEach(columnDefs, function (column) {
                        if (column.displayName && column.name === undefined) {
                            var th = document.createElement('th');
                            th.innerHTML = $filter('translate')(column.displayName);
                            tr.appendChild(th);
                        }
                    });
                    thead.appendChild(tr);
                    table.appendChild(thead);

                    //Add data
                    var tbody = document.createElement('tbody');
                    angular.forEach(data, function (row) {
                        var tr = document.createElement('tr');
                        angular.forEach(columnDefs, function (column) {
                            if (column.displayName && column.name === undefined) {
                                var td = document.createElement('td');
                                //Fix number to long
                                if (column.field === 'Ref') {
                                    td.setAttribute('style', 'mso-number-format: 0; text-align: center;');
                                }
                                if (column.cellFilter) {
                                    td.innerHTML = $filter(column.cellFilter.split(':')[0])(row[column.field],
                                        column.cellFilter.split(':')[1] !== undefined ? column.cellFilter.split(':')[1] : '');
                                } else {
                                    td.innerHTML = row[column.field];
                                }
                                tr.appendChild(td);
                            }
                        });
                        tbody.appendChild(tr);
                    });
                    table.appendChild(tbody);

                    //Wrap table
                    var wrap = document.createElement('div');
                    wrap.appendChild(table.cloneNode(true));

                    return wrap.innerHTML;
                };

                /**
                 * @ngdoc method
                 * @methodOf App.xlsGenerator.service:xlsGeneratorFactory
                 * @name saveFile
                 * @description
                 * Creates and download the file
                 * @param {object} customTemplate The template of the xls.
                 * @param {object} ctx The params of the xls.
                 */
                factory.saveFile = function (customTemplate, ctx) {
                    if (navigator.msSaveOrOpenBlob) {
                        navigator.msSaveOrOpenBlob(tableToExcelIE(customTemplate, ctx), generateXLSName(ctx.title));
                    } else {
                        var a = document.createElement('a');
                        a.id = 'xlsButton';
                        a.href = tableToExcel(customTemplate, ctx);
                        a.download = generateXLSName(ctx.title);
                        document.body.appendChild(a);
                        a.click();
                        setTimeout(function() {
                            document.body.removeChild(a);
                            $window.URL.revokeObjectURL(ctx.title);
                        }, 0);
                    }
                };

                return factory;

                /**
                 * @ngdoc method
                 * @methodOf App.xlsGenerator.service:xlsGeneratorFactory
                 * @name generateXLSName
                 * @description
                 * Generates a xls name based on a document title and the current date and time (with a predefined utf offSet)
                 * @param {string} title The base document title.
                 */
                function generateXLSName(title) {
                    //Parse the title from: 'lorem ipsum sit amet' to 'LoremIpsumSitAmet'
                    var parsedTitle = _.upperFirst(_.camelCase(title));
                    //Get the current date and time with a predefined offset Example  '30 09 2016 031204'
                    var stringDate = moment().utcOffset(timeZoneutfOffset).format('DD MM YYYY HHmmss');

                    return parsedTitle + '_' + _.snakeCase(stringDate) + '.xls';
                }

                /**
                 * @ngdoc method
                 * @methodOf App.xlsGenerator.service:xlsGeneratorFactory
                 * @name tableToExcel
                 * @description
                 * Generates a xls name based on a document title and the current date and time (with a predefined utf offSet)
                 * @param {object} customTemplate The template of the xls.
                 * @param {object} ctx The params of the xls.
                 */
                function tableToExcel(customTemplate, ctx) {
                    return uri + base64(format(template + customTemplate, ctx));
                }

                /**
                 * @ngdoc method
                 * @methodOf App.xlsGenerator.service:xlsGeneratorFactory
                 * @name tableToExcelIE
                 * @description
                 * Generates a xls name based on a document title and the current date and time (with a predefined utf offSet)
                 * @param {object} customTemplate The template of the xls.
                 * @param {object} ctx The params of the xls.
                 */
                function tableToExcelIE(customTemplate, ctx) {
                    return new Blob(['\ufeff' + format(template + customTemplate, ctx)], { type: mimeType + charset});
                }
            });
}());
