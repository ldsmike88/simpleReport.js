var xl = require('excel4node');
var strftime = require('strftime');

const generate = (report={}) => {
  return new Promise((resolve, reject) => {

    // Create new workbook
    wb = new xl.Workbook({
      defaultFont: {
        size: 11,
      },
      dateFormat: 'm/d/yyyy',
      author: 'Simple Report',
    });

    // Define formats
    var reportNameFormat = wb.createStyle({
      font: { size: 24 }
    })
    var sheetNameFormat = wb.createStyle({
      font: { size: 18 }
    })
    var tableNameFormat = wb.createStyle({
      font: { size: 16 }
    })
    var descFormat = wb.createStyle({
      alignment: {
        wrapText: true,
        vertical: 'top',
      }
    })
    var bold = wb.createStyle({
      font: { bold: true },
    });
    var numberFormat = wb.createStyle({
      numberFormat: '#,##0.00; (#,##0.00); -'
    });
    var currencyFormat = wb.createStyle({
      numberFormat: '$#,##0.00; ($#,##0.00); -'
    });
    var tableHeaderFormat = wb.createStyle({
      border: {
        bottom: {
          style: 'thin',
          color: 'black',
        }
      }
    });
    var tableFooterFormat = wb.createStyle({
      border: {
        top: {
          style: 'thin',
          color: 'black',
        }
      }
    });

    // Set standard worksheet options
    const wsOptions = {
      'sheetFormat': {
        'defaultColWidth': 10,
        'defaultRowHeight': 15,
      },
    }

  if (!report.sheets) {
    report.sheets = [{
        name: 'simpleReport.js Demo',
        description: 'You\'re seeing this demo page because report.sheets ' +
          'was not defined. To replace this with your content, try adding:\n ' +
          'report.sheets = [{ name: \'Sheet Name\' }]',
        printSheetName: true,
        "tables": [{
            name: "Simple Table",
            description: "This is a simple table with row banding.",
            printTableName: true,
            rowBanding: true,
            rowBandingColors: ['#CCCCCC', '#FFFFFF'],
            data: [{
                'column A': 'One',
                'date': new Date(),
                'category': 'Trees',
                'total': 64,
              },
              {
                'column A': 'Two',
                'date': new Date(),
                'category': 'Birds',
                'total': 159,
              },
              {
                'column A': 'Three',
                'date': new Date(),
                'category': 'Alligators',
                'total': 78541236.02156,
              },
              {
                'column A': 'Four',
                'date': new Date(),
                'category': 'Cars',
                'total': 1587,
              },
              {
                'column A': 'Five',
                'date': new Date(),
                'category': 'Eyes',
                'total': 20.25,
              },
            ],
            "filters": true,
          }],
        }]
    }

    // Loop through sheets
    report.sheets.map(sheet => {

      var ws = wb.addWorksheet(sheet.name, wsOptions); // add sheet

      var row = 1; // track of where the next row of data goes
      var maxWidth = {}; // track what the column widths should be
      var descRows = {}; // track the description rows for setting the height

      // Write the report name, description, & date if the option is selected
      if (sheet.printReportName) {
        // Write report name
        ws.cell(row, 1, row, 5, true)
          .string(report.name)
          .style(reportNameFormat);
        ws.row(row).setHeight(31.5);
        row++;

        // Write report description if defined
        if (report.description) {
          ws.cell(row, 1, row, 5, true)
            .string(report.description)
            .style(descFormat);
          descRows[row] = report.description;
          row++;
        } // End write report description

        // Write report date
        var now = new Date();
        ws.cell(row, 1, row, 5, true)
          .string(strftime('Report generated on %m/%d/%Y at %I:%M:%S %p'));
        row += 2;
      } // End write report name, description, date

      // Write sheet name & description if option selected
      if (sheet.printSheetName) {
        // Write sheet name
        ws.cell(row, 1, row, 5, true)
          .string(sheet.name)
          .style(sheetNameFormat);
        ws.row(row).setHeight(23.25);
        row++;

        // Write sheet description if defined
        if (sheet.description) {
          ws.cell(row, 1, row, 5, true)
            .string(sheet.description)
            .style(descFormat);
          descRows[row] = sheet.description
          row++;
        } // End write sheet description

        row++; // add blank row between sheet name & description and the table
      } // End write sheet name & description if option selected


      // Loop through tables if tables exist
      sheet.tables && sheet.tables.map(table => {
        // Write table name & description if that option is selected
        if (table.printTableName) {
          // Write table name
          ws.cell(row, 1, row, 5, true)
            .string(table.name)
            .style(tableNameFormat);
          ws.row(row).setHeight(21);
          row++;

          // Write table description
          if (table.description) {
            ws.cell(row, 1, row, 5, true)
              .string(table.description)
              .style(descFormat);
            descRows[row] = table.description;
            row++;
          }
        } // End write table name & description

        const filtersStart = row // Save where the table starts for row banding

        // If there's no data in the table display "No data"
        if (table.data === undefined || table.data.length == 0) {
          ws.cell(row, 1, row, 5, true)
            .string('No data')
            .style(tableFooterFormat);
        } else { // Otherwise write row
          columnNames = Object.keys(table.data[0]); // Get column names
          // Loop through column names
          for (let [idx, colName] of columnNames.entries()) {
            if (colName.startsWith('$')) {
              // If a column name starts with $ the $ will be removed from the
              // column name and the column values will use currencyFormat
              colName = colName.replace('$', '');
            } else if (colName.startsWith('\\$')) {
              // If a column name starts with \$ the backslash will be removed
              // from the column name and the column values will use
              // numberFormat (NOT currencyFormat)
              colName = colName.replace('\\$', '$');
            }

            // Write column name
            ws.cell(row, idx + 1).string(colName).style(tableHeaderFormat);
            maxWidth[idx + 1] = colName.length + 4; // + 4 for filter button
          } // End loop through column names
          row++;

          // Loop through table rows
          for (dataRow of table.data) {
            var col = 1; // Start with column 1
            // Loop through columns of the table row
            for (pair of Object.entries(dataRow)) {
              const [name, item] = pair;
              // Write the data with the appropriate type
              if (typeof item === 'number') {
                ws.cell(row, col)
                  .number(item)
                  .style(name.startsWith('$') ? currencyFormat : numberFormat);
                maxWidth[col] = Math.max(maxWidth[col],
                  item.toString().length,
                  10);
              } else if (typeof item === 'string') {
                ws.cell(row, col).string(item);
                maxWidth[col] = Math.max(maxWidth[col], item.length);
              } else if (typeof item === 'boolean') {
                ws.cell(row, col).bool(item);
                maxWidth[col] = Math.max(maxWidth[col], 9);
              } else if (item instanceof Date) {
                ws.cell(row, col).date(item);
                maxWidth[col] = Math.max(maxWidth[col], 12);
              }
              col++;
            } // End loop through columns of the table row
            row++;
          } // End loop through table rows

          // Add table filters if option selected
          if (table.filters) {
            ws.row(filtersStart).filter();
          }

          // Get start and end cell references for the table
          var startCell = xl.getExcelCellRef(filtersStart + 1, 1);
          var endCell = xl.getExcelCellRef(row - 1, col - 1);

          // Apply row banding if option selected
          if ('rowBanding' in table && table.rowBanding) {
            // Get first row odd or even for consistent row banding starting color
            const oddOrEven = row % 2 == 0 ? 0 : 1;

            // Set font colors. Default to black.
            var fontColors = 'rowBandingFontColors' in table ?
              table.rowBandingFontColors :
              ['#000000', '#000000'];

            // Apply light color row banding
            ws.cell(filtersStart + 1, 1, row - 1, col - 1, false)
              .style({
                fill: {
                  type: 'pattern',
                  patternType: 'solid',
                  fgColor: table.rowBandingColors[0],
                },
                font: { color: fontColors[0] }
              });

            // Apply dark color row banding
            ws.addConditionalFormattingRule(startCell + ':' + endCell, {
              type: 'expression',
              priority: 1,
              formula: 'MOD(ROW(), 2)=' + oddOrEven, // Even rows
              style: wb.createStyle({
                fill: {
                  type: 'pattern',
                  patternType: 'solid',
                  bgColor: table.rowBandingColors[1],
                },
                font: { color: fontColors[1] }
              }),
            })
          }
          // Add totals row if option selected
          if ('totalsRow' in table && table.totalsRow) {
            // hide a row between the data and the totals so that the totals
            // are not filtered with the filter buttons
            ws.row(row).hide();
            row++;

            for (const column of Object.keys(table.totalsRow)) {
              if (typeof table.totalsRow[column] === 'string') {
                ws.cell(row, xl.getExcelRowCol(column + row).col)
                  .string(table.totalsRow[column])
              } else {
                const colNum = xl.getExcelRowCol(column + 1).col - 1;
                const colName = Object.keys(table.data[0])[colNum];
                ws.cell(row, xl.getExcelRowCol(column + 1).col)
                  .formula(
                    'SUBTOTAL(' + table.totalsRow[column] + ', ' +
                    column + filtersStart +':' +
                    column + (row - 2) + ')'
                  )
                  .style(colName.startsWith('$') ?
                    currencyFormat :
                    numberFormat);
              }
            }

            // format as table footer
            ws.cell(
                row, xl.getExcelRowCol(startCell + 1).col,
                row, xl.getExcelRowCol(endCell + 1).col,
                false
              )
              .style(tableFooterFormat);
          }

        } // End otherwise write row
        row += 2;
      }) // End tables loop

      // Auto width all table columns
      for (col in maxWidth) {
        ws.column(col).setWidth(maxWidth[col]);
      }

      // Set description cells height
      // First, we add up the widths of the merged columns. In Excel, the width
      // of a column is roughly how many characters can fit in that column.
      //
      // With this in mind, we split the description by each line break and then
      // divide the number of characters by the width of the merged columns. A
      // single row of text is 15 high. We multiply the result by 15 to get the
      // appropriate number of rows. By splitting the description into chunks by
      // each line break we garuntee that each will get at least 15 height added
      var descWidth = 0;
      for (var i = 1; i < 6; i++) descWidth += i in maxWidth ? maxWidth[i] : 9

      for (const [descRow, desc] of Object.entries(descRows)) {
        var rowHight = 0;
        for (const descChunk of desc.split('\n')) {
          rowHight += Math.ceil(Math.max(descChunk.length, 1) / descWidth) * 15
        }
        ws.row(parseInt(descRow, 10)).setHeight(rowHight);
      }
    }) // End sheets loop

    // If filename is specified default to timestamp. Ex: 2018-12-26 145716.xlsx
    if (!report.filename) {
      	var d = new Date();
      	var timestamp = (
      	  d.getFullYear().toString() +
      	  "-" + ((d.getMonth() + 1).toString().length == 2 ? (d.getMonth() + 1).toString() : "0" + (d.getMonth() + 1).toString()) +
      	  "-" + (d.getDate().toString().length == 2 ? d.getDate().toString() : "0" + d.getDate().toString()) +
      	  " " + (d.getHours().toString().length == 2 ? d.getHours().toString() : "0" + d.getHours().toString()) +
      	  (d.getMinutes().toString().length == 2 ? d.getMinutes().toString() : "0" + d.getMinutes().toString()) +
      	  (d.getSeconds().toString().length == 2 ? d.getSeconds().toString() : "0" + d.getSeconds().toString())
      	);
      report.filename = timestamp + '.xlsx';
      console.log('Saving ' + report.filename);
    }

    // Save report
    wb.write(report.filename, (err, stats) => {
      if (err) {
        reject(err);
      } else {
        resolve();
      }
    });
  })
}

module.exports = { generate }
