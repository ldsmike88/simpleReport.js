# simpleReport.js
This package utilizes excel4node to generate a simple report. It attempts to streamline the process of getting data from the database to the end user. This is ported from and expanded on simpleReport.py.

## Installation
`npm install simplereportjs`

## Usage
#### Demo
There is only one function that takes one optional object. If you run it without the object it will create a demo file in the current working directory with a timestamp for the filename.
```
var simpleReport = require('simplereportjs');
simpleReport.generate();
```

#### Complete Example
Normally you will provide an object that defines your report.
```
var simpleReport = require('simplereportjs');

const report = {
  name: "My Report Name",
  description: "Here is where you put the report description. It can be as long as you want and the row height will expand to fill the space.\\n\\n It does a pretty good job, but isn't perfect.",
  sheets: [{
    name: "Sheet One",
    description: "",
    printReportName: true,
    printSheetName: false,
    tables: [{
        name: "Sample Table",
        description: "",
        printTableName: true,
        rowBanding: true,
        rowBandingColors: ['#333333', '#FFFFFF'],
        rowBandingFontColors: ['#FFFFFF', '#333333'],
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
        filters: true,
        totalsRow: {
          A: 'Totals',
          B: 2,
          D: 109,
        },
      },
      {
        name: "Table Two",
        description: "Here's another description",
        printTableName: true,
        filters: true,
        data: [],
      }, {
        name: "Table Three",
        description: "",
        printTableName: false,
        rowBanding: true,
        rowBandingColors: ['#8EDEFD', '#31C5FD'],
        data: [{
            'column A': 'One',
            'date': new Date(),
            'category': 'Trees',
            '$total': 64,
          },
          {
            'column A': 'Two',
            'date': new Date(),
            'category': 'Birds',
            '$total': 159,
          },
          {
            'column A': 'Three',
            'date': new Date(),
            'category': 'Alligators',
            '$total': 78541236.02156,
          },
          {
            'column A': 'Four',
            'date': new Date(),
            'category': 'Cars',
            '$total': 1587,
          },
          {
            'column A': 'Five',
            'date': new Date(),
            'category': 'Eyes',
            '$total': 20.25,
          },
        ],
        filters: true,
        totalsRow: {
          A: 'Totals',
          B: 2,
          D: 105,
        },
      }
    ]
  }, {
    name: "Sheet Two",
    description: "",
    printReportName: false,
    printSheetName: true,
    tables: [{
      name: "Table Four",
      description: "The description for table four.",
      printTableName: true,
      filters: false,
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
    }]
  }],
}
```

#### Parameters
All keys are optional.
```
{
  name: String,
  description: String,
  sheets: [ // Array of objects
    {
      name: String,
      description: String,
      printReportName: Boolean, // Defaults to false
      printSheetName: Boolean, // Defaults to false
      tables: [ // Array of objects
        {
          name: String,
          description: String,
          printTableName: Boolean, // Defaults to false
          rowBanding: Boolean, // Defaults to false
          rowBandingColors: Array of Strings, // Example: ['#333333', '#FFFFFF']
          rowBandingFontColors: Array of Strings, // Example: ['#FFFFFF', '#333333']
          data: [ // Array of objects
            // Each object represents one row of data.
            // Each object should have the same keys.
            // An object can have multiple key (column name) value pairs.
            {
              'Column Name': String or Date or Number or Boolean,
            },
          ],
          filters: Boolean, // Defaults to false
          totalsRow: {
            Letter: String or Integer,
            // Letter should be the letter of the column.
            // If the value is Integer, the cell will be a formula of
            // "SUBTOTAL(Integer, Range)" (Range is calculated automatically.)
          },
        },
      ]
    },
  ],
}
```

## Limitations
- The tables are not actually formatted as Excel tables, they are just ranges (excel4node limitation). Because of this there can only be one filtered range on any given sheet (MS Excel limitation). If more than one range is specified only the last one will have filters.
