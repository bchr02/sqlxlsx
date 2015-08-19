# [sqlxlsx version 1.0.3](https://github.com/bchr02/sqlxlsx)
Create Excel files from SQL queries using [node-oracledb](https://github.com/oracle/node-oracledb) database driver and [exceljs](https://github.com/guyonroche/exceljs) (Streaming XLSX Writer)!

[![NPM](https://nodei.co/npm/sqlxlsx.png?downloads=true&stars=true)](https://nodei.co/npm/sqlxlsx/)

## Installation:
Run `npm install sqlxlsx` to install from the NPM registry.


## Sample Usage:
````javascript
var Sqlxlsx = require('sqlxlsx'),
    sqlxlsx = new Sqlxlsx({
      sql: 'select somedate from sometable',
      oracledb_cfg: {
        user: 'username',
        pass: 'password',
        connectString: '192.168.1.254/YOGA' // Easy Connect syntax
      },
      oracledb_args: [],
      numRows: 100, // This is the number of rows to pull/stream at a time
      exceljs_wsName: 'worksheet name',
      exceljs_options: {	filename: "report.xlsx",
				useStyles: false,
				useSharedStrings: true,
				creator: "Name"
			}});
		
sqlxlsx.run(function(err) {console.log('report complete!');});

````
