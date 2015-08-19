# [sqlxlsx version 1.0.5](https://github.com/bchr02/sqlxlsx)
Create Excel files from SQL queries using the [node-oracledb](https://github.com/oracle/node-oracledb) database driver and the [exceljs](https://github.com/guyonroche/exceljs) streaming XLSX Writer.

[![NPM](https://nodei.co/npm/sqlxlsx.png?downloads=true&stars=true)](https://nodei.co/npm/sqlxlsx/)

## About:
Currently, [sqlxlsx](https://github.com/bchr02/sqlxlsx) only works for Oracle databases but eventually I would like to make it compatible with more. Such as MySQL.

## Prerequisites:
There are several prerequisites needed to both compile and run the Oracle database driver. Please visit the [node-oracledb INSTALL.md page](https://github.com/oracle/node-oracledb/blob/master/INSTALL.md) for more information.

## Installation:
- Run `npm install sqlxlsx` to install from the NPM registry.


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
