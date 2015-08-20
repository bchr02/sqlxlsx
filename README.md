# [sqlxlsx version 1.0.5](https://github.com/bchr02/sqlxlsx)
Create Excel files from SQL queries using the [node-oracledb](https://github.com/oracle/node-oracledb) database driver and the [exceljs](https://github.com/guyonroche/exceljs) streaming XLSX Writer.

[![NPM](https://nodei.co/npm/sqlxlsx.png?downloads=true&stars=true)](https://nodei.co/npm/sqlxlsx/)

## About
Currently, [sqlxlsx](https://github.com/bchr02/sqlxlsx) only works for Oracle databases but eventually I would like to make it compatible with more. Such as MySQL.

## Prerequisites
There are several prerequisites needed to both compile and run the Oracle database driver. Please visit the [node-oracledb INSTALL.md page](https://github.com/oracle/node-oracledb/blob/master/INSTALL.md) for more information.

## Installation
- Run `npm install sqlxlsx` to install from the NPM registry.

## Interface
```javascript
var Sqlxlsx = require('sqlxlsx');
```

## Create a Report
```javascript
var sqlxlsx = new Sqlxlsx();
```

## Define or Change Report Parameters
````javascript
sqlxlsx.update_cfg({
	sql: 'select somedata from sometable',
	oracledb_cfg: {
		user: 'username',
		pass: 'password',
		connectString: '192.168.1.254/YOGA' // Easy Connect syntax
	},
	oracledb_args: [],
	numRows: 100, // This is the number of rows to pull/stream at a time
	exceljs_wsName: 'worksheet name',
	exceljs_options: {
		filename: "./report.xlsx", // specifies the path to a file to write the XLSX workbook to
		useStyles: false,
		useSharedStrings: true,
		creator: "Name"
	}
});
````

## Change Individual Parameter(s)
````javascript
sqlxlsx.update_cfg({exceljs_wsName: 'new name'});
````

## Create and Define
##### You can create and define a report with one or more parameters all at once, like this:
```javascript
var sqlxlsx = new Sqlxlsx({
	sql: 'select somedata from sometable',
	oracledb_cfg: {
		user: 'username',
		pass: 'password',
		connectString: '192.168.1.254/YOGA' // Easy Connect syntax
	},
	oracledb_args: [],
	numRows: 100, // This is the number of rows to pull/stream at a time
	exceljs_wsName: 'worksheet name',
	exceljs_options: {
		filename: "./report.xlsx", // specifies the path to a file to write the XLSX workbook to
		useStyles: false,
		useSharedStrings: true,
		creator: "Name"
	}
});
```

## Run a Report
##### This runs and saves the report to a file.
```javascript
sqlxlsx.run(function(err) {
	console.log('report complete!');
});
```

## Example
```javascript
var Sqlxlsx = require('sqlxlsx'),
	sqlxlsx = new Sqlxlsx({
		sql: 'select somedata from sometable',
		oracledb_cfg: {
			user: 'username',
			pass: 'password',
			connectString: '192.168.1.254/YOGA' // Easy Connect syntax
		},
		oracledb_args: [],
		numRows: 100, // This is the number of rows to pull/stream at a time
		exceljs_wsName: 'worksheet name',
		exceljs_options: {
			filename: "report.xlsx", // specifies the path to a file to write the XLSX workbook to
			useStyles: false,
			useSharedStrings: true,
			creator: "Name"
		},
		afterConnect: function() {console.log('Oracle connection made.');},
		afterDisconnect: function() {console.log('Oracle connection released.');},
		afterEachFetch: function(row_count) {
			process.stdout.write("Streaming data (" + row_count + " rows) to file... \x1B[0G");
		},
		afterFetch: function(row_count) {
			process.stdout.clearLine();
			console.log('Downloaded ' + row_count + ' rows to an Excel file. ');
		},
		afterExecute: function() {
			console.log('Query executed. ');
		}
	});

sqlxlsx.run(function(err) {console.log('report complete!');});

```

## Parameter Definitions
- **oracledb_cfg:** [see node-oracledb getConnection(): connAttrs connAttrs](https://github.com/oracle/node-oracledb/blob/master/doc/api.md#parameters-1)
- **sql:** The SQL string that is executed. The SQL string may contain bind parameters.
- **oracledb_args:** [see node-oracledb execute(): Bind Parameters](https://github.com/oracle/node-oracledb/blob/master/doc/api.md#-4232-execute-bind-parameters)
- **numRows:** number of rows to stream from each call to the database driver
- **exceljs_wsName:** the name to give the Worksheet
- **exceljs_options:** [same as options defined within ExcelJS.stream.xlsx.WorkbookWriter namespace.](https://github.com/guyonroche/exceljs#streaming-xlsx)
- **afterConnect:** optional function to call after connecting to the database  ```function() {}```
- **afterDisconnect:** optional function to call after disconnecting from the database  ```function() {}```
- **afterEachFetch:** optional function to call after each fetch  ```function(row_count) {}```
- **afterFetch:** optional function to call after last fetch  ```function(row_count) {}```
- **afterExecute:** optional function to call once sql successfully executes  ```function(row_count) {}```
- **callback:** define a function to call once complete ```function(err) {}```
 
