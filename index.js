"use strict";

var Excel		= require('exceljs'),
	oracledb	= require('oracledb');

function Sqlxlsx(args) {
	this.update_cfg(args);
}

Sqlxlsx.prototype = {
	config:	{
		oracledb_cfg: {
			user:			undefined,
			password:		undefined,
			connectString:	undefined
		},
		oracledb_conn:		undefined,
		oracledb_args:		undefined,
		numRows:			100,
		sql:				undefined,
		exceljs_wsName:		'',
		exceljs_options:	undefined,
		afterConnect:		undefined,
		afterDisconnect:	undefined,
		afterEachFetch:		undefined,
		afterFetch:			undefined,
		afterExecute:		undefined,
		callback:			undefined
	},
	update_cfg: function(args) {
		var obj = this;
		Object.keys(args).forEach(function(key) {
			if(args.hasOwnProperty(key)) {
				obj.config[key] = args[key];
			}
		});
	},
	connect: function() {
		var obj = this;
		oracledb.getConnection({
			user          : this.config.oracledb_cfg.user,
			password      : this.config.oracledb_cfg.password,
			connectString : this.config.oracledb_cfg.connectString
		}, function(err, connection) {
			if (err) {
				console.error(err.message);
				return;
			}
			obj.connected(connection);
		});
	},
	release: function() {
		var obj = this;
		this.config.oracledb_conn.release(
			function(err) {
				if(typeof obj.config.afterDisconnect === "function") {
					obj.config.afterDisconnect();
				}
				if (err) {
					console.error(err.message);
				}
			}
		);
	},
	connected: function(connection) {
		if(typeof this.config.afterConnect === "function") {
			this.config.afterConnect();
		}
		this.config.oracledb_conn = connection;
		this.convert();
	},
	run: function(callback) {
		this.config.callback = callback;
		this.connect();
	},
	convert: function() {
		var obj		= this,
			cfg		= this.config,
			doAEF	= (typeof cfg.afterEachFetch === "function") ? true : false,
			options	= cfg.exceljs_options || {
				filename: "report.xlsx",
				useStyles: false,
				useSharedStrings: true
			};

		cfg.oracledb_conn.execute(
			cfg.sql,
			cfg.oracledb_args,
			{ resultSet: true },
			function(err, result) {
				var workbook	= new Excel.stream.xlsx.WorkbookWriter(options),
					worksheet	= workbook.addWorksheet(cfg.exceljs_wsName),
					header		= [],
					row_count	= 0;

				function fetch() {
					result.resultSet.getRows(cfg.numRows, function(err, rows) {
						if (err) {
							obj.release();
							cfg.callback(err);
							return;
						}

						row_count += rows.length;

						rows.forEach(function(element) {
							worksheet.addRow(element).commit();
						});

						if(doAEF) {cfg.afterEachFetch(row_count);}

						if (rows.length === cfg.numRows) {
							fetch();
						} else {
							if(typeof cfg.afterFetch === "function") {cfg.afterFetch(row_count);}
							result.resultSet.close(function(err) {
								worksheet.commit();
								workbook.commit();
								obj.release();
								if (err) {
									cfg.callback(err);
									return;
								}
								cfg.callback();
							});
						}
					});
				}

				if (err) {
					obj.release();
					cfg.callback(err.message);
					return;
				}
				if(typeof cfg.afterExecute === "function") {cfg.afterExecute();}
				result.metaData.forEach(function(element) { // adds Column Header Row
					header.push(element.name);
				});
				worksheet.addRow(header).commit();
				fetch();
			}
		);
	}
};

module.exports = Sqlxlsx;
