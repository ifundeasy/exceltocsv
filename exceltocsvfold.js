var exts = ['xlsb', 'xlsx', 'xlsm', 'xls', 'xml', 'ods'];
var fs = require('fs');
var xlsx = require('xlsx');
var entities = require("entities");
var sys = require('./sys');

var log = function () {
	return console.log.apply(console.log, arguments)
};
var stringy = function (input) {
	return JSON.stringify(input)
};

var input = sys(process.argv[2]);
var force = process.argv[3];
var error = new Error('Error');
error.log = function (message) {
	error.message = message;
	log('Error :', error.message);
};

if (!input) {
	error.log('Invalid path directory/file!');
	process.exit();
}

var toCSV = function (opts, callback) {
	var res = {};
	var proto = opts.sys;
	var file = opts.file;
	var name = proto.poin + proto.name;
	var dir = name.substr(0, name.length - proto.ext.length - 1);
	var workbook = xlsx.readFile(file);

	try {
		fs.mkdirSync(dir);
	} catch (e) {}

	workbook.SheetNames.forEach(function(sheet) {
		var content = xlsx.utils.sheet_to_csv(workbook.Sheets[sheet]);
		if (content.length > 0) {
			var name = dir + '/' + entities.decodeXML(sheet)+ '.csv';
			res[name] = content;
		}
	});

	return callback(res);
};
toCSV.save = function (file) {
	var opts = {
		file : file,
		sys : sys(file)
	};

	log();
	log('Input :', file);

	toCSV(opts, function(res){
		for (var key in res) {
			var check = sys(key) || {};
			var process = true;

			if (check.is == 'file') process = force ? true : false;
			if (process) {
				log('Output:', key)
			} else {
				error.log('Already exist for "' + key + '"');
			}

			fs.writeFileSync(key, res[key].toString("utf-8"), 'utf-8')
		}
	});
};

if (input.is == 'file') {
	var poin = input.poin;
	var name = input.name;
	if (exts.indexOf(input.ext) > -1) {
		var file = poin + name;
		toCSV.save(file);
	}
} else {
	sys.readSync(input.poin, function (res) {
		var poin = res.poin;
		var name = res.name;
		if (exts.indexOf(res.ext) > -1) {
			var file = poin + name;
			toCSV.save(file);
		}
	});
}