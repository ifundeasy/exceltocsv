var fs = require('fs');
var path = require('path');
var sys = function (input) {
	var res = undefined;
	var is = sys.is(input);
	if (is) {
		var obj = path.parse(input);
		var base = path.normalize(obj.dir);
		res = {
			base : base == '/' ? base : (base + '/'),
			ext : is == 'dir' ? '' : obj.ext.substr(1, obj.ext.length),
			is : is,
			name : obj.base,
		};
		res.poin = is == 'dir' ? res.base + res.name + '/' : res.base;
	}

	return res;
};
sys.is = function (input) {
	var res = undefined;
	try {
		res = fs.lstatSync(input).isDirectory() ? 'dir' : 'file';
	} catch (e) {}

	return res;
};
sys.readSync = function (input, exec) {
	var check = sys(input);
	if (check.is == 'dir') {
		var contents = fs.readdirSync(check.poin)
		contents.forEach(function(content, i) {
			sys.readSync(check.poin + content, exec)
		})
	} else {
		if (exec && typeof exec == 'function') exec(check)
	}
};
sys.read = function (input, exec) {
	var check = sys(input);
	if (check.is == 'dir') {
		fs.readdir(check.poin, function(err, contents) {
			contents.forEach(function(content, i) {
				sys.read(check.poin + content, exec)
			})
		})
	} else {
		if (exec && typeof exec == 'function') exec(check)
	}
};

module.exports = sys;