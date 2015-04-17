var fs = require('fs');
var XLSX = require('XLSX');
var _ = require('underscore');
var watch = require ('watch');

var patternList = {};
var statecount = {};
var tabs = ["MKTID","STATE COUNT"];

//dont know why
var wscols = [
	{wch:6},
	{wch:7},
	{wch:10},
	{wch:20}
];

watch.createMonitor('/users/molinae/dev/read-convert-tally/', function (monitor) {
	monitor.files['/users/molinae/dev/read-convert-tally/.zshrc'] // Stat object for my zshrc.
	monitor.on("created", function (f, stat) {
		console.log('new file added');
		monitor.stop();

		var headerrow='fname|lname|add|city|state|zip|id' + '\n';

		fs.writeFile('pipefile.txt', headerrow, function (err) {
			if (err) throw err;
		});

		//main
		function readLines(input, callback) {
			var remaining = '';

			input.on('data', function(data) {
				remaining += data;
				var index = remaining.indexOf('\n');
				while (index > -1) {
					var line = remaining.substring(0, index);
					remaining = remaining.substring(index + 1);
					callback(line);
					index = remaining.indexOf('\n');
				}
			});

			//end of file
			input.on('end', function() {
				if (remaining.length > 0) {
					func(remaining);
				}
				for (var item in patternList) {
					console.log(item, patternList[item]);
				}

				//create excel file
				var index;


				function Workbook() {
					if(!(this instanceof Workbook)) return new Workbook();
					this.SheetNames = [];
					this.Sheets = {};
				}
				var wb = new Workbook();
				var sheetdata = [];

				for (index = 0; index < tabs.length; ++index) {

					var ws_name = tabs[index];
					switch (ws_name) {
						case 'MKTID':
							sheetdata = _.pairs(patternList);
							break;
							case 'STATE COUNT':
								sheetdata = _.pairs(statecount);
								break;
							}


							/* convert an array of arrays in JS to a CSF spreadsheet */
							function sheet_from_array_of_arrays(sheetdata, opts) {
								var ws = {};
								var range = {s: {c:10000000, r:10000000}, e: {c:0, r:0 }};
								for(var R = 0; R != sheetdata.length; ++R) {
									for(var C = 0; C != sheetdata[R].length; ++C) {
										if(range.s.r > R) range.s.r = R;
										if(range.s.c > C) range.s.c = C;
										if(range.e.r < R) range.e.r = R;
										if(range.e.c < C) range.e.c = C;
										var cell = {v: sheetdata[R][C] };

										if(cell.v == null) continue;
										var cell_ref = XLSX.utils.encode_cell({c:C,r:R});

										/* TEST: proper cell types and value handling */
										if(typeof cell.v === 'number') cell.t = 'n';
										else if(typeof cell.v === 'boolean') cell.t = 'b';
										else if(cell.v instanceof Date) {
											cell.t = 'n'; cell.z = XLSX.SSF._table[14];
											cell.v = datenum(cell.v);
										}
										else cell.t = 's';
										ws[cell_ref] = cell;
									}
								}

								/* TEST: proper range */
								if(range.s.c < 10000000) ws['!ref'] = XLSX.utils.encode_range(range);
								return ws;
							}
							var ws = sheet_from_array_of_arrays(sheetdata);

							wb.SheetNames.push(ws_name);
							wb.Sheets[ws_name] = ws;

							/* TEST: column widths */
							ws['!cols'] = wscols;

							/* write file */
							XLSX.writeFile(wb, 'ClientReport.xlsx');
							console.log('Excel File Created.');
						}

					});
				}

				var input = fs.createReadStream('ver.txt');
				readLines(input, function(line) {
					var fname=line.substr(0, 12);
					var lname=line.substr(13, 16);
					var add=line.substr(29, 24);
					var city=line.substr(53, 18);
					var state=line.substr(71, 12)
					var zip=line.substr(83, 10);
					var id=line.substr(93, 2);

					if (patternList[id] === undefined) {
						patternList[id] = 1;
					} else {
						patternList[id] += 1;
					}

					if (statecount[state] === undefined) {
						statecount[state] = 1;
					} else {
						statecount[state] += 1;
					}


					//  console.log(id, patternList[id]);

					var newline=fname.trim() + '|' + lname.trim() + '|' + add.trim() + '|' + city.trim() + '|' + state.trim() + '|' + zip.trim() + '|' + id.trim() + '\n';

					fs.appendFile('pipefile.txt', newline, function (err) {
						if (err) throw err;
					});

				});
				console.log('Fix Width File Converted to Pipe File.');

			})

			//monitor.stop();
			//	monitor.on("changed", function (f, curr, prev) {
			//		// Handle file changes
			//	})
			//	monitor.on("removed", function (f, stat) {
			//		// Handle removed files
			//	})
			//	monitor.stop(); // Stop watching
			//})

		})
