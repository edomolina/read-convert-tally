var fs = require('fs');
var XLSX = require('XLSX');
var _ = require('underscore');
var watch = require ('watch');
var nodemailer = require('nodemailer');

//direct transport
//var transporter = nodemailer.createTransport();

var transporter = nodemailer.createTransport({
	service: 'gmail',
	auth: {
		user: 'edomolina66@gmail.com',
		pass: 'josieedo69'
	}
});


//were here
var combosArray = [];
var numberofgridrows = 1747;

var watchFolder='/users/molinae/dev/read-convert-tally/';
var watchFiles='/users/molinae/dev/read-convert-tally/.zshrc';
var writefilePipe='/users/molinae/verizon_com_to_presort.txt';

var combosWorkbook = XLSX.readFile('combos.xlsx');
var sheet_name_list = combosWorkbook.SheetNames;
sheet_name_list.forEach(function(y) {
	var worksheet = combosWorkbook.Sheets[y];
	for (var i = 2; i < numberofgridrows; i+=1) {
		combosArray.push({combo: worksheet['A' + i].v , lasercode: worksheet['B' + i].v})
	}
	console.log('Grid Combos Loaded');
});


watch.createMonitor(watchFolder, function (monitor) {
	monitor.files[watchFiles] // Stat object for my zshrc.

	monitor.on("created", function (newcomfile, stat) {
		console.log('new file added', newcomfile);

		var mktgidCount = {};
		var mktgidTotal = 0;
		var billstateCount = {};
		var billstateTotal = 0;
		var servicestateCount = {};
		var servicestateTotal = 0;
		var lasercomboCount = {};
		var lasercomboTotal = 0;
		var unassignedCount = {};
		var unassignedTotal = 0;
		var tabs = ["MKTG ID","BILL STATE","SERVICE STATE","LETTERCOPY","NOT ASSIGNED"];
		var bufferString;
		var cellLocation;


		//monitor.stop();

		var headerrow='custid|bucucan|custacctnum|legacysysid|highlvlqual|addressid|filler1|billtelnumber|filler2|jobcontrolnum|cycle|leadtype|filler3|activitydetailid|regionstatecode|marketingtype|filler4|cell14|filler5|accountestabdate|servicestatejuris|vendorname|filler6|customername|billstreetadd|billsubadd|billcity|billstate|billcrrt|billzip|billzip4|servicestreetadd|servicesubadd|servicecity|servicestate|servicecrrt|servicezip|servicezip4|returnmailkeycode|dpbc|chkdig|downstatenylata|couponcode|couponcodemsgexp|couponcodeactualexp|inhomedate|plandropdate|planshipdate|leadoffer|testdescrip|testplan|testcelldescrip|agencyjobnum|dmacode|creativename|mktgid|cellnumber|campaignid|messagecopy|lasercopycode|letterheadcode|legaldisclaimercode|channellineupcode|outerenvcode|optionalmailcode|leadpackprice1|leadpackprice2|leadpackupspeed|leadpackdownspeed|tollfreenumber1|url1|seedrecidentifier|collateraltext|segmentname|controlcode|ethnichighlevel|tollfreenumber2|url2|hoursofop1|hoursofop2|promolength|promoexpdate|buckslipcode|buckslip|datapackagecodea|cicode|campaigntrackcode|misccomponent|spanishlasercopy|spanishletterhead|mailclass|agencyname|alldigitalchan|alldigitalchan2|alldigitalchan3|hdchan|hdchan2|hdchan3|pkg2leadofffer|pkg2upspeed|pkg2downspeed|pkg3leadoffer|pkg3price|pkg3upspeed|pkg3downspeed|inlngchannel|testcontrol|cellstatuscheck|voiceexpdate|internetexpdate|tvexpdate|dmaname|faqcodeaa|ecertcode|hsicustomer|voiceflag|fiostvcust|ntwkevoleligble|entertainsegmen|reposetype|cableprovider|lasercombo' + '\n';


		fs.writeFile(writefilePipe, headerrow, function (err) {
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
						case 'MKTG ID':

							sheetdata = _.pairs(mktgidCount);
							sheetdata.sort();

							for (var i=0; i < sheetdata.length; i+=1){
								mktgidTotal = mktgidTotal + sheetdata[i][1];
								//console.log(mktgidTotal);
							}

							sheetdata.unshift(['MKTG ID', 'COUNT']);
							sheetdata.push([' ', ' ']);
							sheetdata.push([' ', mktgidTotal]);

							var wscols = [
								{wch:15},
								{wch:10}
							];
							break;
						case 'BILL STATE':
							sheetdata = _.pairs(billstateCount);
							sheetdata.sort();

							for (var i=0; i < sheetdata.length; i+=1){
								billstateTotal = billstateTotal + sheetdata[i][1];
							}

							sheetdata.unshift(['BILL STATE', 'COUNT']);
							sheetdata.push([' ', ' ']);
							sheetdata.push([' ', billstateTotal]);

							var wscols = [
								{wch:15},
								{wch:10}
							];
							break;
						case 'SERVICE STATE':
							sheetdata = _.pairs(servicestateCount);
							sheetdata.sort();

							for (var i=0; i < sheetdata.length; i+=1){
								servicestateTotal = servicestateTotal + sheetdata[i][1];
							}

							sheetdata.unshift(['SERVICE STATE', 'COUNT']);
							sheetdata.push([' ', ' ']);
							sheetdata.push([' ', servicestateTotal]);

							var wscols = [
								{wch:15},
								{wch:10}
							];
							break;

						case 'LETTERCOPY':
							sheetdata = _.pairs(lasercomboCount);
							sheetdata.sort();

							for (var i=0; i < sheetdata.length; i+=1){
								lasercomboTotal = lasercomboTotal + sheetdata[i][1];
							}

							sheetdata.unshift(['LASER COPY CODE', 'COUNT']);
							sheetdata.push([' ', ' ']);
							sheetdata.push([' ', lasercomboTotal]);

							var wscols = [
								{wch:30},
								{wch:10}
							];
							break;

						case 'NOT ASSIGNED':
							sheetdata = _.pairs(unassignedCount);
							sheetdata.sort();

							for (var i=0; i < sheetdata.length; i+=1){
								unassignedTotal = unassignedTotal + sheetdata[i][1];
							}

							sheetdata.unshift(['UNASSIGNED LASER COMBO', 'COUNT']);
							sheetdata.push([' ', ' ']);
							sheetdata.push([' ', unassignedTotal]);

							var wscols = [
								{wch:30},
								{wch:10}
							];
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

							var reportfileXlsx='JBXXXXXX VERIZON COM CYCLE ' + newcomfile.substr(47, 2) + ' QC REPORT.xlsx';

							XLSX.writeFile(wb, reportfileXlsx);

							console.log(ws_name, 'sheet added to Excel File.');
						}

				console.log('Process Complete');

			var message = {

				from: 'Edo Molina <edomolina66@gmail.com>',

				// Comma separated list of recipients
				to: '"Edo Molina" <emolina@earthcolor.com>',

				bcc: '<edo.molina@mailersplus.com',

				subject: 'Verizon COM QC Report ✔', //

				//headers: {
				//	'X-Laziness-level': 1000
				//},

				// plaintext body
				text: 'Above referenced Report attached.',

				// HTML body
				//html: '<p><b>Hello</b> to myself <img src="cid:note@example.com"/></p>' +
				//'<p>Here\'s a nyan cat for you as an embedded attachment:<br/><img src="cid:nyan@example.com"/></p>',

				// An array of attachments
				attachments: [

					// String attachment
					{
						filename: reportfileXlsx, //'notes.txt',
						content: 'Some notes about this e-mail'
						//contentType: 'text/plain' // optional, would be detected from the filename
					}//,

				]
			};

			console.log('Sending Mail');
			transporter.sendMail(message, function(error, info) {
				if (error) {
					console.log('Error occurred');
					console.log(error.message);
					return;
				}
				console.log('Message sent successfully!');
				console.log('Server responded with "%s"', info.response);
			});


			});
				}

				var input = fs.createReadStream(newcomfile);
					readLines(input, function (line) {
						if (line.length>500) {

							var custid = line.substr(0, 12);
							var bucucan = line.substr(12, 10);
							var custacctnum = line.substr(22, 13);
							var legacysysid = line.substr(35, 4);
							var highlvlqual = line.substr(39, 7);
							var addressid = line.substr(46, 10);
							var filler1 = line.substr(56, 1);
							var billtelnumber = line.substr(57, 10);
							var filler2 = line.substr(67, 8);
							var jobcontrolnum = line.substr(75, 5);
							var cycle = line.substr(80, 3);
							var leadtype = line.substr(83, 1);
							var filler3 = line.substr(84, 6);
							var activitydetailid = line.substr(90, 12);
							var regionstatecode = line.substr(102, 4);
							var marketingtype = line.substr(106, 1);
							var filler4 = line.substr(107, 12);
							var cell14 = line.substr(119, 1);
							var filler5 = line.substr(120, 1);
							var accountestabdate = line.substr(121, 8);
							var servicestatejuris = line.substr(129, 2);
							var vendorname = line.substr(131, 5);
							var filler6 = line.substr(136, 1);
							var customername = line.substr(137, 35);
							var billstreetadd = line.substr(172, 40);
							var billsubadd = line.substr(212, 25);
							var billcity = line.substr(237, 15);
							var billstate = line.substr(252, 2);
							var billcrrt = line.substr(254, 4);
							var billzip = line.substr(258, 5);
							var billzip4 = line.substr(263, 4);
							var servicestreetadd = line.substr(267, 40);
							var servicesubadd = line.substr(307, 25);
							var servicecity = line.substr(332, 15);
							var servicestate = line.substr(347, 2);
							var servicecrrt = line.substr(349, 4);
							var servicezip = line.substr(353, 5);
							var servicezip4 = line.substr(358, 4);
							var returnmailkeycode = line.substr(362, 26);
							var dpbc = line.substr(388, 2);
							var chkdig = line.substr(390, 1);
							var downstatenylata = line.substr(391, 1);
							var couponcode = line.substr(392, 14);
							var couponcodemsgexp = line.substr(406, 8);
							var couponcodeactualexp = line.substr(414, 8);
							var inhomedate = line.substr(422, 10);
							var plandropdate = line.substr(432, 10);
							var planshipdate = line.substr(442, 10);
							var leadoffer = line.substr(452, 40);
							var testdescrip = line.substr(492, 40);
							var testplan = line.substr(532, 35);
							var testcelldescrip = line.substr(567, 40);
							var agencyjobnum = line.substr(607, 8);
							var dmacode = line.substr(615, 3);
							var creativename = line.substr(618, 40);
							var mktgid = line.substr(658, 10);
							var cellnumber = line.substr(668, 10);
							var campaignid = line.substr(678, 15);
							var messagecopy = line.substr(693, 1);
							var lasercopycode = line.substr(694, 25);
							var letterheadcode = line.substr(719, 25);
							var legaldisclaimercode = line.substr(744, 25);
							var channellineupcode = line.substr(769, 25);
							var outerenvcode = line.substr(794, 25);
							var optionalmailcode = line.substr(819, 25);
							var leadpackprice1 = line.substr(844, 10);
							var leadpackprice2 = line.substr(854, 10);
							var leadpackupspeed = line.substr(864, 5);
							var leadpackdownspeed = line.substr(869, 5);
							var tollfreenumber1 = line.substr(874, 14);
							var url1 = line.substr(888, 40);
							var seedrecidentifier = line.substr(928, 1);
							var collateraltext = line.substr(929, 12);
							var segmentname = line.substr(941, 6);
							var controlcode = line.substr(947, 10);
							var ethnichighlevel = line.substr(957, 3);
							var tollfreenumber2 = line.substr(960, 14);
							var url2 = line.substr(974, 40);
							var hoursofop1 = line.substr(1014, 40);
							var hoursofop2 = line.substr(1054, 40);
							var promolength = line.substr(1094, 20);
							var promoexpdate = line.substr(1114, 10);
							var buckslipcode = line.substr(1124, 30);
							var buckslip = line.substr(1154, 150);
							var datapackagecodea = line.substr(1304, 30);
							var cicode = line.substr(1334, 30);
							var campaigntrackcode = line.substr(1364, 20);
							var misccomponent = line.substr(1384, 50);
							var spanishlasercopy = line.substr(1434, 30);
							var spanishletterhead = line.substr(1464, 100);
							var mailclass = line.substr(1564, 250);
							var agencyname = line.substr(1814, 50);
							var alldigitalchan = line.substr(1864, 40);
							var alldigitalchan2 = line.substr(1904, 4);
							var alldigitalchan3 = line.substr(1908, 4);
							var hdchan = line.substr(1912, 40);
							var hdchan2 = line.substr(1952, 4);
							var hdchan3 = line.substr(1956, 4);
							var pkg2leadoffer = line.substr(1960, 100);
							var pkg2upspeed = line.substr(2060, 20);
							var pkg2downspeed = line.substr(2080, 20);
							var pkg3leadoffer = line.substr(2100, 100);
							var pkg3price = line.substr(2200, 20);
							var pkg3upspeed = line.substr(2220, 20);
							var pkg3downspeed = line.substr(2240, 20);
							var inlngchannel = line.substr(2260, 4);
							var testcontrol = line.substr(2264, 7);
							var cellstatuscheck = line.substr(2271, 50);
							var voiceexpdate = line.substr(2321, 10);
							var internetexpdate = line.substr(2331, 10);
							var tvexpdate = line.substr(2341, 10);
							var dmaname = line.substr(2351, 35);
							var faqcodeaa = line.substr(2386, 15);
							var ecertcode = line.substr(2401, 15);
							var hsicustomer = line.substr(2416, 1);
							var voiceflag = line.substr(2417, 1);
							var fiostvcust = line.substr(2418, 1);
							var ntwkevoleligble = line.substr(2419, 1);
							var entertainsegmen = line.substr(2420, 2);
							var reposetype = line.substr(2422, 25);
							var cableprovider = line.substr(2447, 2);
							//var endpipe = line.substr(2449, 1);

							var lasercombo = mktgid.trim() + '-' + leadtype.trim() + '-' + hsicustomer.trim() + '-' + voiceflag.trim() + '-' + fiostvcust.trim() + '-' + ntwkevoleligble.trim() + '-' + entertainsegmen.trim() + '-' + reposetype.trim();


							if (mktgidCount[mktgid] === undefined) {
								mktgidCount[mktgid] = 1;
							} else {
								mktgidCount[mktgid] += 1;
							}

							if (billstateCount[billstate] === undefined) {
								billstateCount[billstate] = 1;
							} else {
								billstateCount[billstate] += 1;
							}

							if (servicestateCount[servicestate] === undefined) {
								servicestateCount[servicestate] = 1;
							} else {
								servicestateCount[servicestate] += 1;
							}

							var gridCombo;
							for (var i = 0; i < combosArray.length; i += 1) {
								if (combosArray[i].combo === lasercombo) {
									gridCombo = combosArray[i].lasercode;
									break;
								}
							}

							var undefinedCombo;
							if (gridCombo === undefined) {
								undefinedCombo = lasercombo;
								if (unassignedCount[undefinedCombo] === undefined) {
									unassignedCount[undefinedCombo] = 1;
								} else {
									unassignedCount[undefinedCombo] += 1;
								}

							} else {
								lasercopycode=gridCombo;
							}
							// code to handle undefined gridCombo here

							if (lasercomboCount[gridCombo] === undefined) {
								lasercomboCount[gridCombo] = 1;
							} else {
								lasercomboCount[gridCombo] += 1;
							}

							var newline = custid.trim() + '|' + bucucan.trim() + '|' + custacctnum.trim() + '|' + legacysysid.trim() + '|' + highlvlqual.trim() + '|' + addressid.trim() + '|' + filler1.trim() + '|' + billtelnumber.trim() + '|' + filler2.trim() + '|' + jobcontrolnum.trim() + '|' + cycle.trim() + '|' + leadtype.trim() + '|' + filler3.trim() + '|' + activitydetailid.trim() + '|' + regionstatecode.trim() + '|' + marketingtype.trim() + '|' + filler4.trim() + '|' + cell14.trim() + '|' + filler5.trim() + '|' + accountestabdate.trim() + '|' + servicestatejuris.trim() + '|' + vendorname.trim() + '|' + filler6.trim() + '|' + customername.trim() + '|' + billstreetadd.trim() + '|' + billsubadd.trim() + '|' + billcity.trim() + '|' + billstate.trim() + '|' + billcrrt.trim() + '|' + billzip.trim() + '|' + billzip4.trim() + '|' + servicestreetadd.trim() + '|' + servicesubadd.trim() + '|' + servicecity.trim() + '|' + servicestate.trim() + '|' + servicecrrt.trim() + '|' + servicezip.trim() + '|' + servicezip4.trim() + '|' + returnmailkeycode.trim() + '|' + dpbc.trim() + '|' + chkdig.trim() + '|' + downstatenylata.trim() + '|' + couponcode.trim() + '|' + couponcodemsgexp.trim() + '|' + couponcodeactualexp.trim() + '|' + inhomedate.trim() + '|' + plandropdate.trim() + '|' + planshipdate.trim() + '|' + leadoffer.trim() + '|' + testdescrip.trim() + '|' + testplan.trim() + '|' + testcelldescrip.trim() + '|' + agencyjobnum.trim() + '|' + dmacode.trim() + '|' + creativename.trim() + '|' + mktgid.trim() + '|' + cellnumber.trim() + '|' + campaignid.trim() + '|' + messagecopy.trim() + '|' + lasercopycode.trim() + '|' + letterheadcode.trim() + '|' + legaldisclaimercode.trim() + '|' + channellineupcode.trim() + '|' + outerenvcode.trim() + '|' + optionalmailcode.trim() + '|' + leadpackprice1.trim() + '|' + leadpackprice2.trim() + '|' + leadpackupspeed.trim() + '|' + leadpackdownspeed.trim() + '|' + tollfreenumber1.trim() + '|' + url1.trim() + '|' + seedrecidentifier.trim() + '|' + collateraltext.trim() + '|' + segmentname.trim() + '|' + controlcode.trim() + '|' + ethnichighlevel.trim() + '|' + tollfreenumber2.trim() + '|' + url2.trim() + '|' + hoursofop1.trim() + '|' + hoursofop2.trim() + '|' + promolength.trim() + '|' + promoexpdate.trim() + '|' + buckslipcode.trim() + '|' + buckslip.trim() + '|' + datapackagecodea.trim() + '|' + cicode.trim() + '|' + campaigntrackcode.trim() + '|' + misccomponent.trim() + '|' + spanishlasercopy.trim() + '|' + spanishletterhead.trim() + '|' + mailclass.trim() + '|' + agencyname.trim() + '|' + alldigitalchan.trim() + '|' + alldigitalchan2.trim() + '|' + alldigitalchan3.trim() + '|' + hdchan.trim() + '|' + hdchan2.trim() + '|' + hdchan3.trim() + '|' + pkg2leadoffer.trim() + '|' + pkg2upspeed.trim() + '|' + pkg2downspeed.trim() + '|' + pkg3leadoffer.trim() + '|' + pkg3price.trim() + '|' + pkg3upspeed.trim() + '|' + pkg3downspeed.trim() + '|' + inlngchannel.trim() + '|' + testcontrol.trim() + '|' + cellstatuscheck.trim() + '|' + voiceexpdate.trim() + '|' + internetexpdate.trim() + '|' + tvexpdate.trim() + '|' + dmaname.trim() + '|' + faqcodeaa.trim() + '|' + ecertcode.trim() + '|' + hsicustomer.trim() + '|' + voiceflag.trim() + '|' + fiostvcust.trim() + '|' + ntwkevoleligble.trim() + '|' + entertainsegmen.trim() + '|' + reposetype.trim() + '|' + cableprovider.trim() + '|' + lasercombo.trim() + '\n';

							fs.appendFile(writefilePipe, newline, function (err) {
								if (err) throw err;
							});
						}
					});
					console.log('Fix Width File Converted to Pipe File.');
			});

			//monitor.stop();
			//	monitor.on("changed", function (f, curr, prev) {
			//		// Handle file changes
			//	})
			//	monitor.on("removed", function (f, stat) {
			//		// Handle removed files
			//	})
			//	monitor.stop(); // Stop watching
			//})

		});
