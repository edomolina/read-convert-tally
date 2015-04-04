var fs = require('fs');

var patternList = {};

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

  input.on('end', function() {
    if (remaining.length > 0) {
      func(remaining);
    }
    for (var item in patternList) {
      console.log(item, patternList[item]);
    }    
  });
}

var input = fs.createReadStream('ver.txt');
readLines(input, function(line) {
  //console.log(line);
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

  console.log(id, patternList[id]);

  //var datatable[0]=line.substr(0, 12);
  var newline=fname.trim() + '|' + lname.trim() + '|' + add.trim() + '|' + city.trim() + '|' + state.trim() + '|' + zip.trim() + '|' + id.trim() + '\n';
//  console.log(fname);
//  console.log(lname);
//  console.log(add);
//  console.log(city);
//  console.log(state);
//  console.log(zip);
//  console.log(newline);

  fs.appendFile('newfile.txt', newline, function (err) {
  if (err) throw err;
//  console.log('Data appended.');
  });

});
console.log('New File Created.');
