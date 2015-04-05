//dependencies
var fs = require('fs');

var xlsx = require('xlsx')

//variables
var patternList = {};

//write header
var headerrow='fname|lname|add|city|state|zip|id' + '\n';

fs.writeFile('newfile.txt', headerrow, function (err) {
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

// write an XLSX file
   var xlsxWriter = new SimpleExcel.Writer.XLSX();
   var xlsxSheet = new SimpleExcel.Sheet();
   var Cell = SimpleExcel.Cell;
   xlsxSheet.setRecord([
       [new Cell('ID', 'TEXT'), new Cell('Nama', 'TEXT'), new Cell('Kode Wilayah', 'TEXT')],
       [new Cell(1, 'NUMBER'), new Cell('Kab. Bogor', 'TEXT'), new Cell(1, 'NUMBER')],
       [new Cell(2, 'NUMBER'), new Cell('Kab. Cianjur', 'TEXT'), new Cell(1, 'NUMBER')],
       [new Cell(3, 'NUMBER'), new Cell('Kab. Sukabumi', 'TEXT'), new Cell(1, 'NUMBER')],
       [new Cell(4, 'NUMBER'), new Cell('Kab. Tasikmalaya', 'TEXT'), new Cell(2, 'NUMBER')]
   ]);
   xlsxWriter.insertSheet(xlsxSheet);
   // export when button clicked
   document.getElementById('fileExport').addEventListener('click', function () {
       xlsxWriter.saveFile(); // pop! ("Save As" dialog appears)
   });


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

//id iteration object
  if (patternList[id] === undefined) {
    patternList[id] = 1;
  } else {
    patternList[id] += 1;
  }

  //console.log(id, patternList[id]);

  var newline=fname.trim() + '|' + lname.trim() + '|' + add.trim() + '|' + city.trim() + '|' + state.trim() + '|' + zip.trim() + '|' + id.trim() + '\n';

  fs.appendFile('newfile.txt', newline, function (err) {
  if (err) throw err;
  });

});
console.log('New File Created.');
