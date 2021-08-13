var express = require('express');
var router = express.Router();
// Formidble to handle file upload
var formidable = require('formidable');
// Import PythonShell module.
const {PythonShell} =require('python-shell');
const fs = require('fs')

cleanUpDir = function(dirPath, removeSelf) {
  try { var files = fs.readdirSync(dirPath); }
  catch(e) { return; }
  if (files.length > 0)
    for (var i = 0; i < files.length; i++) {
      var filePath = dirPath + '/' + files[i];
      if (fs.statSync(filePath).isFile())
        fs.unlinkSync(filePath);
      else
        cleanUpDir(filePath);
    }
  if (removeSelf)
    fs.rmdirSync(dirPath);
};

/* GET home page. */
router.get('/', function(req, res, next) {
  res.render('index', { title: 'TimeCard Parser' });

  // Cleanup Policies
  cleanUpDir(__dirname + '/uploads/')
  cleanUpDir(__dirname + '/downloads/')
});

router.post('/', function (req, res, next){
    var form = new formidable.IncomingForm();
    form.parse(req);
    // Set path to upload file
    form.on('fileBegin', function (name, file){
        file.path = __dirname + '/uploads/' + file.name
    });
    // Once its uploaded completely, call to python csv parser
    form.on('file', function (name, file){
      console.log('Uploaded ' + file.name);

      // options object to be passed 
      let options = {
          mode: 'text',
          pythonOptions: ['-u'], // get print results in real-time
            scriptPath: __dirname + '/../', // Location of timeCardParser.py is 1 directory back
          args: [file.name] // Pass name of the uploaded file as an argument
      };
        
      // Call timeCardParser with options
      PythonShell.run('timeCardParser.py', options, function (err, result){
            if (err) throw err;
            // result is an array consisting of messages collected 
            //during execution of script.
            const file = __dirname + '/downloads/' + result.toString()
            res.download(file); // Set disposition and send it.
      });

    });

});

module.exports = router;
