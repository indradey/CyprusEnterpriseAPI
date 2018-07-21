// call the packages we need
var express = require('express');
var bodyParser = require('body-parser');
var app = express();
var cors = require('cors');
var morgan = require('morgan');
var ExcelHelper = require('./ExcelUtility/Excelhelper');
var _ = require('lodash');
// configure body parser
app.use(bodyParser.urlencoded({
	extended: true
}));
app.use(bodyParser.json());
app.use(morgan('dev')); // log requests to the console
var reqschema = require('./RequestType/Request.json');
// create our router
var router = express.Router();

// middleware to use for all requests
router.use(function (req, res, next) {
	// do logging
	console.log('Something is happening.');
	next();
});

router.get('/', function (req, res) {
	res.json({
		message: 'hooray! welcome to our api!'
	});
});

router.route('/userDetails')

	// create a bear (accessed at POST http://localhost:3600/userDetails)
	.post(async function (req, res) {
		if (req.body !== null || req.body !== '') {
			await ExcelHelper.writeToFile(req.body);
			res.json('Done');
		}
	})

	// get all the bears (accessed at GET http://localhost:3600/api/userDetails)
	.get(async function (req, res) {
		var data = await ExcelHelper.readFromFile();
		res.json(data);
	});

var port = process.env.PORT || 3600; // set our port

// REGISTER OUR ROUTES -------------------------------
app.use('/api', router);

// START THE SERVER
// =============================================================================
app.listen(port);
console.log('Magic happens on port ' + port);