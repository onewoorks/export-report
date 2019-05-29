var express = require('express')
var bodyParser = require('body-parser')

var Axios = require('axios')
var test = require('./components/test')
var NonCommercial = require('./components/noncommercial')
var VacantPremise = require('./components/vacant-premise')

var app = express()
app.use(bodyParser.json())
app.use(bodyParser.urlencoded({
  extended: true
}))

app.use(function (req, res, next) {
  res.header("Access-Control-Allow-Origin", "*");
  res.header("Access-Control-Allow-Headers", "Origin, X-Requested-With, Content-Type, Accept");
  next();
});

app.get('/', (request, response) => {
  //response.send('Hello World')
  response.send("<a href='output/Excel_2.xlsx'>Report Generate</a>")
})

app.get('/download', function (req, res) {
  var file = __dirname + '/output/Excel_2.xlsx';
  res.download(file); // Set disposition and send it.
});

app.get('/countdown', function (req, res) {
  res.writeHead(200, {
    'Content-Type': 'text/event-stream',
    'Cache-Control': 'no-cache',
    'Connection': 'keep-alive'
  })
  countdown(res, 1)
})

function countdown(res, count) {
  res.write("data: " + count + "\n\n")
  if (count)
    setTimeout(() => countdown(res, count + 1), 1000)
  else
    res.end()
}

app.get('/report/:module/:tab/:seq_start/:seq_end', (req, res, next) => {
  // console.log(req.params)
  // NonCommercial.createReport(result => {
  //   res.json(result)
  // })

  switch (req.params.module) {
    case 'vacant-premise':
      VacantPremise.createReport(result => {
        res.json(result);
      })
      break;
    case 'non-commercial':
      break;

  }
})

app.post('/generate-report', (request, response) => {
  test.foo()
  let data = {
    name: "report name",
    link: "http://senjabina.onewoorks-solutions.com/dashboard"
  }
  response.json(data)
})

var server = app.listen(8081, () => {
  var host = server.address().address
  var port = server.address().port
  console.log("Server is listening at http://%s:%s", host, port)
})