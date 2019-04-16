var express = require('express')
var bodyParser = require('body-parser')

var test = require('./components/test')

var app = express()
app.use(bodyParser.json())
app.use(bodyParser.urlencoded({
    extended: true
}))

app.get('/', (request, response) => {
    response.send('Hello World')
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

