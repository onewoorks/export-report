
var Axios = require('axios')
var xl = require('excel4node')
const fs = require('fs')

module.exports = {
  createReport: async (callback) => {
    var send = {
      status: "processing data to generate report"
    }
    callback(send)
    await Axios.get('http://localhost/senjabina/api/report/generate-excel?module=vacant-premise&tabs=db%20kuala%20lumpur')
      .then(response => {
        let resp = response.data.response
        createBook(resp)
        console.log('im done reporting...')
        //    callback(resp)
      }).catch(err => {
        console.log(err)
      })
  }
}

createBook = (data) => {
  var wb = new xl.Workbook()
  var styleHeader = wb.createStyle({
    font: {
      size: 9,
      bold: true
    },
    alignment: {
      horizontal: 'center',
      vertical: 'center',
      wordWrap: true
    }
  });
  var style = wb.createStyle({
    font: {
      size: 9,
    },
    alignment: {
      vertical: 'top',
      wrapText: true
    }
  });
  var stylecenter = wb.createStyle({
    font: {
      size: 9,
    },
    alignment: {
      vertical: 'top',
      horizontal: 'center'
    }
  });

  var ws = wb.addWorksheet(data.report_tabs)
  ws.row(1).setHeight(14)
  ws.row(2).setHeight(14)
  ws.row(3).setHeight(14)

  ws.cell(1,1,3,1,true)
  ws.cell(1,2,3,2,true)
  ws.cell(1,3,3,3,true)
  ws.cell(1,4,3,4,true)
  ws.cell(1,5,3,5,true)
  ws.cell(1,6,3,6,true)
  ws.cell(1,7,1,9,true).string('COLUMN A').style(styleHeader)
  ws.cell(2,7,2,9,true).string('OCCUPANCY STATUS').style(styleHeader)
  ws.cell(1,10,1,12,true).string('COLUMN B').style(styleHeader)
  ws.cell(2,10,2,12,true).string('ACTUAL CLASSIFICATION').style(styleHeader)
  ws.cell(1,13,2,14,true).string('DOES THIS PREMISE HAVE WATER METER CONNECTED?').style(styleHeader)
  ws.cell(1,15,3,15,true).string('METER NUMBER').style(styleHeader)
  ws.cell(1,16,3,16,true).string('REMARKS OR AMMENDMENT TO PROPERTY ').style(styleHeader)
  ws.cell(1,17,3,17,true).string('DATE VISITED').style(styleHeader)
  ws.cell(1,18,3,18,true).string('NAME OF STAFF VISITED').style(styleHeader)
  ws.column(1).setWidth(5)
  ws.column(2).setWidth(9)
  ws.column(3).setWidth(9)
  ws.column(4).setWidth(30)
  ws.column(5).setWidth(24)
  ws.column(6).setWidth(16)
  ws.column(7).setWidth(10)
  ws.column(8).setWidth(10)
  ws.column(9).setWidth(10)
  ws.column(10).setWidth(10)
  ws.column(11).setWidth(10)
  ws.column(12).setWidth(10)
  ws.column(13).setWidth(10)
  ws.column(14).setWidth(10)
  ws.column(15).setWidth(17)
  ws.column(16).setWidth(16)
  ws.column(17).setWidth(10)
  ws.column(18).setWidth(28)

  ws.cell(1, 1).string('ID').style(styleHeader)
  ws.cell(1, 2).string('SEQ ID').style(styleHeader)
  ws.cell(1, 3).string('SEWACC').style(styleHeader)
  ws.cell(1, 4).string('OWNER NAME').style(styleHeader)
  ws.cell(1, 5).string('PROPERTY ADDRESS').style(styleHeader)
  ws.cell(1, 6).string('CURRENT CLASS').style(styleHeader)
  ws.cell(3, 7).string('OCCUPIED').style(styleHeader)
  ws.cell(3, 8).string('VACANT').style(styleHeader)
  ws.cell(3, 9).string('OTHER').style(styleHeader)
  ws.cell(3, 10).string('COMMERCIAL').style(styleHeader)
  ws.cell(3, 11).string('DOMESTIC').style(styleHeader)
  ws.cell(3, 12).string('OTHER').style(styleHeader)
  ws.cell(3, 13).string('YES').style(styleHeader)
  ws.cell(3, 14).string('N0').style(styleHeader)
  ws.cell(3, 19).string('PHOTOGRAPHS').style(styleHeader)

  data.results.forEach((value, k) => {
    var key = k +4
    ws.row(key).setHeight(145)

    ws.cell(key, 1).string(value.id).style(style)
    ws.cell(key, 2).string(value.seq_id).style(style)
    ws.cell(key, 3).string(value.sewacc).style(style)
    ws.cell(key, 4).string(value.owner_name).style(style)
    ws.cell(key, 5).string(value.property_address).style(style)
    ws.cell(key, 6).string(value.current_class).style(style)

    var task_perform = (typeof value.upload_data.task_perform !== undefined) ? value.upload_data.task_perform : false
    ws.cell(key, 7).string((task_perform && task_perform.vacant_status.toLowerCase() == 'occupied') ? '/' : '').style(stylecenter)
    ws.cell(key, 8).string((task_perform && task_perform.vacant_status.toLowerCase() == 'vacant') ? '/' : '').style(stylecenter)
    ws.cell(key, 9).string('').style(style)
    ws.cell(key, 10).string((task_perform && task_perform.actual_classification.toLowerCase() == 'commercial') ? 'COMMERCIAL' : '').style(stylecenter)
    ws.cell(key, 11).string((task_perform && task_perform.actual_classification.toLowerCase() == 'domestic') ? 'DOMESTIC' : '').style(stylecenter)
    ws.cell(key, 12).string((task_perform && task_perform.actual_classification.toLowerCase() != 'domestic') ? '' : '').style(style)
    ws.cell(key, 13).string((task_perform && task_perform.meter_connected.toLowerCase() == 'yes') ? 'YES' : '').style(stylecenter)
    ws.cell(key, 14).string((task_perform && task_perform.meter_connected.toLowerCase() == 'no') ? 'NO' : '').style(stylecenter)
    ws.cell(key, 15).string((task_perform && task_perform.meter_number) ? task_perform.meter_number.toUpperCase() : '').style(stylecenter)
    ws.cell(key, 16).string((task_perform && task_perform.remarks) ? task_perform.remarks : '').style(style)
    ws.cell(key, 17).string((task_perform && task_perform.perform_date) ? task_perform.perform_date : value.complete_time).style(style)
    ws.cell(key, 18).string(value.name).style(stylecenter)

    var images = value.upload_data.photos
    let column_no = 18
    Object.keys(images).forEach((value, k) => {
      try {
        if (fs.existsSync('/Users/iwang/Sites/senjabina/upload/' + images[value])) {
          column_no = column_no + 1
          ws.column(column_no).setWidth(24)
          ws.addImage({
            path: '/Users/iwang/Sites/senjabina/upload/' + images[value],
            type: 'picture',
            position: {
              type: 'twoCellAnchor',
              from: {
                col: column_no,
                colOff: "1mm",
                row: key,
                rowOff: "1mm"
              },
              to: {
                col: column_no,
                colOff: "50mm",
                row: key,
                rowOff: "50mm"
              }
            }
          })
        }
      } catch (err) {
        console.error(err)
      }
    })
  })
  ws.cell(1,19,1,21,true)
  wb.write('vacant_premise.xlsx')
}
