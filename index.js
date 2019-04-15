var xl = require('excel4node')

var wb = new xl.Workbook()
var ws = wb.addWorksheet('Sheet 1')
var ws2 = wb.addWorksheet('Sheet 2')

var style = wb.createStyle({
    font: {
        color: '#FF0800',
        size: 12,
    },
    numberFormat: '$#,##0.00; ($#.##0.); -'
})

for(var i = 1; i <= 5;  i++){
    ws.row(i).setHeight(120)
    ws.column(5).setWidth(200)

    ws.cell(i,1)
    .number(100)
    .style(style)
    
    ws.cell(i,2)
    .number(200)
    .style(style)
    
    ws.cell(i,3)
    .formula('A1 + B1')
    .style(style)
    
    ws.addImage({
        path: 'images/logoform.jpg',
        type: 'picture',
        position: {
            type: 'oneCellAnchor',
            from: {
              col: 5,
              colOff: '0.1in',
              row: i,
              rowOff: '0.1in',
            }
        }
      })
    
}




wb.write('output/Excel.xlsx');