var Axios = require('axios')
var xl = require('excel4node')

module.exports = {
    createReport: async (callback) => {
        await Axios.get('http://localhost/senjabina/api/report/generate-excel?module=non-commercial&tabs=2019-d-nc16f05a%20melaka')
            .then(response => {
                let resp = response.data.response
                createBook(resp)
                callback(resp)
            }).catch(err => {
                console.log(err)
            })
    }
}

createBook = (data) => {
    var wb = new xl.Workbook()
    var style = wb.createStyle({
        font: {
          size: 9,
        },
        alignment: {
            vertical: 'top'
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
    ws.row(1).setHeight(30)
    ws.cell(1, 1).string('ID')
    ws.cell(1, 2).string('Bill No')
    ws.cell(1, 3).string('SAN')
    ws.cell(1, 4).string('OWNER NAME')
    ws.cell(1, 5).string('PROPERTY ADDRESS')
    ws.cell(1, 6).string('BALANCE AS AT')
    ws.cell(1, 7).string('DIFF BETWEEN BAL AS PER BILL AND BAL@')
    ws.cell(1, 8).string('Occupier (Owner/Tenant')
    ws.cell(1, 9).string('Owner Name Correct')
    ws.cell(1, 10).string('Please specify correct owner name')
    ws.cell(1, 11).string('Owner\'s tel no')
    ws.cell(1, 12).string('Tenant\'s Name')
    ws.cell(1, 13).string('Tenant\s tel no')
    ws.cell(1, 14).string('Occupier Nationality')
    ws.cell(1, 15).string('Property Usage')
    ws.cell(1, 16).string('Property Type')
    ws.cell(1, 17).string('Name of shop/company if Non-Domestic')
    ws.cell(1, 18).string('Nature of business if Non-Domestic')
    ws.cell(1, 19).string('DR Code')
    ws.cell(1, 20).string('Black Area')
    ws.cell(1, 21).string('High Rise')
    ws.cell(1, 22).string('Reason Refuse To Pay')
    ws.cell(1, 23).string('Staff')
    ws.cell(1, 24).string('Photo')

    data.results.forEach((value, k) => {
        var key = k + 2
        ws.row(key).setHeight(30)
        ws.cell(key, 1).string(value.seq_id).style(style)
        ws.cell(key, 2).string(value.upload_content.bill_no).style(style)
        ws.cell(key, 3).string(value.upload_content.san).style(style)
        ws.cell(key, 4).string(value.upload_content.owner_1).style(style)
        var address = value.upload_content.prop_address_1
        address += ` ${value.upload_content.prop_address_2}`
        address += ` ${value.upload_content.prop_address_3}`
        address += ` ${value.upload_content.prop_address_4}`
        address += ` ${value.upload_content.prop_address_5}`
        ws.cell(key, 5).string(address).style(style)
        ws.cell(key, 6).string(value.upload_content.balance_as_bill).style(style)
        ws.cell(key, 6).string('').style(style)
        var task_perform = (value.task_perform != null) ? value.task_perform: false
        ws.cell(key, 7).string((task_perform && task_perform.form_info.owner_name_correct) ? task_perform.form_info.owner_name_correct : '').style(stylecenter)
        ws.cell(key, 8).string((task_perform && task_perform.form_info.correct_ownername) ? task_perform.form_info.correct_ownername : '').style(style)
        ws.cell(key, 9).string((task_perform) ? task_perform.form_info.owner_name_correct : '').style(style)
        var owner_phone = (task_perform && task_perform.form_info.owner_tel_no) ? task_perform.form_info.owner_tel_no : ''
        owner_phone = owner_phone + ' ' + ((task_perform && task_perform.form_info.owner_mobile_no) ? task_perform.form_info.owner_mobile_no : '')
        owner_phone = owner_phone + ' ' + ((task_perform && task_perform.form_info.owner_fax) ? task_perform.form_info.owner_fax : '')
        ws.cell(key, 10).string(owner_phone).style(style)
        ws.cell(key, 11).string((task_perform && task_perform.form_info.tenant_name) ? task_perform.form_info.tenant_name : '').style(style) 
        var tenant_phone = (task_perform && task_perform.form_info.tenant_mobile_no) ? task_perform.form_info.tenant_mobile_no : ''
        tenant_phone = tenant_phone + ' ' + ((task_perform && task_perform.form_info.tenant_fax) ? task_perform.form_info.tenant_fax : '')
        tenant_phone = tenant_phone + ' ' + ((task_perform && task_perform.form_info.tenant_tel_no) ? task_perform.form_info.tenant_tel_no : '')
        ws.cell(key, 12).string(tenant_phone).style(style)  
        ws.cell(key, 13).string((task_perform && task_perform.form_info.occupier_nationality) ? task_perform.form_info.occupier_nationality : '').style(style)
        ws.cell(key, 14).string((task_perform && task_perform.form_info.property_type_usage) ? task_perform.form_info.property_type_usage : '').style(style)
        ws.cell(key, 15).string((task_perform && task_perform.form_info.property_type) ? task_perform.form_info.property_type : '').style(stylecenter)
        ws.cell(key, 16).string((task_perform && task_perform.form_info.name_of_shop_company) ? task_perform.form_info.name_of_shop_company : '').style(style)
        ws.cell(key, 17).string((task_perform && task_perform.form_info.nature_of_business) ? task_perform.form_info.nature_of_business : '').style(style)
        ws.cell(key, 18).string((task_perform && task_perform.form_info.dr_code) ? task_perform.form_info.dr_code : '').style(stylecenter)
        ws.cell(key, 19).string((task_perform && task_perform.form_info.blackarea) ? task_perform.form_info.blackarea : '').style(stylecenter)
        ws.cell(key, 20).string((task_perform && task_perform.form_info.highrise) ? task_perform.form_info.highrise : '').style(stylecenter)
        ws.cell(key, 21).string((task_perform && task_perform.form_info.remarks) ? (task_perform.form_info.remarks ): '').style(style)
        ws.cell(key, 22).string(value.staff).style(style)
        ws.cell(key, 23).string('Photo')
    })

    wb.write('create_report.xlsx')
}

// var style = wb.createStyle({
//     font: {
//         color: '#FF0800',
//         size: 12,
//     },
//     numberFormat: '$#,##0.00; ($#.##0.); -'
// })

// for(var i = 1; i <= 5;  i++){
//     ws.row(i).setHeight(120)
//     ws.column(5).setWidth(200)

//     ws.cell(i,1)
//     .number(100)
//     .style(style)

//     ws.cell(i,2)
//     .number(200)
//     .style(style)

//     ws.cell(i,3)
//     .formula('A1 + B1')
//     .style(style)

//     ws.addImage({
//         path: '../images/logoform.jpg',
//         type: 'picture',
//         position: {
//             type: 'oneCellAnchor',
//             from: {
//               col: 5,
//               colOff: '0.1in',
//               row: i,
//               rowOff: '0.1in',
//             }
//         }
//       })

// }

// wb.write('../output/Excel_2.xlsx');