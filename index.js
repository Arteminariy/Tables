const http = require('http');
const fs = require('fs');
const path = require('path');
const hostname = '127.0.0.1';
const port = 3000;
const ExcelJS = require("ExcelJS")

const server = http.createServer(async function (req, res) {
    if(req.method == "GET"){
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile("Прайс от 07.03.xlsx")

        new Promise((res, req) => {
            workbook.eachSheet((worksheet, sheetId) => {
                res((worksheet.getCell("A1").value = "Success!"))
            })
        }).then(workbook.xlsx.writeFile("New-from-nodejs.xlsx"))
        res.end(".xlsx created")
    }
})


    // let filePath = path.join(__dirname, 'public', req.url === '/' ? 'index.html' : req.url)
    // fs.readFile(filePath, (err, data) => {
    //     if (err) {
    //         throw err
    //     } 
    //     else {
    //         res.writeHead(200, {
    //             'Content-Type': 'text/html'
    //         })
    //     }
    // })
    

//   if (res.url === '/'){
//     fs.readFile(path.join(__dirname, 'public', 'index.html'), (err, data) => {
//         if (err) {
//             throw err
//         }
//         res.writeHead(200, {
//             'Content-Type': 'text/html'
//         })
//         res.end(data);
//     })
// }

server.listen(port, hostname, () => {
  console.log(`Server running at http://${hostname}:${port}/`);
});



// const name = "Окрашенный"

// let reader = new FileReader();

// let sheet1 = workbook.xlsx.readFile(document.getElementById("old").value)
// let sheet2 = workbook.xlsx.readFile(document.getElementById("new").value)

// new Promise((res, rej) => {
//     workbook.eachSheet((worksheet, sheetId) => {
//         res((worksheet.getCell("AS8").value = AS8))
//     })
// }).then(workbook.xlsx.writeFile(`${name}.xlsx`))
// res.end(`${name}.xlsx was created`)

// function showFile(input) {
//     let file = input.files[0];

//     console.log(`File name: ${file.name}`); // например, my.png
//     console.log(`Last modified: ${file.lastModified}`); // например, 1552830408824
// }



// let gradient = ['800000', '8B0000', 'B22222','FF0000','FFD700','9ACD32', 'ADFF2F','7CFC00', '00FF00']
// function converter(num) {
//         if (num == 0) {
//             return 'D'
//         }
//         if (num == 1) {
//             return 'E'
//         }
//         if (num == 2) {
//             return 'F'
//         }
//         if (num == 3) {
//             return 'G'
//         }
//         if (num == 4) {
//             return 'H'
//         }
// }
// function painter(prev, com) {
//         if (old == None || old == 0){
//             return PatternFill(fill_type = 'solid', start_color = 'FFFFFF')
//         }
//         else{
//             let power
//             if (prev - com < 0 && abs(prev - com) > (prev * 0.3)){power = 0}    
//             if (prev - com < 0 && abs(prev - com) > (prev * 0.1)){power = 1}    
//             if (prev - com < 0 && abs(prev - com) > (prev * 0.05)){power = 2}    
//             if (prev - com < 0 && abs(prev - com) <= (prev * 0.05)){power = 3}    
//             if (prev - com == 0) {let power = 4}                        
//             if (prev - com > 0 && abs(prev - com) <= (prev * 0.05)){power = 5}    
//             if (prev - com > 0 && abs(prev - com) > (prev * 0.05)){power = 6}    
//             if (prev - com > 0 && abs(prev - com) > (prev * 0.1)){power = 7}    
//             if (prev - com > 0 && abs(prev - com) > (prev * 0.3)){power = 8}    
//         return PatternFill(fill_type = 'solid', start_color = gradient[power])
//         }
// }
// function main(prev, com) {
//     let prevData = {
//         prevDataDil_1,
//         prevDataDil_2,
//         prevDataDil_3,
//         prevTraderPrice,
//         prevClientPrice
//     }
//     let comData = {
//         comDataDil_1,
//         comDataDil_2,
//         comDataDil_3,
//         comTraderPrice,
//         comClientPrice
//     }
    
//     for (let i = 0; i < 700; i++) {
//         for (let j = 0; j < 5; j++) {
//             prevData[j][i].append(sheet1[converter(j) + str(i)].value)
//         }
//     }
//     for (let i = 0; i < 700; i++) {
//         for (let j = 0; j < 5; j++) {
//             comData[j][i].append(sheet2[converter(j) + str(i)].value)
//         }
//     }
//     for (let i = 0; i < 700; i++) {
//         for (let j = 0; j < 5; j++) {
//             if(prevData[j][i] == None) {
//                 prevData[j][i] == 0
//             }
//             if(comData[j][i] == None) {
//                 comData[j][i] == 0
//             }
            
//         }
        
//     }
//     for (let i = 0; i < sheet2.max_row; i++) {
//         sheet2['D' + str(i + 7)].fill = painter(prevData.prevDataDil_1[i], comData.comDataDil_1[i])
//         sheet2['E' + str(i + 7)].fill = painter(prevData.prevDataDil_2[i], comData.comDataDil_2[i])
//         sheet2['F' + str(i + 7)].fill = painter(prevData.prevDataDil_3[i], comData.comDataDil_3[i])
//         sheet2['G' + str(i + 7)].fill = painter(prevData.prevTraderPrice[i], comData.comTraderPrice[i])
//         sheet2['H' + str(i + 7)].fill = painter(prevData.prevClientPrice[i], comData.comClientPrice[i])
//     }
// }
