const Excel = require('exceljs');
const fs = require("fs");


const setup = {
    folderTarget : "./img",
    imageWidth  : 50,
    imageHeight : 150 
}
async function wireteFile(setup){
    
    const workbook = new Excel.Workbook();
    const worksheet = workbook.addWorksheet("NEW SHEET");

    worksheet.properties.defaultRowHeight = setup.imageHeight;
    worksheet.pageSetup.horizontalCentered = true ;
    worksheet.pageSetup.verticalCentered = true;
    const {imageWidth , imageHeight} = setup;

    worksheet.columns = [
        {header: 'No', key: 'no', width: 10},
        {header: 'Word Aassamble', key: 'wa', width: 32}, 
        {header: 'แก้ไข.', key: 'edit', width: 10},
        {header: 'พอผ่าน.', key: 'ok_pass', width: 10},
        {header: 'ผ่าน.', key: 'pass', width: 10},
        {header: 'Illustration.', key: 'ill', width:  imageWidth},
        {header: 'Case.', key: 'case', width: 10},
    ];
    const first = worksheet.getRow(0);
    first.height = 40;
    console.log(first.height);

    const files = fs.readdirSync(setup.folderTarget);

    console.log(imageWidth , imageHeight);
    console.log(files);

    for(let [index,file] of files.entries()){

        worksheet.addRow({no: (index+1), wa: "", edit : "",ok_pass : "" ,pass:"",ill : "" , case : ""});
        //---------- write image --------------
        const imageId = workbook.addImage({
            filename: `${setup.folderTarget}/${file}`,
            extension: file.substring(".")[1],
        });

        worksheet.addImage(imageId, {
            tl: { col: 5, row: (index+1) },
            ext: { width: imageWidth+300, height: imageHeight+50},
            editAs: 'oneCell'
        });
    }

    await workbook.xlsx.writeFile('export.xlsx');
};

(async ()=>{
    await wireteFile(setup);
    console.log("File is written!");

})();