import * as XLSX from 'xlsx/xlsx.mjs';
import * as fs from 'fs';
import esMain from 'es-main';
import {deleteRow} from './util.mjs';

/* load the codepage support library for extended support with older formats  */
//import * as cpexcel from 'xlsx/dist/cpexcel.full.mjs';
//XLSX.set_cptable(cpexcel);
XLSX.set_fs(fs);

export const dump = (file, opts = {}, workbook = XLSX.readFile(file, opts)) => workbook.SheetNames.reduce((ret, name) => {
    ret[name] = XLSX.utils.sheet_to_json(workbook.Sheets[name], {raw:true});
    return ret;
}, {});

export const clean = (worksheet)=>{
    
     deleteRow(worksheet, 1);
    // console.log('clean', worksheet)

 //   console.log(worksheet);
//    deleteRow(worksheet, 0);

   // console.log(worksheet);
   return worksheet;
}
export function main(files) {
    console.warn(`processing ${files}`);
    files.forEach(v => {
        const workbook = XLSX.readFile(v, {raw:true, headers:0});
        const sheet = clean(workbook.Sheets[workbook.SheetNames.find(v=>/schedule/i.test(v))]);
        console.log(`${v}--\n`);
        console.log(XLSX.utils.sheet_to_row_object_array(sheet, {raw:true, header:1}));
    });
}

if (esMain(import.meta)) {
    main(process.argv.slice(2));
}