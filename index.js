const fs = require("fs");
const XLSX = require("xlsx");
const jsontoxml = require("jsontoxml");
const workbookAdministra = XLSX.readFile("administra.xlsx");
const workbookInventti = XLSX.readFile("inventti.csv");

function getRowsOfSheet(worksheet){
    if(worksheet){
        return XLSX.utils.decode_range(worksheet["!ref"]).e.r;
    }

    console.log("Worksheet passada é null/n");
}

function getNotasAdministra(worksheet){
    let rows = getRowsOfSheet(worksheet);
    let result = [];
    let j = 0;
    for(let i = 0; i<rows; i++){     
        const cupomFiscal = worksheet['D' + i];
        const nota = worksheet['H' + i];
        if(cupomFiscal && cupomFiscal.v == "CUPOM FISCAL"){
            result[j] = parseInt(nota.v.trim());
            j++;
        }
    }

    return result;

}

function getNotasInventti(worksheet){
    let rows = getRowsOfSheet(worksheet);
    let result = [];
    let continuar = true;
    let i = 2;
    let j = 0;
    do{
        const nota = worksheet['O' + i];
        if(nota && nota.v){
            result[j] = parseInt(nota.v);
            j++;
            i++
        }else{
            continuar = false
        }
    }while(continuar)
    
    return result;
}

function getWorksheet(workbook){
    let worksheet;
    for (const sheetName of workbook.SheetNames){
        worksheet = workbook.Sheets[sheetName];
    }
    return worksheet;
}

function getNotasNaoEncontradas(notasAdministra, notasInventti){
    let result = {
        administra: [],
        inventti: []
    }

    result.administra = getNotasOffArray(notasAdministra, notasInventti);
    result.inventti = getNotasOffArray(notasInventti, notasAdministra);

    return result;
}

function getNotasOffArray(a, b){
    let result = [];
    let j = 0;
    for(let i = 0; i < a.length; i++){
        posicaoOnArray = b.indexOf(a[i]);
        if(posicaoOnArray == -1){
            result[j] = a[i];
            j++;
        }
    }

    return result;
}

function sortNotas(notas){
    notas = notasAdministra.sort((a, b) => a - b);
}


let worksheetAdministra = getWorksheet(workbookAdministra);
let worksheetInventti = getWorksheet(workbookInventti)
let notasAdministra = getNotasAdministra(worksheetAdministra);
let notasInventti = getNotasInventti(worksheetInventti);

let notasNaoEncontradas = getNotasNaoEncontradas(notasAdministra, notasInventti);

console.log(notasAdministra);
console.log(notasInventti);
console.log(notasNaoEncontradas);


