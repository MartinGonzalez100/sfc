import XlsxPopulate from "xlsx-populate";

XlsxPopulate.fromBlankAsync()
    .then(workbook => {
        workbook.sheet(0).cell('A1').value('Hola mundo');
        return workbook.toFileAsync("./salida.xlsx");
    })