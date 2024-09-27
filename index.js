import XlsxPopulate from "xlsx-populate";

//crear archivos excel desde cero
/* XlsxPopulate.fromBlankAsync()
    .then(workbook => {
        workbook.sheet(0).cell('A1').value('Hola mundo');
        return workbook.toFileAsync("./salida.xlsx");
    }) */

async function crearXlsx(){    
    /* const workbook = await XlsxPopulate.fromBlankAsync()
        .then(workbook =>{
            //cell siempre en mayuscula!!
            workbook.sheet(0).cell('A1').value('Nombre')
            workbook.sheet(0).cell('B1').value('Apellido')
            workbook.sheet(0).cell('C1').value('Dni')
            workbook.sheet(0).cell('A2').value('Angeles')
            workbook.sheet(0).cell('B2').value('Segovia')
            workbook.sheet(0).cell('C2').value('35261329')
            workbook.sheet(0).cell('A3').value('Martin')
            workbook.sheet(0).cell('B3').value('Gonzalez')
            workbook.sheet(0).cell('C3').value('30841782')
            workbook.toFileAsync("./Salida4.xlsx")
        }) */

    //editar las hojas        
    const workbook = await XlsxPopulate.fromBlankAsync()
    workbook.sheet('Sheet1').cell('A1').value([
        [ 'Nombre', 'Apellido', 'Dni' ],
        [ 'Angeles', 'Segovia', 35261329 ],
        [ 'Angeles', 'Segovia', 35261329 ],
        [ 'Martin', 'Gonzalez', 30841782 ]
    ])
    workbook.sheet('Sheet1').cell('A5').value([
        [ 'Nombre', 'Apellido', 'Dni' ],
        [ 'Martin', 'Gonzalez', 30841782 ],
        [ 'Angeles', 'Segovia', 35261329 ]
    ])
    workbook.toFileAsync('./Salida4.xlsx')
}

async function leerXlsx() {
    const workbook = await XlsxPopulate.fromFileAsync('Salida4.xlsx')
    //leer por celda
    const valor1 = workbook.sheet('Sheet1').cell('A2').value()
    const valor2 = workbook.sheet('Sheet1').cell('B2').value()
    console.log("Nombre y Apellido: "+valor1+" "+valor2)
    //leer por rango
    const rango1 = workbook.sheet('Sheet1').usedRange().value()
    //console.log("Los Datos Son :"+rango1)
    console.log(rango1)
    const rango2 = workbook.sheet('Sheet1').range('A1:B2').value()
    //console.log("Los Datos Son :"+rango1)
    console.log(rango2)
}   
async function crearSheet() {
    //creo el archivo
    const workbook = await XlsxPopulate.fromBlankAsync()
    workbook.addSheet('Hoja1')
    //carga el sheet
    workbook.sheet('Hoja1').cell('A1').value([
        [ 'Nombre', 'Apellido', 'Dni' ],
        [ 'Angeles', 'Segovia', 35261329 ],
        [ 'Angeles', 'Segovia', 35261329 ],
        [ 'Martin', 'Gonzalez', 30841782 ]
    ])
    workbook.sheet('Hoja1').range('E1:F10').value('hola')
    //crea sheet2
    workbook.addSheet('Hoja2')
    //cargha shhet 2
    workbook.sheet('Hoja2').range('A1:B100').value("")
    workbook.sheet('Hoja2').range('A6:B6').value("Hola")
    workbook.sheet('Hoja2').range('A8:B12').value("Hola")
    workbook.toFileAsync('./salida5.xlsx')
    //lista todo de sheet1
    const rango1 = workbook.sheet('Hoja1').usedRange().value()
    console.log(rango1)
    //lista todos los nombre sde la lista
    console.log(
        workbook.sheets().map((sheet)=>sheet.name())
    )
    
}
    
//crearXlsx() 
//leerXlsx()
crearSheet()