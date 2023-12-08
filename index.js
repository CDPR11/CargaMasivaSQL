
const excelInput = document.getElementById("excel-input");
const contentXls = document.getElementById("content-xls");
const btnExportDocument = document.getElementById("btnExportDocument");
const btnClear = document.getElementById("btnClear");
const tableInput = document.getElementById("table-input")
contentXls.value = ""
class Excel{
    constructor(content){
        this.content = content;
    }

    header(){
        return this.content[0]
    }

    rows(){
        return this.content.slice(1,this.content.lenght)
    }
}

excelInput.addEventListener('change',async function(){
    const content = await readXlsxFile(excelInput.files[0])
    
    const excel = new Excel(content)
    let texto = ""
    let tableName = tableInput.value
    for(let i = 0; i < excel.header().length; i++){
        if( i === (excel.header().length - 1))
            texto = texto + excel.header()[i] + ""
        else    
            texto = texto + excel.header()[i] + ","
    }
    
    let contentRow = ''
    for(let i = 0; i < excel.rows().length; i++){
        for( let j = 0; j < excel.rows()[i].length; j++){
            //console.log(excel.rows()[j].length)

            if(j === excel.rows()[j].length - 1 )
                contentRow = content + `'${excel.rows()[i][j]}'` 
            else
                contentRow = content + `'${excel.rows()[i][j]}'` + ","
        }
        console.log(contentRow)
        contentRow = ""
    }

    //console.log(contentRow)
    // let arrayValidaciones = ['ClaveCodigoTransporteAereo','VersionComplemento']
    // let whereString = '';
    // if(arrayValidaciones.length > 0){
    //     if(arrayValidaciones.length>1){
    //         for(let i = 0; i < arrayValidaciones.length; i++){
    //             for(let j = 0; j<=)
    //             whereString = whereString + arrayValidaciones[i] + 'AND'
    //         }
    //     }else{
    //         whereString = arrayValidaciones[0]
    //     }
    // }
    

    let estructura = `IF NOT EXISTS ( SELECT 1 FROM Tabla WHERE Campo = Valor)
        BEGIN
            INSERT INTO ${tableName} (${texto}) VALUES (value1, value2, value3, ...); 
        END
        ELSE
        BEGIN
            PRINT('Este registro ya existe');
        END`

    console.log(texto)
    contentXls.value = estructura
    // console.log(excel.header())
    // console.log(excel.rows())
})

btnExportDocument.addEventListener('click', function(){
   if(contentXls.value){
        exportarTexto(contentXls.value)
   }else{
        alert("No hay contenido que exportar")
   }

})

btnClear.addEventListener('click', function(){
    console.log("clickea limpiar")
    contentXls.value = ""
})

function exportarTexto(content){
    var fileContent = content;
    var fileName = 'sampleFile.txt';

    const blob = new Blob([fileContent], { type: 'text/plain' });
    const a = document.createElement('a');
    a.setAttribute('download', fileName);
    a.setAttribute('href', window.URL.createObjectURL(blob));
    a.click();
}