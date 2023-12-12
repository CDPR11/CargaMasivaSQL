const excelInput = document.getElementById("excel-input");
const databaseInput = document.getElementById("database-input");
const contentXls = document.getElementById("content-xls");
const btnExportDocument = document.getElementById("btnExportDocument");
const btnClear = document.getElementById("btnClear");
const tableInput = document.getElementById("table-input")
const divChecks = document.getElementById("checksD")
const btnRealizarScript = document.getElementById("btnRealizarScript")
let nombreDocumento = ''

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
    let arrayChecks = []
    let tableName = tableInput.value
    for(let i = 0; i < excel.header().length; i++){
        if( i === (excel.header().length - 1))
            texto = texto + excel.header()[i] + ""
        else    
            texto = texto + excel.header()[i] + ","
        arrayChecks.push(excel.header()[i])
    }

    nombreDocumento = tableInput.value
    
    let checksButtons =''
    for(let i = 0; i < arrayChecks.length; i++){
        checksButtons = `${checksButtons}\n<div class="col-4 col-md-4 col-lg-3">
                            <input type="checkbox" class="btn-check" id="btn-check${i}" autocomplete="off">
                            <label class="btn btn-primary" for="btn-check">${arrayChecks[i]}</label>
                        </div>`    
    }
    divChecks.innerHTML = checksButtons;
})

btnRealizarScript.addEventListener('click', async function(){
    const content = await readXlsxFile(excelInput.files[0])
    
    const excel = new Excel(content)
    let texto = ""
    let arrayChecks = []
    let tableName = tableInput.value
    for(let i = 0; i < excel.header().length; i++){
        if( i === (excel.header().length - 1))
            texto = texto + `[${excel.header()[i]}]` + ""
        else    
            texto = texto + `[${excel.header()[i]}]` + ","
        arrayChecks.push(excel.header()[i])
    }
    
    let arrayCheckeds = []
    for(let i = 0; i < arrayChecks.length; i++){
        
        if(document.getElementById('btn-check'+i).checked){
            let json = {
                "posicion":i,
                "name":arrayChecks[i]
            }
            arrayCheckeds.push(json)
        }
    }

    let whereString = ''
    let contentRow = ''
    let estructura =`USE [${databaseInput.value}]
    GO
    BEGIN TRANSACTION;
        BEGIN TRY\n`
    for(let i = 0; i < excel.rows().length; i++){
        contentRow = ""
        whereString = ""
        for( let j = 0; j < excel.rows()[i].length; j++){
            if(arrayCheckeds.length > 0){
                for(let k = 0; k < arrayCheckeds.length; k++){
                    if(arrayCheckeds[k].posicion === j){
                        whereString = `${whereString} ${arrayCheckeds[k].name} = '${excel.rows()[i][j]}' ${k != (arrayCheckeds.length - 1)?'AND':''}`;
                    }
                }
            }
            if(typeof(excel.rows()[i][j]) === 'object'){
                excel.rows()[i][j] = excel.rows()[i][j] != null ? formatearFecha(excel.rows()[i][j]) : ""
            }

            if(j === excel.rows()[j].length - 1 )
                contentRow = `${contentRow} ${excel.rows()[i][j].toString().toLowerCase().includes('null') || excel.rows()[i][j].toString().toLowerCase().includes('getdate()') ?excel.rows()[i][j]:`'${excel.rows()[i][j]}'`}` 
            else
                contentRow = `${contentRow} ${excel.rows()[i][j].toString().toLowerCase().includes('null') || excel.rows()[i][j].toString().toLowerCase().includes('getdate()') ?excel.rows()[i][j]:`'${excel.rows()[i][j]}'`},`
            
        }

        estructura = `${estructura}\n       IF NOT EXISTS ( SELECT 1 FROM ${tableInput.value} WHERE ${whereString})
        BEGIN
            INSERT INTO [dbo].[${tableName}] (${texto}) VALUES (${contentRow}); 
            PRINT('Se inserto con exito el registro ${whereString.replaceAll("'","")}')
        END
        ELSE
        BEGIN
            PRINT('El registro con ${whereString.replaceAll("'","")} ya existe');
        END`
        
    }

    estructura = estructura + `\nEND TRY 
    BEGIN CATCH
        SELECT 'Insert Incorrecto'  AS [Mensaje ejecución]   ,CAST (ERROR_NUMBER() AS NVARCHAR(1000)) AS [Dato_Error_ERROR_NUMBER()  ]
        ,CAST (ERROR_SEVERITY() AS NVARCHAR(1000)) AS [Dato_Error_ERROR_SEVERITY() ]   ,CAST (ERROR_STATE() AS NVARCHAR(1000)) AS [Dato_Error_ERROR_STATE()  ]
        ,CAST (ERROR_PROCEDURE() AS NVARCHAR(1000)) AS [Dato_Error_ERROR_PROCEDURE()]   ,CAST (ERROR_LINE() AS NVARCHAR(1000)) AS [Dato_Error_ERROR_LINE()   ]
        ,CAST (ERROR_MESSAGE() AS NVARCHAR(1000)) AS [Dato_Error_ERROR_MESSAGE() ]  
        IF @@TRANCOUNT > 0 ROLLBACK TRANSACTION;
    END CATCH
        IF @@TRANCOUNT > 0 -- Verifica correcta ejecucion de nivel de transaccion
    BEGIN
        COMMIT TRANSACTION;  
        SELECT 'Insert Ejecutado';
    END`;
    contentXls.value = sqlFormatter.format(estructura)
})

function formatearFecha(fechaStr) {
    var fecha = new Date(fechaStr);
    
    fecha.setUTCHours(0, 0, 0, 0);
    var year = fecha.getUTCFullYear()
    var month = ('0' + (fecha.getUTCMonth() + 1)).slice(-2);
    var day = ('0' + fecha.getUTCDate()).slice(-2);

    var formattedDate = year + '-' + month + '-' + day;

    return formattedDate;
}

const copiarContenido = async () => {
    try {
      await navigator.clipboard.writeText(contentXls.value);
      alert('Contenido copiado al portapapeles');
    } catch (err) {
      console.error('Error al copiar: ', err);
    }
}

btnExportDocument.addEventListener('click', function(){
   if(contentXls.value){
        exportarTexto(contentXls.value)
   }else{
        alert("No hay contenido que exportar")
   }

})

btnClear.addEventListener('click', function(){
    contentXls.value = ""
})

function exportarTexto(content){
    var fileContent = content;
    var fileName = `[INSERT]_${nombreDocumento}.sql`;

    const blob = new Blob([fileContent], { type: 'text/plain' });
    const a = document.createElement('a');
    a.setAttribute('download', fileName);
    a.setAttribute('href', window.URL.createObjectURL(blob));
    a.click();


}
