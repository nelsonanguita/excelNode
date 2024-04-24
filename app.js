const express = require('express');
const app = express();
const excel = require('excel4node');

app.get('/descargar-excel', (req, res) => {
  // Código para crear el archivo Excel como se muestra anteriormente
  
try {
        

        // Crear un nuevo libro de Excel
        const workbook = new excel.Workbook();

        // Añadir una hoja al libro
        const worksheet = workbook.addWorksheet('Sheet 1');

        // Escribir datos en la hoja

        for (let i = 1; i < 200000; i++) {
            worksheet.cell(i, 1).string('Hello');
            worksheet.cell(i, 2).string('World');
            worksheet.cell(i, 3).string('World');
            worksheet.cell(i, 4).string('World');
            worksheet.cell(i, 5).string('World');
            worksheet.cell(i, 6).string('World');
            worksheet.cell(i, 7).string('World');
            worksheet.cell(i, 8).string('World');
            worksheet.cell(i, 9).string('World');

            worksheet.cell(i, 10).string('World');
            worksheet.cell(i, 15).string('World');
         
        }
       
        // Opcionalmente, puedes formatear las celdas
        const headerStyle = workbook.createStyle({
        font: {
            bold: true,
        }
        });
        worksheet.cell(1, 1).style(headerStyle);

        // Guardar el libro en un archivo Excel
        workbook.write('archivo.xlsx');



  // Enviar el archivo Excel como respuesta
  res.download('archivo.xlsx');
} catch (error) {
    console.log(error)
}
});

app.listen(3000, () => {
  console.log('Servidor escuchando en el puerto 3000');
});
