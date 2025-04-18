package edu.com.Excel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

import java.io.FileOutputStream;
import java.io.OutputStream;

public class ManejoEstilos {

    public static void main(String[] args) {

        // 1) crear el libro
        XSSFWorkbook LibroEstilos = new XSSFWorkbook();

        // 2) crear las hojas
        XSSFSheet hoja = LibroEstilos.createSheet("Personas");


        // 3) crear filas  fila
        XSSFRow fila = hoja.createRow(1);

        // 4) crear celdas  columna
        XSSFCell celda = fila.createCell(2);
        XSSFCellStyle estiloCelda = LibroEstilos.createCellStyle();

        /* Configuracion de celda */
        estiloCelda.setFillForegroundColor(IndexedColors.AQUA.getIndex());
        estiloCelda.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        estiloCelda.setBorderBottom(BorderStyle.THIN);
        estiloCelda.setBorderLeft(BorderStyle.THIN);
        estiloCelda.setBorderRight(BorderStyle.THIN);
        estiloCelda.setBorderTop(BorderStyle.THIN);


        /* Configuracion de celda */
        celda.setCellValue("Estilos con apache poi");
        celda.setCellStyle(estiloCelda);

        /* Tamanio hoja */
        hoja.autoSizeColumn(2);
        hoja.setHorizontallyCenter(true);

        // 5)



        try {
            OutputStream out =  new FileOutputStream("ArchivoExcelEstilos.xlsx");
            LibroEstilos.write(out);

            // liberar recursos
            LibroEstilos.close();
            out.close();
        } catch (Exception e){
            e.printStackTrace();
        }

    }
}
