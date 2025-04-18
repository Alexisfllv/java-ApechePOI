package edu.com.Excel;

import org.apache.commons.codec.DecoderException;
import org.apache.commons.codec.binary.Hex;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.*;

import java.io.FileOutputStream;
import java.io.OutputStream;

public class ColoresExcel {
    public static void main(String[] args) {

        /* Colores */
        XSSFColor verdeClaro = crearColor("4EC2D7");


        //Libro
        XSSFWorkbook wb = new XSSFWorkbook();

        //Hoja
        XSSFSheet sheet0 = wb.createSheet();



        // Fila
        XSSFRow row = sheet0.createRow(0);
        XSSFRow row1 = sheet0.createRow(1);

        // Crear Celda  -> Columna
        XSSFCell cell = row.createCell(0);
        XSSFCellStyle estiloCelda = wb.createCellStyle();

        XSSFCell cell2 = row1.createCell(0);
        XSSFCellStyle estiloCelda2 = wb.createCellStyle();


        /* Configuracion estilos */
        estiloCelda.setFillForegroundColor(IndexedColors.AQUA.getIndex());
        estiloCelda.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        estiloCelda2.setFillForegroundColor(verdeClaro);
        estiloCelda2.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        /* Configuracion celda */
        cell.setCellStyle(estiloCelda);
        cell.setCellValue("Coloressssssss");

        cell2.setCellStyle(estiloCelda2);
        cell2.setCellValue("Pruebaaaaa");

        /* Configuacion Hoja */
        sheet0.autoSizeColumn(0);








        // try
        try {
            OutputStream out =  new FileOutputStream("ArchivoColoresExcel.xlsx");
            wb.write(out);

            // liberar recursos
            wb.close();
            out.close();
        } catch (Exception e){
            e.printStackTrace();
        }

    }


    // metodo estatico de color

    public static  XSSFColor crearColor(String Hexa){


        try {
            byte[] rgb = Hex.decodeHex(Hexa);
            return new XSSFColor(rgb);

        } catch (DecoderException e) {
            e.printStackTrace();
            throw new RuntimeException("Error al crear el color.");
        }
    }
}
