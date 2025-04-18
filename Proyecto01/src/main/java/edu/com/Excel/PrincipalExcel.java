package edu.com.Excel;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.OutputStream;

public class PrincipalExcel {

    public static void main(String[] args) {

//        // libro con = .xlsx 2007
//        Workbook libro = new XSSFWorkbook();
//
//        // libro con = .xls 1997 -2003
//        Workbook libro2 = new HSSFWorkbook();

        // 1) crear un libro

        Workbook libroPrincipal = new XSSFWorkbook();

        // 2) crear las hojas
        Sheet sheet = libroPrincipal.createSheet("Personas");

        // 3) crear las filas
        Row cabecera = sheet.createRow(2);
        Row fila2 = sheet.createRow(3);
        Row fila3 = sheet.createRow(4);

        // 4) crear las columnas
        Cell nombre = cabecera.createCell(1);
        Cell edad = cabecera.createCell(2);
        Cell ciudad = cabecera.createCell(3);
        // --
        Cell nombre1 = fila2.createCell(1);
        Cell edad1 = fila2.createCell(2);
        Cell ciudad1 = fila2.createCell(3);
        //
        Cell nombre2 = fila3.createCell(1);
        Cell edad2 = fila3.createCell(2);
        Cell ciudad2 = fila3.createCell(3);

        // 5) setear valor a la celda
        nombre.setCellValue("Nombre");
        edad.setCellValue("Edad");
        ciudad.setCellValue("Ciudad");

        // -- fila 2
        nombre1.setCellValue("Santiago");
        edad1.setCellValue(23);
        ciudad1.setCellValue("Medellin");

        // -- fila3
        nombre2.setCellValue("Anyi");
        edad2.setCellValue(22);
        ciudad2.setCellValue("Bogota");

        try {
            OutputStream out =  new FileOutputStream("ArchivoExcel.xlsx");
            libroPrincipal.write(out);

            // liberar recursos
            libroPrincipal.close();
            out.close();
        } catch (Exception e){
            e.printStackTrace();
        }


    }
}
