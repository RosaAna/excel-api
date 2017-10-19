package com.jruanogalvez.excelapi;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Main {
    
    public static void main(String[] args) {
        transformToArray("C:\\Users\\video\\Downloads\\example.xlsx");
    }
    
    public static String[][] transformToArray(String inputFile) {
        FileInputStream input;
        XSSFWorkbook wb;
                
        int rowCount = 0;
        int cellCount = 0;
        String[][] output = new String[4000][4000];
        
        try {
            input = new FileInputStream(new File(inputFile));
            wb = new XSSFWorkbook(input);
            
            //for(int sheetCount = 0; sheetCount <= wb.getNumberOfSheets(); sheetCount++) {
                Sheet firstSheet = wb.getSheetAt(0);
                Iterator<Row> iterator = firstSheet.iterator();
            
                while(iterator.hasNext()) {
                    Row thisRow = iterator.next();
                    Iterator<Cell> cellIterator = thisRow.cellIterator();
                    rowCount += 1;

                    while(cellIterator.hasNext()) {
                        Cell thisCell = cellIterator.next();
                        cellCount += 1;

                        if (thisCell.getCellTypeEnum() == CellType.STRING) {
                            System.out.println(thisCell.getStringCellValue());
                            output[rowCount][cellCount] = thisCell.getStringCellValue();

                        } else if (thisCell.getCellTypeEnum() == CellType.NUMERIC) {
                            output[rowCount][cellCount] = Double.toString(thisCell.getNumericCellValue());
                        }
                    }
                }
            //}
            
            wb.close();
            input.close();
            
        } catch (FileNotFoundException ex) {
            System.out.println("No se puede encontrar el archivo especificado.");
            
        } catch (IOException ex) {
            System.out.println("No se puede leer o escribir el archivo.");
            
        }
        
        return null;
    }
}
