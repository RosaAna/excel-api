package com.jruanogalvez.excelapi;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Main {
    
    /*public static void main(String[] args) {
        transformToArray("C:\\Users\\video\\Downloads\\example.xlsx");
    }*/
    
    public static ArrayList<ArrayList<String>> transformToArray(String inputFile) {
        FileInputStream input;
        XSSFWorkbook wb;
        
        ArrayList<ArrayList<String>> data = new ArrayList<>();
        ArrayList<String> rowList = new ArrayList<>();
        
        try {
            input = new FileInputStream(new File(inputFile));
            wb = new XSSFWorkbook(input);
            
            for(int sheetCount = 0; sheetCount < wb.getNumberOfSheets(); sheetCount++) {
                Sheet firstSheet = wb.getSheetAt(0);
                
                for(int i = 0; i <= firstSheet.getLastRowNum(); i++) {
                    Row thisRow = firstSheet.getRow(i);
                  
                    if(thisRow != null) {
                        for(int j = 0; j <= thisRow.getLastCellNum(); j++) {
                            Cell thisCell = thisRow.getCell(j);

                            if(thisCell != null) {
                                if(thisCell.getCellTypeEnum() == CellType.STRING)
                                    rowList.add(thisCell.getStringCellValue());
                                if(thisCell.getCellTypeEnum() == CellType.NUMERIC)
                                    rowList.add(Double.toString(thisCell.getNumericCellValue()));
                                if(thisCell.getCellTypeEnum() == CellType.BLANK)
                                    rowList.add(" ");
                            } else {
                                rowList.add(" ");
                            }
                        }
                    } 
                    
                    data.add((ArrayList<String>) rowList.clone());
                    rowList.clear();
                    
                }
            }
            
            wb.close();
            input.close();
            
        } catch (FileNotFoundException ex) {
            System.out.println("No se puede encontrar el archivo especificado.");
            
        } catch (IOException ex) {
            System.out.println("No se puede leer o escribir el archivo.");
            
        }
        
        /*data.forEach((element) -> {
            System.out.println(element);
        });*/
        
        //transformToExcel(data, "C:\\Users\\video\\Downloads\\example2.xlsx");
        return data;
    }

    public static void transformToExcel(ArrayList<ArrayList<String>> data, String path) {
        ExcelBook e = new ExcelBook(data, path);
        e.writeExcelSheet();
        e.writeExcelFile();
      
    }
}
