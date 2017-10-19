package com.jruanogalvez.excelapi;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Main {
    
    public static void main(String[] args) {
        transformToArray("C:\\Users\\matinal\\Downloads\\example.xlsx");
    }
    
    public static ArrayList<ArrayList<String>> transformToArray(String inputFile) {
        FileInputStream input;
        XSSFWorkbook wb;
                
        int rowCount = 0;
        int cellCount = 0;
        
        ArrayList<ArrayList<String>> data = new ArrayList<>();
        ArrayList<String> rowList = new ArrayList<>();
        
        try {
            input = new FileInputStream(new File(inputFile));
            wb = new XSSFWorkbook(input);
            
            for(int sheetCount = 0; sheetCount < wb.getNumberOfSheets(); sheetCount++) {
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
                            rowList.add(thisCell.getStringCellValue());
                            
                        } else if (thisCell.getCellTypeEnum() == CellType.NUMERIC) {
                            rowList.add(Double.toString(thisCell.getNumericCellValue()));
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
        
        for (List<String> element : data) {
            System.out.println(element);
        }
        
        return data;
    }
    
    public void transformToExcel(ArrayList<ArrayList<String>> data) {
        
        
        
    }
}
