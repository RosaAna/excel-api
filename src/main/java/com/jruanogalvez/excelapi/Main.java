package com.jruanogalvez.excelapi;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author Jose Ruano
 */
public class Main {
    
    /**
     * This method transform the XLSX file into an dynamic Array.
     *      * 
     * @param inputFile the file input to transform it into an Array, must be an .xlsx file
     * @return a bidimensional ArrayList
     * @throws ExcelAPISyntaxException 
     */
    
    public static ArrayList<ArrayList<String>> transformToArray(String inputFile) throws ExcelAPISyntaxException {
        FileInputStream input = null;
        XSSFWorkbook wb = null;
        
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
            
        } catch (FileNotFoundException ex) {
            System.out.println("No se puede encontrar el archivo especificado.");
            
        } catch (IOException ex) {
            System.out.println("No se puede leer o escribir el archivo.");
            
        } finally {
            try {
                wb.close();
                input.close();
            } catch (IOException ex) {
                Logger.getLogger(Main.class.getName()).log(Level.SEVERE, null, ex);
            }
        }
        
        return data;
    }

    /**
     * 
     * This method transform the ArrayList introduced into a XLSX document
     * in the specified path.
     * 
     * @param data the ArrayList, must be a bidimensional ArrayList
     * @param path the path to save the returned document.
     * @throws ExcelAPISyntaxException 
     */
    
    public static void transformToExcel(ArrayList<ArrayList<String>> data, 
            String path) throws ExcelAPISyntaxException {
        
        ExcelBook e = new ExcelBook(data, path);
        e.writeExcelSheet();
        e.writeExcelFile();
      
    }
}
