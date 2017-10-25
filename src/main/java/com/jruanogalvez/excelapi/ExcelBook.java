package com.jruanogalvez.excelapi;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelBook {
    private XSSFWorkbook wb;
    private ArrayList<ArrayList<String>> data;
    private String path;
    private File output;
    private String inputFile;

    public ExcelBook(ArrayList<ArrayList<String>> data, String path) {
        this.data = data;
        this.path = path;
    }
    
    public ExcelBook(String inputFile) {
        this.inputFile = inputFile;
    }
    
    public boolean checkPathSyntax() throws ExcelAPISyntaxException {
        if(path.endsWith(".xlsx"))
            return true;
        else
            throw new ExcelAPISyntaxException("El nombre del archivo " + path +
                    " es incorrecto.");
    }
    
    public ArrayList<ArrayList<String>> readExcelFile() {
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
    
    public void writeExcelSheet() throws ExcelAPISyntaxException {
        if(checkPathSyntax()) {
            wb = new XSSFWorkbook();
            Sheet thisSheet = wb.createSheet();

            for(int i = 0; i < data.size(); i++) {
                Row thisRow = thisSheet.createRow(i);

                for(int j = 0; j < data.get(i).size(); j++) {
                    Cell thisCell = thisRow.createCell(j);
                    thisCell.setCellValue(data.get(i).get(j));
                }
            }
        }
    }
    
    public void writeExcelFile() throws ExcelAPISyntaxException {
        if(checkPathSyntax()) {
            
            try {
                output = new File(path);

                if(!output.exists())
                    output.createNewFile();
                
            } catch (IOException ex) {
                System.out.println("Error al crear el fichero.");
            }
            
            try(FileOutputStream out = new FileOutputStream(output);) {
                wb.write(out);
                
            } catch (IOException ex) {
                System.out.println("Error de Entrada/Salida.");

            } finally {
                try {
                    wb.close();
                } catch (IOException ex) {
                    Logger.getLogger(ExcelBook.class.getName()).log(Level.SEVERE, null, ex);
                }
            }
        }
    }
    
}
