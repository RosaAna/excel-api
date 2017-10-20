package com.jruanogalvez.excelapi;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelBook {
    private XSSFWorkbook wb;
    private ArrayList<ArrayList<String>> data;
    private String path;
    private File output;

    public ExcelBook(ArrayList<ArrayList<String>> data, String path) {
        this.data = data;
        this.path = path;
    }
    
    public void writeExcelSheet() {
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
    
    public void writeExcelFile() {
        try {
            output = new File(path);
            
            if(!output.exists())
                output.createNewFile();
            
            FileOutputStream out = new FileOutputStream(output);
            wb.write(out);
            
        } catch (FileNotFoundException ex) {
            System.out.println("No se puede encontrar el archivo específicado"
                    + " para la salida.");
        } catch (IOException ex) {
            System.out.println("Error de Entrada/Salida.");
        }
    }
    
}
