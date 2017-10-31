package com.jruanogalvez.excelapi;

import java.util.ArrayList;

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
        ExcelBook e = new ExcelBook(inputFile);
        
        return e.readExcelFile();
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
