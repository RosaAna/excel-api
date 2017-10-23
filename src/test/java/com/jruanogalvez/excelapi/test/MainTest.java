package com.jruanogalvez.excelapi.test;

import com.jruanogalvez.excelapi.ExcelAPISyntaxException;
import com.jruanogalvez.excelapi.ExcelBook;
import com.jruanogalvez.excelapi.Main;
import java.util.ArrayList;
import org.junit.After;
import org.junit.AfterClass;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;

/**
 *
 * @author Jose Ruano
 */
public class MainTest {
    
    public MainTest() {
    }
    
    @BeforeClass
    public static void setUpClass() {
    }
    
    @AfterClass
    public static void tearDownClass() {
    }
    
    @Before
    public void setUp() {
    }
    
    @After
    public void tearDown() {
    }
    
    ArrayList<ArrayList<String>> lista = new ArrayList<>();
    ArrayList<String> listaSec = new ArrayList<>();
    
    @Test
    public void testSyntaxCheck() throws ExcelAPISyntaxException {
        ArrayList<ArrayList<String>> lista = null;
        ExcelBook e = new ExcelBook(lista, "asfasfasjfsñl0987654321-+ç´`<>.xlsx");
        
        e.checkPathSyntax();
    }
    
    @Test
    public void testExcelSheetWritting() throws ExcelAPISyntaxException {
        listaSec.add("Hola");
        listaSec.add("Adios");
        lista.add(listaSec);
                
        ExcelBook e = new ExcelBook(lista, "asfasfasjfsñl0987654321-+ç´`<>.xlsx");
        e.writeExcelSheet();
    }
    
    @Test
    public void testExcelFileWritting() throws ExcelAPISyntaxException {
        ExcelBook e = new ExcelBook(lista, "asfasfasjfsñl0987654321-+ç´`<>.xlsx");
        
        e.writeExcelSheet();
        e.writeExcelFile();
    }
    
    @Test
    public void testAllTransform() throws ExcelAPISyntaxException {
        Main.transformToExcel(Main.transformToArray("C:\\Users\\matinal\\Documents\\NetBeansProj"
                + "ects\\ExcelAPI\\src\\main\\resources\\testFiles"
                + "\\example.xlsx"), "C:\\Users\\matinal\\Documents\\NetBeansProj"
                + "ects\\ExcelAPI\\src\\main\\resources\\testFiles"
                + "\\outputExample.xlsx");
    }
}
