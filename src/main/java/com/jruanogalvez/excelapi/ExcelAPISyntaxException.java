package com.jruanogalvez.excelapi;

public class ExcelAPISyntaxException extends Exception{
    
    public ExcelAPISyntaxException() {
        
    }
    
    /**
     * Exception shown when the syntax of the introduced path is not valid.
     * 
     * @param msg error message from the exception.
     */
    public ExcelAPISyntaxException(String msg) {
        super("ExcelAPISyntaxException" + msg);
    }
}
