package excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook; 
import org.apache.poi.ss.usermodel.Cell; 
 
/** 
 * Utility class, where we will create methods for training read and write excel files,
 * with <a href="https://poi.apache.org/">Apache POI</a>, we use 
 * <a href="https://poi.apache.org/spreadsheet/">POI-HSSF and POI-XSSF - Java API To Access Microsoft</a>
 * HSSF is the POI Project's pure Java implementation of the Excel '97(-2007) file.
 * 
 * Clase de utilidades, donde crearemos m�todos
 * para el aprendizaje de la lectura y escritura de ficheros excel con 
 * <a href="https://poi.apache.org/">Apache POI</a>, usaremos
 * <a href="https://poi.apache.org/spreadsheet/">POI-HSSF and POI-XSSF - Java API To Access Microsoft</a>
 * HSSF es el proyecto POI de implementaci�n total en Java para ficheros Excel '97(-2007).
 *
 * @author Xules You can follow me on my website http://www.codigoxules.org/en
 * Puedes seguirme en mi web http://www.codigoxules.org).
 */

public class ImportarExcel {

	/**
     * Explanation of the method by which we read the excel file we pass as
     * parameter if exists, in this example we print the content in the console.
     * Explicaci�n del m�todo con el que leemos el fichero excel que pasamos como
     * par�metro si existe, en este ejemplo mostramos el contenido por la consola.
     * <h3>Example (Ejemplo)</h3>
     * <pre>
     * JavaPoiUtils javaPoiUtils = new JavaPoiUtils();
     * javaPoiUtils.readExcelFile(new File("/home/xules/codigoxules/apachepoi/PaisesIdiomasMonedas.xls"));    
     * </pre>
     *
     * @param excelFile <code>String</code> 
     *      excel File we are going to read. 
     *      Fichero excel que vamos a leer. 
     */
    public void readExcelFile(File excelFile){
        InputStream excelStream = null;
        try {
            excelStream = new FileInputStream(excelFile);
            // High level representation of a workbook.
            // Representaci�n del m�s alto nivel de la hoja excel.
            HSSFWorkbook hssfWorkbook = new HSSFWorkbook(excelStream);
            // We chose the sheet is passed as parameter. 
            // Elegimos la hoja que se pasa por par�metro.
            HSSFSheet hssfSheet = hssfWorkbook.getSheetAt(0);
            // An object that allows us to read a row of the excel sheet, and extract from it the cell contents.
            // Objeto que nos permite leer un fila de la hoja excel, y de aqu� extraer el contenido de las celdas.
            HSSFRow hssfRow;
            // Initialize the object to read the value of the cell 
            // Inicializo el objeto que leer� el valor de la celda
            HSSFCell cell;                        
            // I get the number of rows occupied on the sheet
            // Obtengo el n�mero de filas ocupadas en la hoja
            int rows = hssfSheet.getLastRowNum();
            // I get the number of columns occupied on the sheet
            // Obtengo el n�mero de columnas ocupadas en la hoja
            int cols = 0;            
            // A string used to store the reading cell
            // Cadena que usamos para almacenar la lectura de la celda
            String cellValue;  
            // For this example we'll loop through the rows getting the data we want
            // Para este ejemplo vamos a recorrer las filas obteniendo los datos que queremos            
            for (int r = 0; r < rows; r++) {
                hssfRow = hssfSheet.getRow(r);
                if (hssfRow == null){
                    break;
                }else{
                    System.out.print("Row: " + r + " -> ");
                    for (int c = 0; c < (cols = hssfRow.getLastCellNum()); c++) {
                        /* 
                            We have those cell types (tenemos estos tipos de celda): 
                                CELL_TYPE_BLANK, CELL_TYPE_NUMERIC, CELL_TYPE_BLANK, CELL_TYPE_FORMULA, CELL_TYPE_BOOLEAN, CELL_TYPE_ERROR
                        */
                        cellValue = hssfRow.getCell(c) == null?"":
                                (hssfRow.getCell(c).getCellType() == Cell.CELL_TYPE_STRING)?hssfRow.getCell(c).getStringCellValue():
                                (hssfRow.getCell(c).getCellType() == Cell.CELL_TYPE_NUMERIC)?"" + hssfRow.getCell(c).getNumericCellValue():
                                (hssfRow.getCell(c).getCellType() == Cell.CELL_TYPE_BOOLEAN)?"" + hssfRow.getCell(c).getBooleanCellValue():
                                (hssfRow.getCell(c).getCellType() == Cell.CELL_TYPE_BLANK)?"BLANK":
                                (hssfRow.getCell(c).getCellType() == Cell.CELL_TYPE_FORMULA)?"FORMULA":
                                (hssfRow.getCell(c).getCellType() == Cell.CELL_TYPE_ERROR)?"ERROR":"";                       
                        System.out.print("[Column " + c + ": " + cellValue + "] ");
                    }
                    System.out.println();
                }
            }            
        } catch (FileNotFoundException fileNotFoundException) {
            System.out.println("The file not exists (No se encontr� el fichero): " + fileNotFoundException);
        } catch (IOException ex) {
            System.out.println("Error in file procesing (Error al procesar el fichero): " + ex);
        } finally {
            try {
                excelStream.close();
            } catch (IOException ex) {
                System.out.println("Error in file processing after close it (Error al procesar el fichero despu�s de cerrarlo): " + ex);
            }
        }
    }
    /**     
     * Main method for the tests for the methods of the class <strong>Java
     * read excel</strong> and <strong>Java create excel</strong> 
     * with <a href="https://poi.apache.org/">Apache POI</a>. 
     * <br />
     * M�todo main para las pruebas para los m�todo de la clase,
     * pruebas de <strong>Java leer excel</strong> y  <strong>Java crear excel</strong>
     * con <a href="https://poi.apache.org/">Apache POI</a>.     
     * @param args 
     */
    public static void main(String[] args){
        ImportarExcel javaPoiUtils = new ImportarExcel();
        javaPoiUtils.readExcelFile(new File("C:\\Users\\formacion\\Desktop\\WorkSpace\\Excel\\AlumnosAytos.xls "));       
    } 
	
}
