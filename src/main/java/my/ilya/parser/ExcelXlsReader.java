package my.ilya.parser;

import java.io.File;
import java.io.FileInputStream;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.Iterator;
import java.util.List;

import javax.swing.text.SimpleAttributeSet;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelXlsReader 
{
	static String SQL_STRING = "INSERT INTO ZME_NCCODES (FAILURE_MODE, TRANSLATED_TEXT, ERDAT, ERZET, ERNAM) VALUES ('?','?',TO_DATE('?', 'mm/dd/yyyy hh24:mi:ss'),TO_DATE('?', 'mm/dd/yyyy hh24:mi:ss'),'?');";
	static SimpleDateFormat df = new SimpleDateFormat();
	
    public static void main( String[] args )
    {
    	try
        {
    		FileInputStream file = new FileInputStream(new File("ZME_NCCODES.XLS"));
            //Create Workbook instance holding reference to .xlsx file
            HSSFWorkbook  workbook =new HSSFWorkbook(file);
            
            //Get first/desired sheet from the workbook
            HSSFSheet  sheet = workbook.getSheetAt(0);
 
            //Iterate through each rows one by one
            
            Iterator<Row> rowIterator = sheet.iterator();
            int i = 0;
            while (rowIterator.hasNext())
            {	i++;
                Row row = rowIterator.next();
                List<String> params = new ArrayList<String>();
                String sss = null;
                //For each row, iterate through all the columns
                Iterator<Cell> cellIterator = row.cellIterator();
                
                while (cellIterator.hasNext() )                	
                {
                	
                    Cell cell = cellIterator.next();
                    //Check the cell type and format accordingly
                    
                    switch (cell.getCellType())
                    {
                        case Cell.CELL_TYPE_NUMERIC:
                            //System.out.print(cell.getDateCellValue() + "     ");
                        	df.applyPattern("MM/dd/yyyy HH:mm:ss");                               
                        	params.add(df.format(cell.getDateCellValue()));
                        	//System.out.print(df.format(cell.getDateCellValue()) + "     ");
                            
                            break;
                        case Cell.CELL_TYPE_STRING:
                            //System.out.print(cell.getStringCellValue() + "      ");
                        	params.add(cell.getStringCellValue());
                            break;                       
                    }
                    
                }
                sss = String.format(SQL_STRING.replace("?", "%s"), params.toArray());
                System.out.println(sss);
            }
            file.close();            
            System.out.println(i);
        }
        catch (Exception e)
        {
            e.printStackTrace();
        }
    }
}
