import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Iterator;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

/*
@ author Sankalp Saxena
*/

public class save implements Runnable{
    
             public File f;
             public Sheet s;
             public int i;
            
             public save(File outputFile , Sheet sheet , int i) throws FileNotFoundException{
                        f = outputFile;
                        s = sheet;
                        this.i = i;
    
             }
    
    
 public void run(){
        // For storing data into CSV files
        StringBuffer data = new StringBuffer();
        try 
        {
                        try (FileOutputStream fos = new FileOutputStream(f)) {
                            Cell cell;
                            Row row;
                            Date date;
                           
                            // Iterate through each rows from first sheet
                            Iterator<Row> rowIterator = s.iterator();
                            while (rowIterator.hasNext())
                            {
                                    row = rowIterator.next();
               
                                    // For each row, iterate through each columns
                                    Iterator<Cell> cellIterator = row.cellIterator();
                   
                                    while (cellIterator.hasNext())
                                    {
                                     
                                            cell = cellIterator.next();
                                            switch (cell.getCellType())
                                            {
                                          
                                                    case Cell.CELL_TYPE_BOOLEAN:
                                                    data.append(cell.getBooleanCellValue() + ",");
                                                    break;
                                            
                                                    case Cell.CELL_TYPE_NUMERIC:
                                                     
                                                    if (DateUtil.isCellDateFormatted(cell))
                                                    {
                                                              SimpleDateFormat sdf = new SimpleDateFormat("dd/MM/yyyy");
                                                              data.append( sdf.format(cell.getDateCellValue())+ ",");
                                                    }
                                            
                                                    else
                                                              data.append(cell.getNumericCellValue() + ",");
                                                    break;
                                            
                                                    case Cell.CELL_TYPE_STRING:
                                                    data.append(cell.getStringCellValue() + ",");
                                                    break;
                                            
                                                    case Cell.CELL_TYPE_FORMULA:
                                                    data.append(cell.getNumericCellValue() + ",");
                                                    break;
                                        
                                                    case Cell.CELL_TYPE_BLANK:
                                                    data.append(cell.getStringCellValue()+ ", " );
                                                    break;
                                            
                                                    case Cell.CELL_TYPE_ERROR:
                                                    data.append(cell.getErrorCellValue() + ",");
                                                    break;
                                        
                                                    default:
                                                    data.append(cell + ",");
                                            }
                                     }
                                 data.append('\n');
                             }
                            
                    fos.write(data.toString().getBytes());
              }
        }
        catch (FileNotFoundException e) 
        {
                e.printStackTrace();
        }
        catch (IOException e) 
        {
                e.printStackTrace();
        }
                   
      }
    }
    
    

