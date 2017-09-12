
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.Scanner;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.TimeUnit;
import org.apache.commons.io.FilenameUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

/**
 * Language Java
 * @author Sankalp Saxena
 */
public class convert {

    public static void main(String [ ] args) throws InterruptedException {
        
        InputStream inp = null;
           try {
               
            int i;
            String inputPath;   
            String outputPath;
            String csvPath;
            String withExtension;
            String workbookName;
            
            
            Scanner sc = new Scanner (System.in);
            System.out.println("Please enter the Excel file path ");
            inputPath = sc.nextLine();
            
            inp = new FileInputStream(inputPath);
            Workbook wb = WorkbookFactory.create(inp);
     
            System.out.println("Please provide the path where you want your CSV File ");
            csvPath = sc.nextLine();
            
            File f = new File(inputPath);       // CLASS HAVING getName() TO GET WORKBOOK NAME 
            withExtension = f.getName();  // GET THE NAME OF WORKBOOK WITH EXTENSION
            workbookName = FilenameUtils.removeExtension(withExtension);     //GET THE NAME OF WORKBOOK
            
          
            int sheetCount = wb.getNumberOfSheets();
            
            ExecutorService executor = Executors.newFixedThreadPool(sheetCount);
            
                       for( i=0 ; i<sheetCount ; i++) {
 
                                    String sheetName = wb.getSheetAt(i).getSheetName();
                
                                    outputPath = csvPath + workbookName + "_" + sheetName +".csv";
                                    File outputFile = new File(outputPath);
        
                                    executor.submit(new save (outputFile , wb.getSheetAt(i) , i ));
                      }
          
                System.out.println( i + " CSV Files created" );
                
              executor.shutdown();
              executor.awaitTermination(1, TimeUnit.DAYS);
        
            
        } catch (InvalidFormatException ex) {
                      System.out.println("Error : " + ex);
        } catch (FileNotFoundException ex) {
                      System.out.println("Error : FILE NOT FOUND No such file or directory exists! ");
        } catch (IOException ex) {
                      System.out.println("Error : " + ex);
        } finally {
            try {
                            if(inp != null)
                            inp.close();
            } catch (IOException ex) {
                      System.out.println("Error : " + ex);
            }
        }
    }
}
