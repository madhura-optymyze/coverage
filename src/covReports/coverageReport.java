package covReports;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.util.HashMap;
import java.util.Map;

/**
 * Created by Madhura Nakate on 20-06-2016.
 */
public class coverageReport {
    public static String path="";
    public static void main(String[] args){
        path = "E:\\autocoverage\\"+args[0];
     //   WriteExcel.deleteFile(new File("E:\\autocoverage\\Stats.xlsx"));

        int totalRows=0;

        File file = new File(path);
        if (!file.exists())
            System.out.println("File does not exist!");
        else {
            System.out.println("File exists");
            try {
                FileInputStream fileInputStream = new FileInputStream(file);
                XSSFWorkbook inputWorkbook = new XSSFWorkbook(fileInputStream);
                for( Sheet sheet: inputWorkbook){
                    totalRows=new Calculate(path,sheet).generateCoverageData();
                }
                WriteExcel.writeTotals(path,totalRows);

            }catch (Exception e){
                e.printStackTrace();
            }

        }



    }
}
