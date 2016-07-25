package covReports;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.Iterator;
import java.util.Map;
import java.util.Set;

/**
 * Created by Nakate on 20-06-2016.
 */
public class WriteExcel {

    public static void writeTotals(String fpath,int rows){
        String path=fpath;
        try {
            System.out.println("Writing totals");
            File file = new File(path);
            FileInputStream fileInputStream = new FileInputStream(file);
            XSSFWorkbook inputWorkbook = new XSSFWorkbook(fileInputStream);
            XSSFSheet xssfSheet = inputWorkbook.getSheet("Statistics");
            FileOutputStream fileOutputStream = null;
            int feat=rows;
            rows++;
            xssfSheet.createRow(rows);
            xssfSheet.getRow(rows).createCell(0).setCellValue("Totals");

            for (int i=1;i<19;i++) {
                if(i==7 || i==8 ||i==12)
                {}
                else if (i==5){
                    String strFormula = "B"+(rows+1)+"/D"+(rows+1)+"*100";
                    xssfSheet.getRow(rows).createCell(i).setCellType(XSSFCell.CELL_TYPE_FORMULA);
                    xssfSheet.getRow(rows).createCell(i).setCellFormula(strFormula);
                }
                else if (i==6){
                    String strFormula = "C"+(rows+1)+"/D"+(rows+1)+"*100";
                    xssfSheet.getRow(rows).createCell(i).setCellType(XSSFCell.CELL_TYPE_FORMULA);
                    xssfSheet.getRow(rows).createCell(i).setCellFormula(strFormula);
                }else {
                    String col = getCharForNumber(i);
                    String strFormula = "SUM(" + col + "2:" + col + feat + ")";
                    xssfSheet.getRow(rows).createCell(i).setCellType(XSSFCell.CELL_TYPE_FORMULA);
                    xssfSheet.getRow(rows).createCell(i).setCellFormula(strFormula);
                }
            }

            try{
                fileOutputStream = new FileOutputStream(file);
                inputWorkbook.write(fileOutputStream);

                fileOutputStream.flush();
                fileOutputStream.close();
            }catch (FileNotFoundException e) {
                e.printStackTrace();
            }
            finally {
                fileInputStream.close();
            }

        }catch(IOException e){
            e.printStackTrace();
        }


    }

    private static String getCharForNumber(int i) {
        return i > -1 && i < 26 ? String.valueOf((char)(i + 65)) : null;
    }

    public static void deleteFile(File file){
        file.delete();
    }

}

