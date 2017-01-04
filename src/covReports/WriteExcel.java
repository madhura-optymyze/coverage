package covReports;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.xssf.usermodel.*;

import java.io.*;

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
                String strFormula;
                switch(i) {
                    case 4:
                    case 7:
                    case 8: break;

                    case 5:
                        strFormula = "B" + (rows + 1) + "/D" + (rows + 1) + "*100";
                        xssfSheet.getRow(rows).createCell(i).setCellType(XSSFCell.CELL_TYPE_FORMULA);
                        xssfSheet.getRow(rows).createCell(i).setCellFormula(strFormula);
                        break;

                    case 6:
                        strFormula = "C" + (rows + 1) + "/D" + (rows + 1) + "*100";
                        xssfSheet.getRow(rows).createCell(i).setCellType(XSSFCell.CELL_TYPE_FORMULA);
                        xssfSheet.getRow(rows).createCell(i).setCellFormula(strFormula);
                        break;

                    case 12:
                        strFormula = "J" + (rows + 1) + "+K" + (rows + 1) + "+L" + (rows + 1);
                        xssfSheet.getRow(rows).createCell(i).setCellType(XSSFCell.CELL_TYPE_FORMULA);
                        xssfSheet.getRow(rows).createCell(i).setCellFormula(strFormula);
                        break;

                     default:
                        String col = getCharForNumber(i);
                        strFormula = "SUM(" + col + "2:" + col + feat + ")";
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

    public static void highlighter(String fpath,int rows){

        String path=fpath;
        try {
        //    System.out.println("Highlighting");
            File file = new File(path);
            FileInputStream fileInputStream = new FileInputStream(file);
            XSSFWorkbook inputWorkbook = new XSSFWorkbook(fileInputStream);
            XSSFSheet xssfSheet = inputWorkbook.getSheet("Statistics");
            FileOutputStream fileOutputStream = null;
            rows++;

            double fTot = xssfSheet.getRow(rows).getCell(3).getNumericCellValue();
            double pTot = xssfSheet.getRow(rows).getCell(12).getNumericCellValue();
            if (fTot != pTot) {
                XSSFCellStyle style = inputWorkbook.createCellStyle();
                style.setFillForegroundColor(HSSFColor.LIME.index);
                style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);

                XSSFFont font = inputWorkbook.createFont();
                font.setColor(HSSFColor.RED.index);
                style.setFont(font);

                xssfSheet.getRow(rows + 1).getCell(12).setCellStyle(style);
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



