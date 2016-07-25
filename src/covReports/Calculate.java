package covReports;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.Iterator;

/**
 * Created by Madhura Nakate on 20-06-2016.
 */
public class Calculate {
    private int[] P1 = {0, 0};
    private int[] P2 = {0, 0};
    private int[] P3 = {0, 0};
    private int totalGoals = 0, coveredGoals = 0, uncoveredGoals = 0, fniGoals = 0, manualGoals = 0;
    static int iterator = 0;
    double perc=0, autoPerc=0;
    String goalNum="";
    String sheetName="";
    String path = "";
    DataFormatter formatter = new DataFormatter();
    Sheet sheet;
    public Calculate(String fpath,Sheet wsheet){
        this.sheet=wsheet;
        this.path=fpath;
    }

    public int generateCoverageData() {
        sheetName = sheet.getSheetName();
        Iterator<Row> iter = sheet.iterator();
        if(iter.hasNext())
            iter.next();

        while (iter.hasNext()) {
            Row nextRow = iter.next();
            Cell cell0 = nextRow.getCell(0);
            if (cell0 != null && (formatter.formatCellValue(cell0)) != "") {
                totalGoals++;
               goalNum = formatter.formatCellValue(cell0);

                Cell cell3 = nextRow.getCell(3);
                if (cell3 == null||formatter.formatCellValue(cell3)==""||!(formatter.formatCellValue(cell3).equalsIgnoreCase("manual"))) {
                    Cell cell2 = nextRow.getCell(2);
                    if (cell2 == null||formatter.formatCellValue(cell2)=="") {
                        uncoveredGoals++;
                        priorityCount(nextRow, 1);
                    } else if (formatter.formatCellValue(cell2).equalsIgnoreCase("FNI"))
                        fniGoals++;
                    else {
                        coveredGoals++;
                        priorityCount(nextRow, 0);
                    }
                }else if (formatter.formatCellValue(cell3).equalsIgnoreCase("manual")){
                    manualGoals++;
                    priorityCount(nextRow, 1);
                }


            }
        }
        perc = ((double)coveredGoals/(double)(totalGoals-fniGoals))*100;
        autoPerc = (((double)(totalGoals-fniGoals-manualGoals))/(double)(totalGoals-fniGoals))*100;
//        System.out.println("Sheet name: " + sheetName + "\nTotal goals: " + totalGoals + "\tCovered goals: " + coveredGoals + "\tUncovered goals: " + uncoveredGoals + "\tFNI goals: " + fniGoals+"\tPercentage: "+perc);
//        System.out.println("P1 covered: " + P1[0] + "\tP1 uncovered: " + P1[1] + "\nP2 covered: " + P2[0] + "\tP2 uncovered: " + P2[1] + "\nP3 covered: " + P3[0] + "\tP3 uncovered: " + P3[1]);
            writeResults();
        return iterator;
    }

    private void priorityCount(Row currentRow, int type) {
        Cell cell1 = currentRow.getCell(1);
        if (cell1 != null) {
            if (formatter.formatCellValue(cell1)=="")
                System.out.println("Priority missing for goal "+goalNum+" in sheet "+sheetName);
            else {
                switch (formatter.formatCellValue(cell1).charAt(1)) {
                    case '1':
                        P1[type]++;
                        break;
                    case '2':
                        P2[type]++;
                        break;
                    case '3':
                        P3[type]++;
                        break;
                }
            }
        }
    }

    private void writeResults() {
        try {
            File file = new File(path);
//            if (!file.exists())
            if (iterator==0)
                createStatsFile(file);
            writeSheetData(file);
        }catch(IOException e){
            e.printStackTrace();
        }
    }

    private void createStatsFile(File file) throws IOException {

//        try {

            FileInputStream fileInputStream = new FileInputStream(file);
            XSSFWorkbook inputWorkbook = new XSSFWorkbook(fileInputStream);
            XSSFSheet xssfSheet = inputWorkbook.createSheet("Statistics");
            FileOutputStream fileOutputStream = null;

//            FileOutputStream fileOutputStream = new FileOutputStream(file);
//            XSSFWorkbook inputWorkbook = new XSSFWorkbook();
//            XSSFSheet xssfSheet = inputWorkbook.createSheet("Statistics");
            xssfSheet.createRow(iterator);
            xssfSheet.getRow(iterator).createCell(0).setCellValue("Feature");
            xssfSheet.getRow(iterator).createCell(1).setCellValue("Covered goals");
            xssfSheet.getRow(iterator).createCell(2).setCellValue("Automatable goals");
            xssfSheet.getRow(iterator).createCell(3).setCellValue("Implemented goals");
            xssfSheet.getRow(iterator).createCell(4).setCellValue("Total goals");
            xssfSheet.getRow(iterator).createCell(5).setCellValue("Coverage percentage");
            xssfSheet.getRow(iterator).createCell(6).setCellValue("Automatable percentage");

            xssfSheet.getRow(iterator).createCell(8).setCellValue("Feature");
            xssfSheet.getRow(iterator).createCell(9).setCellValue("P1");
            xssfSheet.getRow(iterator).createCell(10).setCellValue("P2");
            xssfSheet.getRow(iterator).createCell(11).setCellValue("P3");

            xssfSheet.getRow(iterator).createCell(13).setCellValue("P1 covered");
            xssfSheet.getRow(iterator).createCell(14).setCellValue("P1 uncovered");
            xssfSheet.getRow(iterator).createCell(15).setCellValue("P2 covered");
            xssfSheet.getRow(iterator).createCell(16).setCellValue("P2 uncovered");
            xssfSheet.getRow(iterator).createCell(17).setCellValue("P3 covered");
            xssfSheet.getRow(iterator).createCell(18).setCellValue("P3 uncovered");

            iterator++;

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

    }

    private void writeSheetData(File file) throws IOException {
        FileInputStream fileInputStream = new FileInputStream(file);
        XSSFWorkbook inputWorkbook = new XSSFWorkbook(fileInputStream);
        XSSFSheet xssfSheet = inputWorkbook.getSheet("Statistics");
        FileOutputStream fileOutputStream = null;
        String sheetName = sheet.getSheetName();

        xssfSheet.createRow(iterator);
        xssfSheet.getRow(iterator).createCell(0).setCellValue(sheetName);
        xssfSheet.getRow(iterator).createCell(1).setCellValue(coveredGoals);
        xssfSheet.getRow(iterator).createCell(2).setCellValue(totalGoals-fniGoals-manualGoals);
        xssfSheet.getRow(iterator).createCell(3).setCellValue(totalGoals-fniGoals);
        xssfSheet.getRow(iterator).createCell(4).setCellValue(totalGoals);
        if((totalGoals-fniGoals)!=0)
            xssfSheet.getRow(iterator).createCell(5).setCellValue(perc);
        else
            xssfSheet.getRow(iterator).createCell(5).setCellValue("0");
        xssfSheet.getRow(iterator).createCell(6).setCellValue(autoPerc);

        xssfSheet.getRow(iterator).createCell(8).setCellValue(sheetName);
        xssfSheet.getRow(iterator).createCell(9).setCellValue(P1[0]+P1[1]);
        xssfSheet.getRow(iterator).createCell(10).setCellValue(P2[0]+P2[1]);
        xssfSheet.getRow(iterator).createCell(11).setCellValue(P3[0]+P3[1]);

        xssfSheet.getRow(iterator).createCell(13).setCellValue(P1[0]);
        xssfSheet.getRow(iterator).createCell(14).setCellValue(P1[1]);
        xssfSheet.getRow(iterator).createCell(15).setCellValue(P2[0]);
        xssfSheet.getRow(iterator).createCell(16).setCellValue(P2[1]);
        xssfSheet.getRow(iterator).createCell(17).setCellValue(P3[0]);
        xssfSheet.getRow(iterator).createCell(18).setCellValue(P3[1]);

        iterator++;

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
    }

}





