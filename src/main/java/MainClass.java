import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.Date;

public class MainClass{
    private static final Logger log = Logger.getLogger(MainClass.class);
    static final int a = 45; // Number of items with goods
    public static void main(String[] args) throws IOException {
        FileInputStream oldInputStream = null;
        try {
            oldInputStream = new FileInputStream(new File("E:/1Vending-proga/dispense.xls"));
        } catch (FileNotFoundException e) {
            log.error("File 'dispense.xls' not found");
            e.printStackTrace();
        }
        HSSFWorkbook oldWorkbook = new HSSFWorkbook(oldInputStream);
        HSSFSheet oldSheet = oldWorkbook.getSheetAt(0);

        File file = null;
        FileInputStream inputStream = null;
        try {
            file = new File("E:/1Vending-proga/sclad.xls");
            inputStream = new FileInputStream(new File("E:/1Vending-proga/sclad.xls"));
        } catch (FileNotFoundException e) {
            log.error("File 'sclad.xls' not found");
            e.printStackTrace();
        }
        HSSFWorkbook workbook = new HSSFWorkbook(inputStream);

        SimpleDateFormat dateFormat = new SimpleDateFormat("dd");
        Date dateDay = new Date();
        String today = dateFormat.format(dateDay);
        //String today = "03"; //If I need to manually correct the dates

        addLastDay(workbook, today);

        int numberOfSheets = workbook.getNumberOfSheets();

        int i = Integer.parseInt(today);

        HSSFSheet sheet = workbook.getSheetAt(numberOfSheets-1);
        HSSFSheet lastSheet = workbook.getSheetAt(numberOfSheets-2);

        File fileZayavka = null;
        FileInputStream zayavkaInputStream = null;
        try {
            fileZayavka = new File("E:/1Vending-proga/zayavkaItog.xls");
            zayavkaInputStream = new FileInputStream(fileZayavka);
        } catch (FileNotFoundException e) {
            log.error("File 'zayavkaItog.xls' not found");
            e.printStackTrace();
        }
        HSSFWorkbook zayavkaWorkbook = new HSSFWorkbook(zayavkaInputStream);
        HSSFSheet zayavkaSheet = zayavkaWorkbook.getSheetAt(0);
        HSSFSheet zayavkaBigSheet = zayavkaWorkbook.getSheetAt(1);
        HSSFSheet sverkaSheet = zayavkaWorkbook.getSheetAt(2);

        dataChangeAll(sheet, zayavkaSheet, zayavkaBigSheet, sverkaSheet, i, dateDay);

        WorkWithYestarday.transportScladFromYesterday(sheet, lastSheet, a);

        WorkWithZayavka.buyAndMovedToAutomat(zayavkaSheet, sheet, a);

        WorkWithSclad.sclad(sheet, a, log);

        WorkWithYestarday.transportAutomatBeginDay(lastSheet, sheet, a);

        WorkWithSclad.automatWithNewProduct(sheet, a);

        WorkWithSclad.saleFromToday(oldSheet, sheet, sverkaSheet, lastSheet, a, log);
        oldInputStream.close();

        WorkWithSclad.transportAutomatEndDay(sheet, a);

        WorkWithSverka.checkSverka(sheet, sverkaSheet, a, log);

        WorkWithZayavka.writeBigRequest(sheet, zayavkaBigSheet, a);

        WorkWithZayavka.writeShortRequest(zayavkaBigSheet, zayavkaSheet, a, log);

        WorkWithSclad.writeAllPriceCells(sheet, a);

        WorkWithSclad.writeAllItog(sheet, a);

        WorkWithSclad.itogWriteCells(sheet, today, a);


        zayavkaInputStream.close();

        FileOutputStream outZayavka = new FileOutputStream(fileZayavka);
        zayavkaWorkbook.write(outZayavka);
        outZayavka.close();

        inputStream.close();

        FileOutputStream out = new FileOutputStream(file);
        workbook.write(out);
        out.close();

        log.info(" ");
        log.info(" ");
    }

    private static void addLastDay(HSSFWorkbook workbook, String today){
        int numberOfSheets = workbook.getNumberOfSheets();

        for (int i = 0; i < numberOfSheets; i++) {
            String sheetName = workbook.getSheetName(i);

            if(sheetName.equals(today)){
                int intSheet = Integer.parseInt(sheetName);
                System.out.println("Найден лист номер" + intSheet);
                if(intSheet == 0) {
                    workbook.removeSheetAt(i);
                    numberOfSheets = numberOfSheets - 1;
                    break;
                } else {
                    for (int j = i; j >=0 ; j--) {
                        System.out.println("Удален лист номер " + j);
                        workbook.removeSheetAt(j);
                        numberOfSheets = numberOfSheets - 1;
                    }
                    break;
                }
            }
        }
        workbook.cloneSheet(numberOfSheets-1);
        workbook.setSheetName(numberOfSheets, today);
    }

    private static void dataChangeAll(HSSFSheet sheet, HSSFSheet zayavkaSheet, HSSFSheet zayavkaBigSheet, HSSFSheet sverkaSheet, int i, Date date){
        SimpleDateFormat dateFormat = new SimpleDateFormat("MM.yyyy");
        String newB = Integer.toString(i);
        String newBB = Integer.toString(i+1);

        String myDate = dateFormat.format(date);
        String today = newB + "." + myDate;
        String tomorrow = newBB + "." + myDate;

        log.info("Otchet ot " + today);

        HSSFCell cell1 = sheet.getRow(0).getCell(0);
        cell1.setCellValue(today);

        HSSFCell cell2 = zayavkaSheet.getRow(1).getCell(0);
        cell2.setCellValue(tomorrow);

        HSSFCell cell3 = zayavkaBigSheet.getRow(1).getCell(0);
        cell3.setCellValue(tomorrow);

        HSSFCell cell4 = sverkaSheet.getRow(1).getCell(0);
        cell4.setCellValue(today);
    }
}
