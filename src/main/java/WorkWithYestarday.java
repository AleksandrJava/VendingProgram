import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFSheet;

public class WorkWithYestarday {
    public static void transportScladFromYesterday(HSSFSheet sheet, HSSFSheet lastSheet, int a){
        for (int i = 13; i <=14 ; i++) {
            for (int j = 2; j < a; j++) {
                HSSFCell bilo = lastSheet.getRow(j).getCell(i);
                double biloDoub = bilo.getNumericCellValue();
                HSSFCell stalo = sheet.getRow(j).getCell(i-11);
                stalo.setCellValue(biloDoub);
            }
        }
    }

    public static void transportAutomatBeginDay(HSSFSheet oldSheet, HSSFSheet newSheet, int a){
        for (int i = 25; i <=26 ; i++) {
            for (int j = 2; j <= a; j++) {
                HSSFCell oldCell = oldSheet.getRow(j).getCell(i);
                HSSFCell newCell = newSheet.getRow(j).getCell(i-8);
                newCell.setCellValue(oldCell.getNumericCellValue());
            }
        }
    }
}
