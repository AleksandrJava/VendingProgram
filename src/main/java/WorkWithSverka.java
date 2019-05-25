import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFSheet;

import static java.lang.Math.abs;

public class WorkWithSverka {
    public static void checkSverka(HSSFSheet sheet, HSSFSheet sverkaSheet, int a, Logger log){
        HSSFCell yesOrNoSverka = sverkaSheet.getRow(0).getCell(1);
        String yesOrNoSverkaStr = yesOrNoSverka.getStringCellValue();


        if (yesOrNoSverkaStr.equals("Да") || yesOrNoSverkaStr.equals("да") || yesOrNoSverkaStr.equals("ДА") || yesOrNoSverkaStr.equals("дА")) {
            vAvtomatePoComp(sheet, sverkaSheet, a);
            for (int i = 3; i < a + 1; i++) {
                HSSFCell errorToNull = sverkaSheet.getRow(i).getCell(5);
                errorToNull.setCellValue("нет");

                HSSFCell vAvtomate = sverkaSheet.getRow(i).getCell(2);
                HSSFCell poComp = sverkaSheet.getRow(i).getCell(4);
                HSSFCell spisanie = sverkaSheet.getRow(i).getCell(6);
                HSSFCell name = sverkaSheet.getRow(i).getCell(0);
                HSSFCell cellNumber = sverkaSheet.getRow(i).getCell(1);

                double doubVavtom = vAvtomate.getNumericCellValue();
                double doubPoComp = poComp.getNumericCellValue();
                double doubSpisanie = spisanie.getNumericCellValue();
                double doubCellName = cellNumber.getNumericCellValue();

                HSSFCell spis = sheet.getRow(i - 1).getCell(28);
                spis.setCellValue(doubSpisanie);

                String strName = name.getStringCellValue();

                double minus = doubVavtom - doubPoComp;
                if (abs(minus) == abs(doubSpisanie)) {
                    minus = 0;
                }

                HSSFCell sver = sheet.getRow(i - 1).getCell(27);
                sver.setCellValue(minus);

                HSSFCell cell = sheet.getRow(i - 1).getCell(25);

                if (minus == 0) {
                    cell.setCellValue(cell.getNumericCellValue() - doubSpisanie);
                    System.out.println("Всё правильно!");
                } else if (minus > 0) {
                    log.error("Owibka v sverke - " + doubCellName + ". V avtomate bol'we tovara na " + minus);
                    System.out.println("Ошибка в сверке. В автомате больше товара на " + minus);

                    cell.setCellValue(cell.getNumericCellValue() + minus - doubSpisanie);

                    errorToNull.setCellValue("Ошибка");

                } else {
                    log.error("Owibka v sverke - " + doubCellName + ". V avtomate men'we tovara na " + abs(minus));
                    System.out.println("Ошибка в сверке. В автомате меньше товара на " + minus);
                    //HSSFCell cell = sheet.getRow(i - 1).getCell(25);
                    cell.setCellValue(cell.getNumericCellValue() + minus - doubSpisanie);

                    errorToNull.setCellValue("Ошибка");
                }


            }
        }

    }

    private static void vAvtomatePoComp(HSSFSheet sheet, HSSFSheet sverkaSheet, int a){
        for (int i = 3; i < a+1; i++) {
            HSSFCell itog = sverkaSheet.getRow(i).getCell(4);

            HSSFCell automatHaveBegin = sheet.getRow(i-1).getCell(19);

            HSSFCell saleCell = sverkaSheet.getRow(i).getCell(3);

            itog.setCellValue(0);

            double minus = abs(automatHaveBegin.getNumericCellValue() - saleCell.getNumericCellValue());
            itog.setCellValue(minus);
        }
    }

}
