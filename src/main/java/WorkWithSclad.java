import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFSheet;


import java.util.Date;

import static java.lang.Math.abs;

public class WorkWithSclad {
    public static void sclad(HSSFSheet sheet, int a, Logger log){
        for (int i = 13; i <=14 ; i++) {
            for (int j = 2; j < a; j++) {
                HSSFCell scladBegin = sheet.getRow(j).getCell(i-11);
                if(scladBegin.getNumericCellValue() != 0) {
                    HSSFCell itog = sheet.getRow(j).getCell(i);
                    HSSFCell transport = sheet.getRow(j).getCell(i - 2);
                    itog.setCellValue(abs(scladBegin.getNumericCellValue() - transport.getNumericCellValue()));
                    if (itog.getNumericCellValue() == 0) {
                        HSSFCell cell = sheet.getRow(j).getCell(1);
                        log.info("Zakonchilsya segodnya " + cell.getNumericCellValue() +". Na sklade teper' pusto");
                    }
                }
            }
        }
    }

    public static void automatWithNewProduct(HSSFSheet sheet, int a){
        for (int i = 19; i <=20 ; i++) {
            for (int j = 2; j < a; j++) {
                HSSFCell itog = sheet.getRow(j).getCell(i);
                HSSFCell transport = sheet.getRow(j).getCell(i-8);
                HSSFCell automatBegin = sheet.getRow(j).getCell(i-2);
                itog.setCellValue(automatBegin.getNumericCellValue() + transport.getNumericCellValue());
            }
        }
    }

    public static void saleFromToday(HSSFSheet sheet,HSSFSheet bigtable, HSSFSheet sverkaSheet, HSSFSheet lastSheet, int a, Logger log){
        for (int i = 21; i <= 21 ; i++) {
            for (int j = 2; j < a; j++) {
                HSSFCell cell = bigtable.getRow(j).getCell(i);
                cell.setCellValue(0);
            }
        }

        for (int i = 1; i <= 1; i++) {
            for (int j = 3; j < a+1; j++) {
                HSSFCell saleIntoSverka = sverkaSheet.getRow(j).getCell(3);
                saleIntoSverka.setCellValue(0);
            }
        }

        try {
            HSSFCell cellDate = lastSheet.getRow(1).getCell(0);
            Date dateYesterday = cellDate.getDateCellValue();


            HSSFCell dateOfSverka = sverkaSheet.getRow(0).getCell(2);
            Date dateSverka = dateOfSverka.getDateCellValue();


            for (int i = 1; i <= 1 ; i++) {
                for (int j = 1; j < 150; j++) {
                    HSSFCell cell = sheet.getRow(j).getCell(i);

                    Date gg = cell.getDateCellValue();


                    if (gg.after(dateYesterday)) {
                        if (yesOrNo(sheet, j) == true) {
                            double cellNumber = sheet.getRow(j).getCell(2).getNumericCellValue();
                            addSale(bigtable, cellNumber, gg, dateSverka, sverkaSheet, a);
                        } else {
                            log.error("Owibka v prodaje " + sheet.getRow(j).getCell(2).getNumericCellValue() + " " + gg);
                        }
                    } else {
                        break;
                    }
                }
            }

            HSSFCell newCellDate = sheet.getRow(1).getCell(1);
            Date newDate = newCellDate.getDateCellValue();
            System.out.println("timeDisp " +newDate);
            newDate.setSeconds(newDate.getSeconds() + 10);
            HSSFCell newNewDate = bigtable.getRow(1).getCell(0);
            System.out.println("timeScl " + newNewDate);
            newNewDate.setCellValue(newDate);

        } catch (NullPointerException e){
            log.error("Owibka s yacheikoi A2 v tablice 'Sclad' ili s yacheikoi C1 v tablice 'zayavkaItog - 3 list");
            throw new NullPointerException();
        }
    }

    private static boolean yesOrNo(HSSFSheet sheet, int j){
        for (int i = 3; i <= 3; i++) {
            HSSFCell cell = sheet.getRow(j).getCell(i);
            String checkCell = cell.getStringCellValue();
            if(checkCell.equals("Да")){
                return true;
            }
        }
        return false;
    }

    private static void addSale(HSSFSheet sheet, double cellNumber, Date gg, Date dateSverka, HSSFSheet sverkaSheet, int a){

        if(gg.before(dateSverka)) {

            for (int i = 1; i <= 1; i++) {
                for (int j = 2; j < a; j++) {
                    HSSFCell cell = sheet.getRow(j).getCell(i);

                    double doub = cell.getNumericCellValue();

                    if (doub == cellNumber) {
                        HSSFCell saleCell = sheet.getRow(j).getCell(21);
                        HSSFCell sverkaCell = sverkaSheet.getRow(j+1).getCell(3);

                        saleCell.setCellValue(saleCell.getNumericCellValue() + 1);
                        sverkaCell.setCellValue(sverkaCell.getNumericCellValue() + 1);
                        break;
                    }
                }
            }


        } else {
            for (int i = 1; i <= 1; i++) {
                for (int j = 2; j < a; j++) {
                    HSSFCell cell = sheet.getRow(j).getCell(i);

                    double doub = cell.getNumericCellValue();

                    if (doub == cellNumber) {
                        HSSFCell saleCell = sheet.getRow(j).getCell(21);
                        saleCell.setCellValue(saleCell.getNumericCellValue() + 1);
                        break;
                    }
                }
            }
        }
    }

    public static void transportAutomatEndDay(HSSFSheet sheet, int a){
        for (int i = 25; i <= 26 ; i++) {
            for (int j = 2; j < a ; j++) {
                HSSFCell before = sheet.getRow(j).getCell(i-6);
                HSSFCell after = sheet.getRow(j).getCell(i-4);
                HSSFCell cell = sheet.getRow(j).getCell(i);
                cell.setCellValue(abs(after.getNumericCellValue() - before.getNumericCellValue()));
            }
        }
    }

    public static void writeAllPriceCells(HSSFSheet sheet, int a){
        for (int i = 2; i < a; i++) {
            HSSFCell sumOpt = sheet.getRow(i).getCell(6);
            sumOpt.setCellValue(0);
            HSSFCell sumRozn = sheet.getRow(i).getCell(7);
            sumRozn.setCellValue(0);
            HSSFCell peremestVavtSum = sheet.getRow(i).getCell(12);
            peremestVavtSum.setCellValue(0);
            HSSFCell prodSumRozn = sheet.getRow(i).getCell(22);
            prodSumRozn.setCellValue(0);
            HSSFCell prodSumOpt = sheet.getRow(i).getCell(23);
            prodSumOpt.setCellValue(0);
            HSSFCell newPribil = sheet.getRow(i).getCell(24);
            newPribil.setCellValue(0);
            HSSFCell sumSpis = sheet.getRow(i).getCell(29);
            sumSpis.setCellValue(0);

        }

        for (int i = 2; i < a; i++) {
            HSSFCell roznica = sheet.getRow(i).getCell(15);
            double doubRoznica = roznica.getNumericCellValue();
            HSSFCell opt = sheet.getRow(i).getCell(5);
            double doubOpt = opt.getNumericCellValue();

            HSSFCell sumOpt = sheet.getRow(i).getCell(6);
            HSSFCell kupili = sheet.getRow(i).getCell(4);
            sumOpt.setCellValue(kupili.getNumericCellValue() * doubOpt);
            HSSFCell sumRozn = sheet.getRow(i).getCell(7);
            sumRozn.setCellValue(kupili.getNumericCellValue()*doubRoznica);

            HSSFCell pribil = sheet.getRow(i).getCell(10);
            pribil.setCellValue(sumRozn.getNumericCellValue() - sumOpt.getNumericCellValue());

            HSSFCell peremestVavtSum = sheet.getRow(i).getCell(12);
            HSSFCell peremVavt = sheet.getRow(i).getCell(11);
            peremestVavtSum.setCellValue(peremVavt.getNumericCellValue() * doubRoznica);

            HSSFCell proverka = sheet.getRow(i).getCell(31);
            if(proverka.getNumericCellValue() != 1) {
                HSSFCell scladEndSht = sheet.getRow(i).getCell(13);
                HSSFCell scladEndSum = sheet.getRow(i).getCell(14);
                HSSFCell scladBegin = sheet.getRow(i).getCell(2);
                HSSFCell scladBeginSum = sheet.getRow(i).getCell(3);
                scladEndSht.setCellValue(scladBegin.getNumericCellValue() - peremVavt.getNumericCellValue());
                scladEndSum.setCellValue(scladBeginSum.getNumericCellValue() - peremestVavtSum.getNumericCellValue());
            }

            HSSFCell prodaja = sheet.getRow(i).getCell(21);
            double doubProd = prodaja.getNumericCellValue();
            HSSFCell prodSumRozn = sheet.getRow(i).getCell(22);
            prodSumRozn.setCellValue(doubProd * doubRoznica);
            HSSFCell prodSumOpt = sheet.getRow(i).getCell(23);
            prodSumOpt.setCellValue(doubProd * doubOpt);
            HSSFCell newPribil =  sheet.getRow(i).getCell(24);
            newPribil.setCellValue(prodSumRozn.getNumericCellValue() - prodSumOpt.getNumericCellValue());

            HSSFCell sumSpis = sheet.getRow(i).getCell(29);
            HSSFCell spis = sheet.getRow(i).getCell(28);
            sumSpis.setCellValue(spis.getNumericCellValue()*doubOpt);
        }
    }

    public static void writeAllItog(HSSFSheet sheet, int a){
        double skladBeginSumDoub = 0;
        double kupiliShtukDoub = 0;
        double sumOptDoub = 0;
        double sumRoznDoub = 0;
        double pribilDoub = 0;
        double peremestVavtDoub = 0;
        double sumPeremestVavtDoub = 0;
        double scladEndSumDoub = 0;
        double automWithNewShtDoub = 0;
        double automWithNewSumDoub = 0;
        double prodajaDoub = 0;
        double prodajaSumRoznDoub = 0;
        double prodajaSumOptDoub = 0;
        double prodajaPribilDoub = 0;
        double automatEndDayShtDoub = 0;
        double automatEndDaySumDoub = 0;
        double sumSpisanieDoub = 0;

        for (int i = 2; i < a; i++) {
            HSSFCell skladBeginSum = sheet.getRow(i).getCell(3);
            HSSFCell kupiliShtuk = sheet.getRow(i).getCell(4);
            HSSFCell sumOpt = sheet.getRow(i).getCell(6);
            HSSFCell sumRozn = sheet.getRow(i).getCell(7);
            HSSFCell pribil = sheet.getRow(i).getCell(10);
            HSSFCell peremestVavt = sheet.getRow(i).getCell(11);
            HSSFCell sumPeremestVavt = sheet.getRow(i).getCell(12);
            HSSFCell scladEndSum = sheet.getRow(i).getCell(14);
            HSSFCell automWithNewSht = sheet.getRow(i).getCell(19);
            HSSFCell automWithNewSum = sheet.getRow(i).getCell(20);
            HSSFCell prodaja = sheet.getRow(i).getCell(21);
            HSSFCell prodajaSumRozn = sheet.getRow(i).getCell(22);
            HSSFCell prodajaSumOpt = sheet.getRow(i).getCell(23);
            HSSFCell prodajaPribil = sheet.getRow(i).getCell(24);
            HSSFCell automatEndDaySht = sheet.getRow(i).getCell(25);
            HSSFCell automatEndDaySum = sheet.getRow(i).getCell(26);
            HSSFCell sumSpisanie = sheet.getRow(i).getCell(29);
            skladBeginSumDoub += skladBeginSum.getNumericCellValue();
            kupiliShtukDoub += kupiliShtuk.getNumericCellValue();
            sumOptDoub += sumOpt.getNumericCellValue();
            sumRoznDoub += sumRozn.getNumericCellValue();
            pribilDoub += pribil.getNumericCellValue();
            peremestVavtDoub += peremestVavt.getNumericCellValue();
            sumPeremestVavtDoub += sumPeremestVavt.getNumericCellValue();
            scladEndSumDoub += scladEndSum.getNumericCellValue();
            automWithNewShtDoub += automWithNewSht.getNumericCellValue();
            automWithNewSumDoub += automWithNewSum.getNumericCellValue();
            prodajaDoub += prodaja.getNumericCellValue();
            prodajaSumRoznDoub += prodajaSumRozn.getNumericCellValue();
            prodajaSumOptDoub += prodajaSumOpt.getNumericCellValue();
            prodajaPribilDoub += prodajaPribil.getNumericCellValue();
            automatEndDayShtDoub += automatEndDaySht.getNumericCellValue();
            automatEndDaySumDoub += automatEndDaySum.getNumericCellValue();
            sumSpisanieDoub += sumSpisanie.getNumericCellValue();

        }


        HSSFCell skladBeginSum = sheet.getRow(a).getCell(3);
        skladBeginSum.setCellValue(skladBeginSumDoub);
        HSSFCell kupiliShtuk = sheet.getRow(a).getCell(4);
        kupiliShtuk.setCellValue(kupiliShtukDoub);
        HSSFCell sumOpt = sheet.getRow(a).getCell(6);
        sumOpt.setCellValue(sumOptDoub);
        HSSFCell sumRozn = sheet.getRow(a).getCell(7);
        sumRozn.setCellValue(sumRoznDoub);
        HSSFCell pribil = sheet.getRow(a).getCell(10);
        pribil.setCellValue(pribilDoub);
        HSSFCell peremestVavt = sheet.getRow(a).getCell(11);
        peremestVavt.setCellValue(peremestVavtDoub);
        HSSFCell sumPeremestVavt = sheet.getRow(a).getCell(12);
        sumPeremestVavt.setCellValue(sumPeremestVavtDoub);
        HSSFCell scladEndSum = sheet.getRow(a).getCell(14);
        scladEndSum.setCellValue(scladEndSumDoub);
        HSSFCell automWithNewSht = sheet.getRow(a).getCell(19);
        automWithNewSht.setCellValue(automWithNewShtDoub);
        HSSFCell automWithNewSum = sheet.getRow(a).getCell(20);
        automWithNewSum.setCellValue(automWithNewSumDoub);
        HSSFCell prodaja = sheet.getRow(a).getCell(21);
        prodaja.setCellValue(prodajaDoub);
        HSSFCell prodajaSumRozn = sheet.getRow(a).getCell(22);
        prodajaSumRozn.setCellValue(prodajaSumRoznDoub);
        HSSFCell prodajaSumOpt = sheet.getRow(a).getCell(23);
        prodajaSumOpt.setCellValue(prodajaSumOptDoub);
        HSSFCell prodajaPribil = sheet.getRow(a).getCell(24);
        prodajaPribil.setCellValue(prodajaPribilDoub);
        HSSFCell automatEndDaySht = sheet.getRow(a).getCell(25);
        automatEndDaySht.setCellValue(automatEndDayShtDoub);
        HSSFCell automatEndDaySum = sheet.getRow(a).getCell(26);
        automatEndDaySum.setCellValue(automatEndDaySumDoub);
        HSSFCell sumSpisanie = sheet.getRow(a).getCell(29);
        sumSpisanie.setCellValue(sumSpisanieDoub);

    }

    public static void itogWriteCells(HSSFSheet sheet, String today, int a){
        HSSFCell itog = sheet.getRow(a+1).getCell(24);
        HSSFCell pribil = sheet.getRow(a).getCell(24);

        HSSFCell itogSumSpis = sheet.getRow(a+1).getCell(29);
        HSSFCell sumSpis = sheet.getRow(a).getCell(29);

        if(today.equals("01")){
            itog.setCellValue(pribil.getNumericCellValue());
            itogSumSpis.setCellValue(sumSpis.getNumericCellValue());
        } else{
            itog.setCellValue(itog.getNumericCellValue() + pribil.getNumericCellValue());
            itogSumSpis.setCellValue(itogSumSpis.getNumericCellValue() + sumSpis.getNumericCellValue());
        }

    }

}
