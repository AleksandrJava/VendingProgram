import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ss.usermodel.CellType;

public class WorkWithZayavka {
    public static void buyAndMovedToAutomat(HSSFSheet zayavkaSheet, HSSFSheet sheet, int a) {
        for (int i = 1; i <= 1; i++) {
            for (int j = 2; j < a; j++) {
                HSSFCell delete1 = sheet.getRow(j).getCell(4);
                HSSFCell delete2 = sheet.getRow(j).getCell(11);
                HSSFCell delete3 = sheet.getRow(j).getCell(12);
                delete1.setCellValue(0);
                delete2.setCellValue(0);
                delete3.setCellValue(0);
            }
        }

        HSSFCell allCell = zayavkaSheet.getRow(0).getCell(1);
        double doubleAllCell = allCell.getNumericCellValue();
        double doubSneaksOrTea = 0;

        for (int i = 1; i <= 1; i++) {
            for (int j = 3; j < 3 + doubleAllCell; j++) {
                HSSFCell currentCell = zayavkaSheet.getRow(j).getCell(i);
                double doubleCurrentCell = currentCell.getNumericCellValue();
                for (int k = 2; k < a; k++) {
                    HSSFCell cellOnBigTable = sheet.getRow(k).getCell(1);
                    double doubleCellOnBigTable = cellOnBigTable.getNumericCellValue();
                    if (doubleCurrentCell == doubleCellOnBigTable) {
                        HSSFCell sneaksOrTea = sheet.getRow(k).getCell(31);
                        doubSneaksOrTea = sneaksOrTea.getNumericCellValue();
                        break;
                    }
                }

                if (doubSneaksOrTea == 1) {
                    addBuyAndMovedPieTea(sheet, zayavkaSheet, doubleCurrentCell, i, j, a);
                } else {
                    addBuyAndMovedSneaks(sheet, zayavkaSheet, doubleCurrentCell, i, j, a);
                }


            }
        }
        allCell.setCellValue(0);
    }
    private static void addBuyAndMovedPieTea(HSSFSheet sheet,HSSFSheet zayavkaSheet, double currentCell, int ii, int jj, int a){
        for (int i = 1; i <= 1 ; i++) {
            for (int j = 2; j < a ; j++) {
                HSSFCell cell = sheet.getRow(j).getCell(i);

                double doub = cell.getNumericCellValue();

                if(doub == currentCell){
                    HSSFCell mustHaveCell = zayavkaSheet.getRow(jj).getCell(ii + 1);
                    double mustCellDouble = mustHaveCell.getNumericCellValue();
                    HSSFCell cell1 = sheet.getRow(j).getCell(4);
                    HSSFCell cell2 = sheet.getRow(j).getCell(11);
                    cell1.setCellValue(mustCellDouble);
                    cell2.setCellValue(mustCellDouble);

                    HSSFCell cell4 = zayavkaSheet.getRow(jj).getCell(0);
                    cell4.setCellType(CellType.STRING);
                    cell4.setCellValue("нет");

                    for (int k = 1; k <= 2; k++) {
                        HSSFCell cell3 = zayavkaSheet.getRow(jj).getCell(k);
                        cell3.setCellType(CellType.NUMERIC);
                        cell3.setCellValue(0);
                    }

                    break;
                }
            }
        }
    }

    private static void addBuyAndMovedSneaks(HSSFSheet sheet,HSSFSheet zayavkaSheet, double currentCell, int ii, int jj, int a){
        for (int i = 1; i <= 1 ; i++) {
            for (int j = 2; j < a ; j++) {
                HSSFCell cell = sheet.getRow(j).getCell(i);

                double doub = cell.getNumericCellValue();

                if(doub == currentCell){
                    HSSFCell mustHaveCell = zayavkaSheet.getRow(jj).getCell(ii + 1);
                    double mustCellDouble = mustHaveCell.getNumericCellValue();
                    HSSFCell cell1 = sheet.getRow(j).getCell(11);
                    cell1.setCellValue(mustCellDouble);

                    HSSFCell cell4 = zayavkaSheet.getRow(jj).getCell(0);
                    cell4.setCellType(CellType.STRING);
                    cell4.setCellValue("нет");

                    for (int k = 1; k <= 2; k++) {
                        HSSFCell cell3 = zayavkaSheet.getRow(jj).getCell(k);
                        cell3.setCellType(CellType.NUMERIC);
                        cell3.setCellValue(0);
                    }
                    break;
                }
            }
        }
    }

    public static void writeBigRequest(HSSFSheet sheet, HSSFSheet requestSheet, int a){
        for (int i = 1; i <= 1; i++) {
            for (int j = 3; j < a+1; j++) {
                HSSFCell needCell = requestSheet.getRow(j).getCell(2);
                needCell.setCellValue(0);
            }
        }

        for (int i = 1; i <= 1; i++) {
            for (int j = 2; j < a; j++) {
                HSSFCell max = sheet.getRow(j).getCell(30);
                HSSFCell automatEndDay = sheet.getRow(j).getCell(25);
                HSSFCell currentSclad = sheet.getRow(j).getCell(13);
                HSSFCell numberOfCell = sheet.getRow(j).getCell(1);

                double numberOfCellDoub = numberOfCell.getNumericCellValue();
                double currentScladDoub = currentSclad.getNumericCellValue();
                double automatEndDayDoub = automatEndDay.getNumericCellValue();
                double maxDoub = max.getNumericCellValue();

                if(currentScladDoub != 0 || (numberOfCellDoub<140 && numberOfCellDoub > 129) || (numberOfCellDoub<960 && numberOfCellDoub > 949)){
                    double weNeed = maxDoub - automatEndDayDoub;
                    if(weNeed < currentScladDoub){
                        writeBigRequestDop(requestSheet, numberOfCellDoub, weNeed, a);
                    }else if((numberOfCellDoub<140 && numberOfCellDoub > 129) || (numberOfCellDoub<960 && numberOfCellDoub > 949)){
                        writeBigRequestDop(requestSheet, numberOfCellDoub, weNeed, a);
                    }else {
                        weNeed = currentScladDoub;
                        writeBigRequestDop(requestSheet, numberOfCellDoub, weNeed, a);
                    }
                }
            }
        }
    }

    private static void writeBigRequestDop(HSSFSheet requestSheet, double numberOfCellDoub, double weNeed, int a){
        for (int i = 1; i <= 1; i++) {
            for (int j = 3; j < a+1; j++) {
                HSSFCell cell = requestSheet.getRow(j).getCell(i);
                double doub = cell.getNumericCellValue();
                if(doub == numberOfCellDoub){
                    HSSFCell needCell = requestSheet.getRow(j).getCell(2);
                    needCell.setCellValue(weNeed);
                    break;
                }
            }
        }
    }

    public static void writeShortRequest(HSSFSheet zayavkaBigSheet, HSSFSheet zayavkaSheet, int a, Logger log){
        try {
            for (int i = 1; i <= 1; i++) {
                for (int j = 3; j < a+1; j++) {
                    HSSFCell weNeed = zayavkaBigSheet.getRow(j).getCell(2);

                    double weNeedDouble = weNeed.getNumericCellValue();

                    if(weNeedDouble != 0){
                        HSSFCell cell = zayavkaBigSheet.getRow(j).getCell(1);
                        HSSFCell name = zayavkaBigSheet.getRow(j).getCell(0);

                        double cellDoub = cell.getNumericCellValue();
                        String nameStr = name.getStringCellValue();

                        HSSFCell allCellInShort = zayavkaSheet.getRow(0).getCell(1);
                        double allCellDoub = allCellInShort.getNumericCellValue();
                        int allCellInt = (int) allCellDoub;

                        HSSFCell shortName = zayavkaSheet.getRow(3+allCellInt).getCell(0);
                        HSSFCell shortCell = zayavkaSheet.getRow(3+allCellInt).getCell(1);
                        HSSFCell shortNeed = zayavkaSheet.getRow(3+allCellInt).getCell(2);



                        shortName.setCellValue(nameStr);
                        shortCell.setCellValue(cellDoub);
                        shortNeed.setCellValue(weNeedDouble);

                        allCellInShort.setCellValue(allCellDoub + 1);
                    }
                }
            }
        }catch (NullPointerException e){
            log.error("Ne mogu zapisat novuIO zayavku na zavtra. Prover'te fail 'zayavkaItog - list 1' ");
            throw new NullPointerException();
        }
    }

}
