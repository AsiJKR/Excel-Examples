package utils;

import org.apache.poi.ss.usermodel.Sheet;
import org.testng.annotations.DataProvider;
import utils.ExcelUtil;

import java.io.IOException;
import java.security.GeneralSecurityException;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;


public class DataProviderUtil {

    private static String excelFilePath = "src/main/resources/excel/DataProviderSheet.xlsx";

    public static ArrayList<ArrayList<Object>> getDataFromExcelSheet() {

        //local sheet reader
//        Sheet s = ExcelUtil.getExcelSheet(sheetName);
//        int noOfRows = ExcelUtil.getNumberOfValidRows(s, s.getLastRowNum());
//        ExcelUtil.colMapByName = ExcelUtil.getIndexByColumnNames(s);

        //for gsheet
        List<List<Object>> list = null;
        try {
            list = ExcelUtil.readGoogleDoc();
        } catch (GeneralSecurityException | IOException e) {
            e.printStackTrace();
        }

        ArrayList<ArrayList<Object>> arrayList1 = new ArrayList<>();

        for (int i = 1; i < Objects.requireNonNull(list).size(); i++) {
            String[] skus = list.get(i).get(0).toString().split(",");
            String[] qtys = list.get(i).get(1).toString().split(",");
            String[] lensType = list.get(i).get(2).toString().split(",");

//            if (!lensType[0].equals("")){
                for (int k=0;k<qtys.length;k++){
                    for (int z=k;z<lensType.length;z++){
                        if (!lensType[z].equals("")){
                            qtys[k] = qtys[k]+"*"+lensType[z];
                        }
                        break;
                    }
                }
//            }


            String streetAddress = list.get(i).get(3).toString();
            String city = list.get(i).get(4).toString();
            String phone = list.get(i).get(5).toString();
            String discountCode = list.get(i).get(6).toString();
            String firstName = list.get(i).get(7).toString();
            String lastName = list.get(i).get(8).toString();
            String countryCode = list.get(i).get(9).toString();
            String type = list.get(i).get(10).toString();

            ArrayList<Object> arrayList = new ArrayList<>();
            arrayList.add(0, skus);
            arrayList.add(1, qtys);
            arrayList.add(2, streetAddress);
            arrayList.add(3, city);
            arrayList.add(4, phone);
            arrayList.add(5, discountCode);
            arrayList.add(6, firstName);
            arrayList.add(7, lastName);
            arrayList.add(8, countryCode);
            arrayList.add(9, type);

            arrayList1.add(arrayList);
        }
        return arrayList1;
    }

    public static ArrayList<ArrayList<String[]>> getSKUList(String sheetName) {
        Sheet s = ExcelUtil.getExcelSheet(sheetName);
        int noOfRows = ExcelUtil.getNumberOfValidRows(s, s.getLastRowNum());
        ExcelUtil.colMapByName = ExcelUtil.getIndexByColumnNames(s);

        List<List<Object>> list = null;

        try {
            list = ExcelUtil.readGoogleDoc();
        } catch (GeneralSecurityException | IOException e) {
            e.printStackTrace();
        }

        ArrayList<ArrayList<String[]>> arrayList1 = new ArrayList<>();

        for (int i = 1; i < Objects.requireNonNull(list).size(); i++) {
            String[] skus = list.get(i).get(0).toString().split(",");

            ArrayList<String[]> arrayList = new ArrayList<>();
            arrayList.add(0, skus);

            arrayList1.add(arrayList);
        }
        return arrayList1;
    }

    @DataProvider(name = "getDataAE")
    public static Object[][] getDataAE() {

        ArrayList<ArrayList<Object>> arrayList = getDataFromExcelSheet();
        Object[][] objects = new Object[arrayList.size()][10];

        for (int i = 0; i < arrayList.size(); i++) {
            for (int j = 0; j < arrayList.get(i).size(); j++) {
                objects[i][j] = arrayList.get(i).get(j);
            }
        }

        return objects;
    }

}
