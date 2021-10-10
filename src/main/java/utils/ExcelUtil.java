package utils;

import com.google.api.client.auth.oauth2.Credential;
import com.google.api.client.extensions.java6.auth.oauth2.AuthorizationCodeInstalledApp;
import com.google.api.client.extensions.jetty.auth.oauth2.LocalServerReceiver;
import com.google.api.client.googleapis.auth.oauth2.GoogleAuthorizationCodeFlow;
import com.google.api.client.googleapis.auth.oauth2.GoogleClientSecrets;
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.http.javanet.NetHttpTransport;
import com.google.api.client.json.JsonFactory;
import com.google.api.client.json.jackson2.JacksonFactory;
import com.google.api.client.util.store.FileDataStoreFactory;
import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.SheetsScopes;
import com.google.api.services.sheets.v4.model.*;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Sheet;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.*;
import java.security.GeneralSecurityException;
import java.util.*;


public class ExcelUtil {

    public static Map<String, Integer> colMapByName = null;
    public static String excelFilePath = "src/main/resources/excel/DataProviderSheet.xlsx";
    static Logger logger = LoggerFactory.getLogger(ExcelUtil.class);
    private Map<String, Integer> colMapByNameForSales = null;

    private static final String APPLICATION_NAME = "name"; //name of the project in dev console
    private static final JsonFactory JSON_FACTORY = JacksonFactory.getDefaultInstance();
    private static final String TOKENS_DIRECTORY_PATH = "tokens"; //directory for token

    private static final List<String> SCOPES = Collections.singletonList(SheetsScopes.SPREADSHEETS);
    private static final String CREDENTIALS_FILE_PATH = "/client_secret.json"; //json file you get from oauth

    public static Sheet readExcelSheet(String sheetName) {

        Sheet excelSheet = null;

        try {

            if (excelFilePath.isEmpty()) {
                logger.error("No excel FilePath specified. Please set environment variable EXCEL_FILE_LOCATION");
                return null;
            }
            logger.info("excelFilePath: " + excelFilePath);

            // Creating a Workbook from an Excel file (.xls or .xlsx)
            Workbook workbook = WorkbookFactory.create(new File(excelFilePath));

            // Retrieving the number of sheets in the Workbook
            logger.info("Workbook has " + workbook.getNumberOfSheets() + " Sheets");

            for (Sheet sheet : workbook) {
                logger.info("=> " + sheet.getSheetName());
                if (sheetName.equals(sheet.getSheetName())) {
                    excelSheet = sheet;
                }
            }

        } catch (IOException e) {
            logger.error("IO Exception");
        } catch (InvalidFormatException e) {
            logger.error("InvalidFormat Exception");
        }

        return excelSheet;
    }

    public static Sheet getExcelSheet(String sheetName) {

        logger.info("Reading excel file");
        Sheet sheet = readExcelSheet(sheetName);

        logger.info("Calculating the number of rows and columns");
        int noOfColumns = sheet.getRow(0).getLastCellNum();
        int noOfRows = sheet.getLastRowNum();

        logger.info("No Of rows rough count: " + noOfRows);
        logger.info("No Of columns rough count: " + noOfColumns);
        logger.info("" + getIndexByColumnNames(sheet));

        return convertColumnNamesToCamelCase(sheet, noOfColumns);
    }

    public static String capitalizeWord(String str) {
        StringBuffer s = new StringBuffer();

        // Declare a character of space
        // To identify that the next character is the starting
        // of a new word
        char ch = ' ';
        for (int i = 0; i < str.length(); i++) {

            // If previous character is space and current
            // character is not space then it shows that
            // current letter is the starting of the word
            if (ch == ' ' && str.charAt(i) != ' ')
                s.append(Character.toUpperCase(str.charAt(i)));
            else
                s.append(str.charAt(i));
            ch = str.charAt(i);
        }

        // Return the string with trimming
        return s.toString().trim();
    }

    private static Sheet convertColumnNamesToCamelCase(Sheet sheet, int noOfColumns) {

        Row row;
        try {
            row = sheet.getRow(0);
        } catch (Exception e) {
            logger.error("Null row. Please provide a valid excel sheet.");
            throw e;
        }

        for (int i = 0; i < noOfColumns; i++) {
            try {
                Cell cell = row.getCell(i, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                String formattedColumnName = capitalizeWord(getCellValue(cell));

                if (formattedColumnName.isEmpty()) {
                    logger.error("Column heading cannot be empty");
                    throw new NullPointerException("One of column headings is empty");
                }

                sheet.getRow(0).getCell(i).setCellValue(formattedColumnName);
            } catch (Exception e) {
                logger.error("Error reading the cell value");
                throw e;
            }
        }
        return sheet;
    }

    public static int getNumberOfValidRows(Sheet sheet, int noOfRows) {

        logger.info("Calculating the valid number of rows present, for rough row count: {}", noOfRows);

        int numOfValidRows = 0;

        for (int i = 0; i <= noOfRows; i++) {
            if (sheet.getRow(i) == null) {
                logger.error("Row is null for row number: {}, number of all the rows present with invalid rows: {}", i,
                        noOfRows);
                logger.error("Ignore this error message. Row count taken as: {}", i);
                numOfValidRows = i;
                break;
            }
            Cell orderIdCell = sheet.getRow(i).getCell(0);
            numOfValidRows = i;
            if (orderIdCell.toString().equals("")) {
                break;
            }
        }
        return numOfValidRows;
    }

    public static Map<String, Integer> getIndexByColumnNames(Sheet sheet) {
        Map<String, Integer> colMapByName = new HashMap<>();

        DataFormatter dataFormatter = new DataFormatter();

        if (sheet.getRow(0).cellIterator().hasNext()) {
            for (int j = 0; j < sheet.getRow(0).getLastCellNum(); j++) {
                colMapByName.put(dataFormatter.formatCellValue(sheet.getRow(0).getCell(j)), j);
            }
        }

        return colMapByName;
    }

    public static String getCell(Sheet sheet, int row, String colName) {

        Row r = sheet.getRow(row);
        Integer i = 0;

        i = colMapByName.get(colName);

        logger.info("row: {} | colName: {} | i: {}", row, colName, i);
        try {
            Cell cell = r.getCell(i, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
            return getCellValue(cell);
        } catch (Exception e) {
            logger.error("Error in reading excel cell value. Check if column name is correct");
            throw e;
        }

    }

    public static String getCellValue(Cell cell) {
        if (cell == null) {
            return StringUtils.EMPTY;
        }

        switch (cell.getCellTypeEnum()) {
            case BOOLEAN:
                return (cell.getBooleanCellValue() + StringUtils.EMPTY).trim();
            case STRING:
                return (cell.getRichStringCellValue().getString() + StringUtils.EMPTY).trim();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return (cell.getDateCellValue() + StringUtils.EMPTY).trim();
                } else {
                    return (cell.getNumericCellValue() + StringUtils.EMPTY).trim();
                }
            case FORMULA:
                if (cell.getCachedFormulaResultTypeEnum() == CellType.NUMERIC) {
                    return (cell.getNumericCellValue() + StringUtils.EMPTY).trim();
                } else if (cell.getCachedFormulaResultTypeEnum() == CellType.STRING) {
                    return (cell.getRichStringCellValue() + StringUtils.EMPTY).trim();
                }
                break;
            default:
                return (StringUtils.EMPTY);
        }
        return StringUtils.EMPTY;
    }

    public static void writeToExcel(int rowIndex, String orderID) throws Exception {

        File file = new File(excelFilePath);
        FileInputStream inputStream = new FileInputStream(file);
        Workbook workbook = WorkbookFactory.create(inputStream);
        Sheet sheet1 = workbook.getSheet("AE");

        int colIndex = 12;

        sheet1.getRow(rowIndex).getCell(colIndex, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).setCellValue(orderID);

        FileOutputStream outputstream = new FileOutputStream(file);
        workbook.write(outputstream);
        workbook.close();
        outputstream.close();
    }

    private void printCellValue(Cell cell) {
        switch (cell.getCellTypeEnum()) {
            case BOOLEAN:
                System.out.print(cell.getBooleanCellValue());
                break;
            case STRING:
                System.out.print(cell.getRichStringCellValue().getString());
                break;
            case NUMERIC:
                System.out.print(cell.getDateCellValue());

                System.out.print(cell.getNumericCellValue());
                break;
            case FORMULA:
                if (cell.getCachedFormulaResultTypeEnum() == CellType.NUMERIC) {
                    System.out.print(" " + cell.getNumericCellValue());
                } else if (cell.getCachedFormulaResultTypeEnum() == CellType.STRING) {
                    System.out.print(" " + cell.getRichStringCellValue());
                }
                break;
            default:
                System.out.print("");
        }

        System.out.print("\t");
    }

    private static Credential getCredentials(final NetHttpTransport HTTP_TRANSPORT) throws IOException {
        // Load client secrets.
        InputStream in = ExcelUtil.class.getResourceAsStream(CREDENTIALS_FILE_PATH);
        if (in == null) {
            throw new FileNotFoundException("Resource not found: " + CREDENTIALS_FILE_PATH);
        }
        GoogleClientSecrets clientSecrets = GoogleClientSecrets.load(JSON_FACTORY, new InputStreamReader(in));

        // Build flow and trigger user authorization request.
        GoogleAuthorizationCodeFlow flow = new GoogleAuthorizationCodeFlow.Builder(
                HTTP_TRANSPORT, JSON_FACTORY, clientSecrets, SCOPES)
                .setDataStoreFactory(new FileDataStoreFactory(new File(TOKENS_DIRECTORY_PATH)))
                .setAccessType("offline")
                .build();
        LocalServerReceiver receiver = new LocalServerReceiver.Builder().setPort(8888).build();
        return new AuthorizationCodeInstalledApp(flow, receiver).authorize("user");
    }

    public static List<List<Object>> readGoogleDoc() throws GeneralSecurityException, IOException {

        // Build a new authorized API client service.
        final NetHttpTransport HTTP_TRANSPORT = GoogleNetHttpTransport.newTrustedTransport();
        final String spreadsheetId = "sheet id"; //id of google sheet
        final String range = "sheet"; //sheet name
        Sheets service = new Sheets.Builder(HTTP_TRANSPORT, JSON_FACTORY, getCredentials(HTTP_TRANSPORT))
                .setApplicationName(APPLICATION_NAME)
                .build();
        ValueRange response = service.spreadsheets().values()
                .get(spreadsheetId, range)
                .execute();
        return response.getValues();
    }

    public static void updategsheet(int row,String orderID) throws GeneralSecurityException, IOException {

        final NetHttpTransport HTTP_TRANSPORT = GoogleNetHttpTransport.newTrustedTransport();

        Sheets service = new Sheets.Builder(HTTP_TRANSPORT, JSON_FACTORY, getCredentials(HTTP_TRANSPORT))
                .setApplicationName(APPLICATION_NAME)
                .build();

        List<List<Object>> values = Arrays.asList(
                Arrays.asList(orderID)
                // Additional rows ...
        );
        ValueRange body = new ValueRange()
                .setValues(values);
        UpdateValuesResponse result =
                service.spreadsheets().values().update("gsheetid", "AE!M"+(row+1), body) //SheetName!range
                        .setValueInputOption("RAW")
                        .execute();
        System.out.printf("%d cells updated.", result.getUpdatedCells());
    }
}
