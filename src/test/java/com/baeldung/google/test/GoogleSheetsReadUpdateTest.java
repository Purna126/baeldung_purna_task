package com.baeldung.google.test;
import com.baeldung.google.util.SheetServiceUtil;
import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.model.BatchUpdateValuesRequest;
import com.google.api.services.sheets.v4.model.BatchUpdateValuesResponse;
import com.google.api.services.sheets.v4.model.ValueRange;
import org.apache.commons.lang3.RandomStringUtils;
import java.io.IOException;
import java.security.GeneralSecurityException;
import java.util.ArrayList;
import java.util.List;

public class GoogleSheetsReadUpdateTest {
    private static Sheets sheetsService;
    private static final String SPREADSHEET_ID = "1GItiL_r3A8f5DW_cq9hxvl8-nJ9E-qOm7IJMMIoUMxs";
    public static void main(String... args) throws GeneralSecurityException, IOException {
        final String range = "A:E";
        sheetsService = SheetServiceUtil.getSheetsService();
        System.out.println("sheetsService:::" + sheetsService);
        String columnName = "category_description";
        ValueRange response = sheetsService.spreadsheets().values().get(SPREADSHEET_ID, range).execute();
        response.setMajorDimension("COLUMNS");
        List<List<Object>> values = response.getValues();
        List<Object> rowList = new ArrayList<>();
        if (values == null || values.isEmpty()) {
            System.out.println("No data found.");
        } else {
            int columnIndex = SheetServiceUtil.getColumnIndexByName(sheetsService, SPREADSHEET_ID, range, columnName);
            int headerSize = values.get(0).size();
            for (List<Object> row : values) {
                String genString = RandomStringUtils.random(20, true, false);
                if (headerSize == row.size()) {
                    if (!row.get(columnIndex).equals("category_description")) {
                        row.set(columnIndex, row.get(columnIndex) +
                                genString);
                    }
                } else {
                    row.add(columnIndex, genString);
                }
            }
        }
        //calling Update category_description
        randomUpdateToSpreadSheet(response, range);
    }

    //This method purpose is update the spread sheet.
    public static void randomUpdateToSpreadSheet(ValueRange response, String range) throws IOException {
        List<ValueRange> data = new ArrayList<>();
        data.add(new ValueRange()
                .setRange(range)
                .setValues(response.getValues()));
        BatchUpdateValuesRequest batchBody = new BatchUpdateValuesRequest()
                .setValueInputOption("USER_ENTERED")
                .setData(data);
        BatchUpdateValuesResponse batchResult = sheetsService.spreadsheets().values()
                .batchUpdate(SPREADSHEET_ID, batchBody)
                .execute();
    }
}
