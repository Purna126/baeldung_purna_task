package com.baeldung.google.util;

import com.google.api.client.auth.oauth2.Credential;
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.http.javanet.NetHttpTransport;
import com.google.api.client.json.jackson2.JacksonFactory;
import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.model.ValueRange;

import java.io.IOException;
import java.security.GeneralSecurityException;
import java.util.List;
import java.util.stream.Collectors;

public class SheetServiceUtil {
    private static final String APPLICATION_NAME = "Google Sheets Example";

    public static Sheets getSheetsService() throws IOException, GeneralSecurityException {
        final NetHttpTransport HTTP_TRANSPORT = GoogleNetHttpTransport.newTrustedTransport();
        Credential credential = GoogleAuthorizeUtil.authorize(HTTP_TRANSPORT);
        return new Sheets.Builder(GoogleNetHttpTransport.newTrustedTransport(), JacksonFactory.getDefaultInstance(), credential).setApplicationName(APPLICATION_NAME).build();
    }

    public static int getColumnIndexByName(Sheets sheetsService, String spreadsheetId, String range, String columnName)
            throws IOException {
        // Get the column names
        ValueRange response = sheetsService.spreadsheets().values().get(spreadsheetId, range).execute();
        List<Object> headerRow = response.getValues().get(0);
        List<String> columnNames = headerRow.stream().map(Object::toString).collect(Collectors.toList());

        // Find the column index by name
        int columnIndex = columnNames.indexOf(columnName);
        if (columnIndex == -1) {
            throw new IllegalArgumentException("Column name not found: " + columnName);
        }
        return columnIndex;
    }
}
