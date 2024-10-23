package org.example;

import com.google.api.client.auth.oauth2.Credential;
import com.google.api.client.extensions.java6.auth.oauth2.AuthorizationCodeInstalledApp;
import com.google.api.client.extensions.jetty.auth.oauth2.LocalServerReceiver;
import com.google.api.client.googleapis.auth.oauth2.GoogleAuthorizationCodeFlow;
import com.google.api.client.googleapis.auth.oauth2.GoogleClientSecrets;
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.http.javanet.NetHttpTransport;
import com.google.api.client.json.JsonFactory;
import com.google.api.client.json.gson.GsonFactory;
import com.google.api.client.util.store.FileDataStoreFactory;
import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.SheetsScopes;
import com.google.api.services.sheets.v4.model.ClearValuesRequest;
import com.google.api.services.sheets.v4.model.UpdateValuesResponse;
import com.google.api.services.sheets.v4.model.ValueRange;

import java.io.FileReader;
import java.io.IOException;
import java.security.GeneralSecurityException;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.Random;

public class GoogleSheetsExample {

    private static final String APPLICATION_NAME = "Google Sheets API Java";
    private static final JsonFactory JSON_FACTORY = GsonFactory.getDefaultInstance();
    private static final String SPREADSHEET_ID = "1wDYeOv5YQaYnhoa496o3SD8KjP-TJrNrbukuA6MWpKY";
    private static final String TOKENS_DIRECTORY_PATH = "token";
    private static final String CREDENTIALS_FILE_PATH = "src/main/resources/credentials.json";
    private static final String RANGE = "Sheet3!A1:E500";


    private static Credential getCredentials(final NetHttpTransport HTTP_TRANSPORT) throws IOException {
        GoogleClientSecrets clientSecrets = GoogleClientSecrets.load(JSON_FACTORY, new FileReader(CREDENTIALS_FILE_PATH));
        GoogleAuthorizationCodeFlow flow = new GoogleAuthorizationCodeFlow.Builder(
                HTTP_TRANSPORT, JSON_FACTORY, clientSecrets, Collections.singletonList(SheetsScopes.SPREADSHEETS))
                .setDataStoreFactory(new FileDataStoreFactory(new java.io.File(TOKENS_DIRECTORY_PATH)))
                .setAccessType("offline")
                .build();
        LocalServerReceiver receiver = new LocalServerReceiver.Builder().setPort(8881).build();
        return new AuthorizationCodeInstalledApp(flow, receiver).authorize("user");
    }

    private static Sheets getSheetsService() throws IOException, GeneralSecurityException {
        final NetHttpTransport HTTP_TRANSPORT = GoogleNetHttpTransport.newTrustedTransport();
        Credential credentials = getCredentials(HTTP_TRANSPORT);
        return new Sheets.Builder(HTTP_TRANSPORT, JSON_FACTORY, credentials)
                .setApplicationName(APPLICATION_NAME)
                .build();
    }
    public static void generateDemoData() throws IOException, GeneralSecurityException {
        Sheets service = getSheetsService();
        List<List<Object>> data = new ArrayList<>();
        Random random = new Random();
        DateTimeFormatter dateFormatter = DateTimeFormatter.ofPattern("yyyy-MM-dd");
        for (int i = 0; i < 500; i++) {
            int demoId = i + 1; // Demo ID
            String demoName = "Demo Name " + demoId;
            String demoEmail = "demo" + demoId + "@example.com";
            String demoPhone = "123-456-789" + (random.nextInt(10));
            LocalDate date = LocalDate.now().minusDays(500 - i);
            String formattedDate = date.format(dateFormatter);
            data.add(List.of(demoId, demoName, demoEmail, demoPhone, formattedDate));
        }
        ValueRange body = new ValueRange().setValues(data);
        UpdateValuesResponse result = service.spreadsheets().values()
                .update(SPREADSHEET_ID, RANGE, body)
                .setValueInputOption("RAW")
                .execute();
        System.out.println(result.getUpdatedCells() + " cells updated.");
    }

    public static void deleteOldRecords() throws IOException, GeneralSecurityException {
        Sheets service = getSheetsService();
        ValueRange response = service.spreadsheets().values()
                .get(SPREADSHEET_ID, RANGE)
                .execute();
        List<List<Object>> records = response.getValues();

        if (records != null && !records.isEmpty()) {
            String latestDateStr = (String) records.get(records.size() - 1).get(4);
            LocalDate latestDate;

            try {
                latestDate = LocalDate.parse(latestDateStr);
            } catch (DateTimeParseException e) {
                System.out.println("Invalid date format: " + latestDateStr);
                return;
            }
            LocalDate thresholdDate = latestDate.minusDays(3);
            List<List<Object>> updatedRecords = new ArrayList<>();
            for (List<Object> record : records) {
                if (record.size() > 4) {
                    String dateStr = (String) record.get(4);
                    LocalDate recordDate;
                    try {
                        recordDate = LocalDate.parse(dateStr);
                    } catch (DateTimeParseException e) {
                        System.out.println("Invalid date format: " + dateStr);
                        continue;
                    }
                    if (!recordDate.isBefore(thresholdDate)) {
                        updatedRecords.add(record);
                    }
                }
            }
            System.out.println(updatedRecords.size() + " records kept.");
            ClearValuesRequest clearRequest = new ClearValuesRequest();
            service.spreadsheets().values().clear(SPREADSHEET_ID, RANGE, clearRequest).execute();
            System.out.println("Cleared the original records.");
            if (!updatedRecords.isEmpty()) {
                ValueRange body = new ValueRange().setValues(updatedRecords);
                UpdateValuesResponse result = service.spreadsheets().values()
                        .update(SPREADSHEET_ID, RANGE, body)
                        .setValueInputOption("RAW")
                        .execute();
                System.out.println(result.getUpdatedCells() + " cells updated after deletion.");
            } else {
                System.out.println("All records were older than 3 days; no records to update.");
            }
        } else {
            System.out.println("No records found in the specified range.");
        }
    }

    public static void main(String[] args) {
        try {
            //Generate demo data
            generateDemoData();
            //Wait for a while before deleting records (for testing purposes)
            //Thread.sleep(60000);
            //Delete records older than 3 days
            deleteOldRecords();
        } catch (IOException | GeneralSecurityException e) {
            e.printStackTrace();
        }
    }
}