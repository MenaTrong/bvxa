package com.example.springweb.controller;

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
import com.google.api.services.sheets.v4.model.UpdateValuesResponse;
import com.google.api.services.sheets.v4.model.ValueRange;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.security.GeneralSecurityException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestParam;

@Controller
public class SheetsQuickstart {
	private static final JsonFactory JSON_FACTORY = GsonFactory.getDefaultInstance();
	private static final String TOKENS_DIRECTORY_PATH = "tokens/path";
	private static final String extistingSpreadSheetID = "1ef-dLEi0b0nlNuHzqMU3OcnGG2D2ngKqirF7WKdQWfo";

	private static final List<String> SCOPES = Arrays.asList(SheetsScopes.SPREADSHEETS, SheetsScopes.DRIVE);
	private static final String CREDENTIALS_FILE_PATH = "/credentials.json";
	static Sheets.Spreadsheets spreadsheets;

	private static Credential getCredentials(final NetHttpTransport HTTP_TRANSPORT) throws IOException {
		// Load client secrets.
		InputStream in = SheetsQuickstart.class.getResourceAsStream(CREDENTIALS_FILE_PATH);
		if (in == null) {
			throw new FileNotFoundException("Resource not found: " + CREDENTIALS_FILE_PATH);
		}
		GoogleClientSecrets clientSecrets = GoogleClientSecrets.load(JSON_FACTORY, new InputStreamReader(in));

		// Build flow and trigger user authorization request.
		GoogleAuthorizationCodeFlow flow = new GoogleAuthorizationCodeFlow.Builder(HTTP_TRANSPORT, JSON_FACTORY,
				clientSecrets, SCOPES)
				.setDataStoreFactory(new FileDataStoreFactory(new java.io.File(TOKENS_DIRECTORY_PATH)))
				.setAccessType("offline").build();
		LocalServerReceiver receiver = new LocalServerReceiver.Builder().setPort(8888).build();
		return new AuthorizationCodeInstalledApp(flow, receiver).authorize("user");
	}

	public static void getSpreadsheetInstance() throws GeneralSecurityException, IOException {
		final NetHttpTransport HTTP_TRANSPORT = GoogleNetHttpTransport.newTrustedTransport();
		spreadsheets = new Sheets.Builder(GoogleNetHttpTransport.newTrustedTransport(),
				GsonFactory.getDefaultInstance(), getCredentials(HTTP_TRANSPORT))
				.setApplicationName("Google Sheet Java Integrate").build().spreadsheets();

	}

	public static void writeSheet(List<Object> inputData, String sheetAndRange, String existingSpreadSheetID)
			throws IOException {
		List<List<Object>> values = Arrays.asList(inputData);
		ValueRange body = new ValueRange().setValues(values);
		UpdateValuesResponse result = spreadsheets.values().update(existingSpreadSheetID, sheetAndRange, body)
				.setValueInputOption("RAW").execute();
		System.out.printf("%d cells updated. \n", result.getUpdatedCells());
	}

	public static void writeDataGoogleSheets(String sheetName, List<Object> data, String existingSpreadSheetID)
			throws IOException {
		int nextRow = getRows(sheetName, existingSpreadSheetID) + 1;
		writeSheet(data, "!A" + nextRow, existingSpreadSheetID);
	}

	public static int getRows(String sheetName, String existingSpreadSheetID) throws IOException {
		List<List<Object>> values = spreadsheets.values().get(existingSpreadSheetID, sheetName).execute().getValues();
		int numRows = values != null ? values.size() : 0;
		return numRows;
	}

	@GetMapping("/signup-information")
	public String signupInformationString (@RequestParam("fullname") String fullName,
			   							   @RequestParam("email") String email,
			   							   @RequestParam("numberphone") String numberPhone,
			   							   @RequestParam("tuoithai") String tuoiThai,
			   							   @RequestParam("coso") String coSo,
			   							   @RequestParam("thongtinthem") String thongTinThem) throws IOException, GeneralSecurityException {
		try {
			getSpreadsheetInstance();
			writeDataGoogleSheets("Trang tính1", new ArrayList<Object>(Arrays.asList(fullName, email, numberPhone, tuoiThai, coSo, thongTinThem)), extistingSpreadSheetID);
	
			return "redirect:/?signup-information=success";
		} catch (Exception e) {
			return "redirect:/?signup-information=failure";
		}
	}
}
