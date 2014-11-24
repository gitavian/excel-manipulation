package com.excelparser;

import java.io.File;
import java.io.FileInputStream;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.codehaus.jackson.map.ObjectMapper;
import org.json.simple.JSONArray;
import org.json.simple.JSONObject;

import com.google.gson.Gson;
import com.google.gson.GsonBuilder;
import com.google.gson.JsonArray;
import com.google.gson.JsonObject;
import com.google.gson.stream.JsonReader;
import com.model.Client;

public class POIReadExcelFile {

	static {
		Map<String, Class<?>> map = new HashMap<String, Class<?>>();
		map.put("clinet", Client.class);

	}

	@SuppressWarnings("unchecked")
	public static JSONObject excelParser(Map<Integer, String> headerMapper) {
		JSONObject excelJson = new JSONObject();
		try {

			FileInputStream excelFile = new FileInputStream(new File(
					"/Users/octavian/workspace/excel-manipulation/test.xlsx"));

			Workbook workbook = WorkbookFactory.create(excelFile);
//			System.out.println("Number of sheets: "
//					+ workbook.getNumberOfSheets());
			FormulaEvaluator evaluator = workbook.getCreationHelper()
                    .createFormulaEvaluator();
			

			for (int i = 0; i < workbook.getNumberOfSheets(); i++) {

				int rowCounter = 0;
				boolean isFirstRow = true;

				final Sheet sheet = workbook.getSheetAt(i);
				final String sheetName = sheet.getSheetName();
//				System.out.println("Sheet Name: " + sheetName);

				JSONArray sheetJArray = new JSONArray();

				// Iterate through each row
				Iterator<Row> rowIterator = sheet.iterator();

				while (rowIterator.hasNext()) {

					int cellCounter = 0;
					JSONObject rowJObject = null;

					if (!isFirstRow)
						rowJObject = new JSONObject();

					final Row row = rowIterator.next();

					rowCounter++;

					// for each row, iterate through all the columns
					Iterator<Cell> cellIterator = row.cellIterator();

					while (cellIterator.hasNext()) {

						final Cell cell = cellIterator.next();
						cellCounter++;
						final String cellContent = getCellContentToString(cell, evaluator, cellCounter);

						if (isFirstRow) {
							loadHeaderMapper(headerMapper, cellCounter, cellContent);
						} else {
							rowJObject.put(headerMapper.get(cellCounter), cellContent);
						}

					}

					if (!isFirstRow)
						sheetJArray.add(rowJObject);
					isFirstRow = false;
				}
				excelJson.put(sheetName, sheetJArray);
			}
			excelFile.close();
//			System.out.println(excelJson.toJSONString());

		} catch (Exception e) {
			e.printStackTrace();
		}
		return excelJson;
		
	}

	private static String getCellContentToString(Cell cell,
			FormulaEvaluator evaluator, int cellCounter) {

		Object obj = null;
		switch (evaluator.evaluateInCell(cell).getCellType()) {
		case Cell.CELL_TYPE_BLANK:
			return "Blank_"+cellCounter;
		case Cell.CELL_TYPE_NUMERIC:
			obj = cell.getNumericCellValue();
			break;
		case Cell.CELL_TYPE_STRING:
			obj = cell.getStringCellValue();
			break;
		case Cell.CELL_TYPE_BOOLEAN:
			obj = cell.getBooleanCellValue();
		case Cell.CELL_TYPE_FORMULA:
			break;
		}

		return obj.toString();
	}

	public static void main(String args[]) {
		Map<Integer, String> headerMapper = new HashMap<Integer, String>();
		JSONObject excelFileToJson = excelParser(headerMapper);
		
		Gson gson = new GsonBuilder().setPrettyPrinting().create();
		System.out.println(gson.toJson(excelFileToJson));
		System.out.println("END");
		
		int size = excelFileToJson.size();
		
		JSONArray sheet1 = (JSONArray) excelFileToJson.get("Sheet1");
		String row1 = sheet1.get(0).toString();
		
		//String clientList = gson.toJson(sheet1, Client.class);
		
		Client c = gson.fromJson(row1, Client.class);
		System.out.println(c.getBalance());
		
		//System.out.println("The Sheet1 array: "+ gson.toJson(sheet1));
		
		//convertJsonToJavaObject(excelFileToJson);
		
		
	}

	/**
	 * Expected Json format to be parsed:	<br>
	 * {									<br>
	 *	  "Sheet1": [						<br>
	 *	    {								<br>
	 *	      "balance": "300.0",			<br>
	 *	      "name": "John",				<br>
	 *	      "blank_5": "Blank_5",			<br>
	 *	      "surname": "Blank_3",			<br>
	 *	      "number": "1.0"				<br>
	 *	    } ...							<br>
	 * }	    							<br>
	 * @param excelFileToJson
	 */
	private static void convertJsonToJavaObject(JSONObject excelFileToJson) {
		Gson gsonTransformer = new Gson();
		Set a = excelFileToJson.entrySet();
		
		Iterator<?> sheetIterator = a.iterator();
		
		while (sheetIterator.hasNext()) {
			System.out.println(sheetIterator.next());
//			System.out.println(o.toString());
		}
		
		
		
		
		
		//List<Client> clients = (List<Client>) gsonTransformer.fromJson(sheet1, Client.class);
		//System.out.println(clients.getBalance());		
	}

	/**
	 * Create map for the excel header (first row) <br>
	 * E.g.:<br>
	 * headerMapper<1,"name"> <br>
	 * headerMapper<2,"surname"> <br>
	 * headerMapper<3,"age"> <br>
	 * 
	 * @param headerMapper
	 *            header map
	 * @param hashIndex
	 *            column index to be used to build the map as a key
	 * @param headerName
	 *            column name to be used in the key
	 */
	static void loadHeaderMapper(Map<Integer, String> headerMapper,
			int hashIndex, String headerName) {
		if (!(headerMapper != null)) {
			headerMapper = new HashMap<Integer, String>();
		}
			headerMapper.put(hashIndex, headerName.toLowerCase());
	}
}
