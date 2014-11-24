package com.excelparser;

import java.io.File;
import java.io.FileInputStream;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.json.simple.JSONArray;
import org.json.simple.JSONObject;

public class JsonMapper {

	public JSONObject excelParser(Map<Integer, String> headerMapper) {
		JSONObject excelJson = new JSONObject();
		try {

			FileInputStream excelFile = new FileInputStream(new File(
					"/Users/octavian/workspace/excel-manipulation/test.xlsx"));

			Workbook workbook = WorkbookFactory.create(excelFile);
			// System.out.println("Number of sheets: "
			// + workbook.getNumberOfSheets());
			FormulaEvaluator evaluator = workbook.getCreationHelper()
					.createFormulaEvaluator();

			for (int i = 0; i < workbook.getNumberOfSheets(); i++) {

				int rowCounter = 0;
				boolean isFirstRow = true;

				final Sheet sheet = workbook.getSheetAt(i);
				final String sheetName = sheet.getSheetName();
				// System.out.println("Sheet Name: " + sheetName);

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
						final String cellContent = getCellContentToString(cell,
								evaluator, cellCounter);

						if (isFirstRow) {
							loadHeaderMapper(headerMapper, cellCounter,
									cellContent);
						} else {
							rowJObject.put(headerMapper.get(cellCounter),
									cellContent);
						}

					}

					if (!isFirstRow)
						sheetJArray.add(rowJObject);
					isFirstRow = false;
				}
				excelJson.put(sheetName, sheetJArray);
			}
			excelFile.close();

		} catch (Exception e) {
			e.printStackTrace();
		}
		return excelJson;

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
	public void loadHeaderMapper(Map<Integer, String> headerMapper,
			int hashIndex, String headerName) {
		if (!(headerMapper != null)) {
			headerMapper = new HashMap<Integer, String>();
		}
		headerMapper.put(hashIndex, headerName.toLowerCase());
	}

	private String getCellContentToString(Cell cell,
			FormulaEvaluator evaluator, int cellCounter) {

		Object obj = null;
		switch (evaluator.evaluateInCell(cell).getCellType()) {
		case Cell.CELL_TYPE_BLANK:
			return "Blank_" + cellCounter;
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
}
