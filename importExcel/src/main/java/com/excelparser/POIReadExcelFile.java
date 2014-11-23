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

import com.model.Client;

public class POIReadExcelFile {

	static {
		Map<String, Class<?>> map = new HashMap<String, Class<?>>();
		map.put("clinet", Client.class);
		
	}	
	
	public static void excelParser(Map<Integer, String> headerMapper) {
		try {

			FileInputStream excelFile = new FileInputStream(new File(
					"D:\\excels\\test.xlsx"));

			Workbook workbook = WorkbookFactory.create(excelFile);
			System.out.println("Number of sheets: "
					+ workbook.getNumberOfSheets());

			// needed to evaluate cells that contain formulas
			FormulaEvaluator evaluator = workbook.getCreationHelper()
					.createFormulaEvaluator();

			for (int i = 0; i < 1; i++) {

				Sheet sheet = workbook.getSheetAt(i);
				System.out.println("Sheet Name: " + sheet.getSheetName());
				
				int rowCounter = -1;
				
				JSONArray Jrows = new JSONArray();
				
				// Iterate through each row
				Iterator<Row> rowIterator = sheet.iterator();
				
				while (rowIterator.hasNext()) {
					
					JSONObject Jobject = new JSONObject();
					
					rowCounter++;
					int cellCounter = 0;
					Row row = rowIterator.next();
					// for each row, iterate through all the columns
					Iterator<Cell> cellIterator = row.cellIterator();

					
					while (cellIterator.hasNext()) {
						
						

						// CellReference cellReference = new CellReference

						Cell cell = cellIterator.next();
						// Check cell type and format --
						// if it's a formula cell it will be evaluated,
						// otherwise nothing happens

						switch (evaluator.evaluateInCell(cell).getCellType()) {
						case Cell.CELL_TYPE_BLANK:
							System.out.println("N/A");
							break;
						case Cell.CELL_TYPE_NUMERIC:
							System.out.println(cell.getNumericCellValue());
							break;
						case Cell.CELL_TYPE_STRING:
							System.out.println(cell.getStringCellValue());
							cellCounter++;
							if(!(rowCounter>0)) {
								loadHeaderMapper(headerMapper, cellCounter, cell.getStringCellValue().toString());
							}
							else {
								Jobject.put(headerMapper.get(cellCounter), cell.getStringCellValue().toString());
							}
							break;
						case Cell.CELL_TYPE_BOOLEAN:
							System.out.println(cell.getBooleanCellValue());

						case Cell.CELL_TYPE_FORMULA:
							break;
						}
					}
					
					
					System.out.println("");
					Jrows.add(Jobject);
				}
				System.out.println(Jrows.toJSONString());
				excelFile.close();
			}

		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public static void main(String args[]) {
		Map<Integer, String> headerMapper = new HashMap<Integer, String>();
		excelParser(headerMapper);
		System.out.println("sd");
		
	}
	
	static void loadHeaderMapper(Map<Integer, String> headerMapper, int hashIndex, String headerName) {
		if(!(headerMapper!=null)) {
			headerMapper = new HashMap<Integer, String>();
		}
		headerMapper.put(hashIndex, headerName.toLowerCase());
	}
}
