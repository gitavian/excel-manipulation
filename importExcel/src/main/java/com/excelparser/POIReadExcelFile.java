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
import org.codehaus.jackson.map.JsonMappingException;
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

	public static void main(String args[]) {
		Map<Integer, String> headerMapper = new HashMap<Integer, String>();
		JsonMapper jMapper = new JsonMapper();
		JSONObject excelFileToJson = jMapper.excelParser(headerMapper);
		
		Gson gson = new GsonBuilder().setPrettyPrinting().create();
		System.out.println(gson.toJson(excelFileToJson));
		System.out.println("END");
		
		int size = excelFileToJson.size();
		
		JSONArray sheet1 = (JSONArray) excelFileToJson.get("Sheet1");
		String row1 = sheet1.get(0).toString();
		
		//String clientList = gson.toJson(sheet1, Client.class);
		
		Client c = gson.fromJson(row1, Client.class);
		System.out.println(c.getBalance());
		
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
	}

	
}
