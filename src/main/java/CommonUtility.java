import java.io.BufferedReader;
import java.io.InputStreamReader;
import java.net.URL;
import java.net.URLConnection;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.json.simple.JSONArray;
import org.json.simple.JSONObject;
import org.json.simple.parser.JSONParser;

public class CommonUtility {
	public static final String urlString = "https://devapi.insight360.io/v3/data/companies/US30303M1027/series?start_date=2016-04-01&end_date=2016-04-30&metrics=allmetrics&score_type=pulse&token=eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJpc3MiOiJodHRwczovL2RldmF1dGguaW5zaWdodDM2MC5pbyIsInN1YiI6ImZha2VfaW50ZXJuYWxfdHZsX2FwaV91c2VyQHRydXZhbHVlbGFicy5jb20iLCJleHAiOiIyMDE4LTA1LTA5VDEwOjE3OjM1LjQwMFoiLCJpYXQiOiIyMDE3LTA1LTA5VDEwOjE3OjM1LjQwMFoiLCJuYW1lIjoiSW50ZXJuYWwgQXBpIFVzZXIiLCJlbmMiOiJlMDRmYjE5ZDQwZDc5OTMzZDBiMzdkNmZjNGEzYzAzN2Q1NDVlY2MyOTdjZTY5Y2VmNzZhMzk2NGQzN2FjMGI1YmYyMDYzNTUyMDY1MjA4MWRmYzRkOGY2ZDU2OGY3ZWQ1ODliMDIyODcwOGI2MDk1ODEyYmQ4Yzg2NDJmZDZiYjAwZWY5MDNlMzQ5MjZhMzM1MTRhZjRiZjBiMDY5NTMyYmM4ZmZiNjNjZWM1ZGEyZWRjMDgwMjZlMDhlOTBjZGNkOWU3NTE0NmJmOGNiNmE5NGRlYzIxOTgyMGU4ZDRlYTc2NjU2NmZjNDkxYmY3OGNhYjk1YjU0YmNmMjM0ZGJjOTAyMDhlMTBhODFjY2NjN2UyYjQ2ODhhZTYzMDM1OWIyYmRjMjViZTAxNDZhYzFiMDhkNTdiZWQ4MjZiYWQzMmRiNjVjNDk4MTRhZmI4MjhmM2UxYzQ1NTlhMzhjMzA2ZDI0MmY3NGRjYmM3OTgxOTE5N2ZkOWNlNjM1MjNjZWJmMmU2YWU1ZjkxNTQxN2I4MTIwNWViZjAzYmRjYzM3OTU2OGM0NTk2YjBhNTdiZTBjYjNiNDRiOTIyNGZlYzg5MmRlMWVlMDhhYTlhOWUxODljNTBkZGFkN2EwNmIzNDVlNDFhNGYyMTgxODMwMWI1ODUzZjYxZmU2ZmU2MWM5NDllOTQ0OTQzZWQ4ZWY1OWZhZjc1YzE4N2I0In0.ViYLzDvwPysQsjZrWTeItwX74xVzVhdByIdTRdTU748";
    
	public static void main(String[] args) {
    	JSONArray jsonArray=getDataFromAPI(urlString);
    	JSONObject jsonObject=getValue(jsonArray, 1459468800000l);
		System.out.println(jsonObject);
	}
	public static JSONArray getDataFromAPI(String urlString) {
		String apiResponse = getAPIRespose(urlString);
		JSONObject jsonObject = null;
		try {
			JSONParser jsonParser = new JSONParser();
			jsonObject = (JSONObject) jsonParser.parse(apiResponse);

		} catch (Exception ex) {
			ex.printStackTrace();
		}
		return (JSONArray) jsonObject.get("series");
	}
	public static JSONObject getValue(JSONArray array, long key)
	{
		JSONObject value = null;
	    for (int i = 0; i < array.size(); i++)
	    {
	    	JSONObject jsonObject = (JSONObject) array.get(i);
	        Long hourMs=(Long) jsonObject.get("hourMs");
	        if (hourMs.equals(key))
	        {
	        	value=jsonObject;
	            break;
	            
	        }
	    }

	    return value;
	}
	public static void updateCurrentRowActualScore(int currentrow, JSONObject currentObj, HSSFSheet sheet) {
        		
		//get row to update which is a key in map
		//iterate over pulse score object and update the pulse score in actual field
		
		for(Iterator iterator = currentObj.keySet().iterator(); iterator.hasNext();) {
		    String key = (String) iterator.next();
		    String value = currentObj.get(key).toString();
		    
		    int cellNumber = ReadExcelWithFormula.searchKey(sheet, 8, key);
		    System.out.println("writing data to " + currentrow);
		    
		    setcellValue(sheet, currentrow, cellNumber + 2, Double.parseDouble(value));		   		    
		}
		        
	}
	public static void setcellValue(Sheet sheet, int row, int column,Double value) {	
		if(sheet.getRow(row) == null) {
			sheet.createRow(row);
		}
		if(value!=null){
			sheet.getRow(row).getCell(column, Row.CREATE_NULL_AS_BLANK).setCellValue(value);
			System.out.println(sheet.getRow(row).getCell(column, Row.CREATE_NULL_AS_BLANK));

		}		 
	}
	
	public static String getAPIRespose(String urlString) {
		StringBuilder result = new StringBuilder();
		try {
			URL url = new URL(urlString);
			URLConnection conn = url.openConnection();

			BufferedReader rd = new BufferedReader(new InputStreamReader(conn.getInputStream()));
			String line;
			while ((line = rd.readLine()) != null) {
				result.append(line);
			}

			rd.close();
		} catch (Exception ex) {
			ex.printStackTrace();
		}
		return result.toString();
	}

}
