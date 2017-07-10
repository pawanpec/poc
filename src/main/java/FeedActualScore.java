import java.io.BufferedReader;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStreamReader;
import java.net.URL;
import java.net.URLConnection;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.TimeZone;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFEvaluationWorkbook;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.formula.SharedFormula;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.json.simple.JSONArray;
import org.json.simple.JSONObject;
import org.json.simple.parser.JSONParser;


public class FeedActualScore {
	public static final String urlString = "https://devapi.insight360.io/v3/data/companies/US30303M1027/series?start_date=2016-04-01&end_date=2016-04-30&metrics=allmetrics&score_type=pulse&token=eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJpc3MiOiJodHRwczovL2RldmF1dGguaW5zaWdodDM2MC5pbyIsInN1YiI6ImZha2VfaW50ZXJuYWxfdHZsX2FwaV91c2VyQHRydXZhbHVlbGFicy5jb20iLCJleHAiOiIyMDE4LTA1LTA5VDEwOjE3OjM1LjQwMFoiLCJpYXQiOiIyMDE3LTA1LTA5VDEwOjE3OjM1LjQwMFoiLCJuYW1lIjoiSW50ZXJuYWwgQXBpIFVzZXIiLCJlbmMiOiJlMDRmYjE5ZDQwZDc5OTMzZDBiMzdkNmZjNGEzYzAzN2Q1NDVlY2MyOTdjZTY5Y2VmNzZhMzk2NGQzN2FjMGI1YmYyMDYzNTUyMDY1MjA4MWRmYzRkOGY2ZDU2OGY3ZWQ1ODliMDIyODcwOGI2MDk1ODEyYmQ4Yzg2NDJmZDZiYjAwZWY5MDNlMzQ5MjZhMzM1MTRhZjRiZjBiMDY5NTMyYmM4ZmZiNjNjZWM1ZGEyZWRjMDgwMjZlMDhlOTBjZGNkOWU3NTE0NmJmOGNiNmE5NGRlYzIxOTgyMGU4ZDRlYTc2NjU2NmZjNDkxYmY3OGNhYjk1YjU0YmNmMjM0ZGJjOTAyMDhlMTBhODFjY2NjN2UyYjQ2ODhhZTYzMDM1OWIyYmRjMjViZTAxNDZhYzFiMDhkNTdiZWQ4MjZiYWQzMmRiNjVjNDk4MTRhZmI4MjhmM2UxYzQ1NTlhMzhjMzA2ZDI0MmY3NGRjYmM3OTgxOTE5N2ZkOWNlNjM1MjNjZWJmMmU2YWU1ZjkxNTQxN2I4MTIwNWViZjAzYmRjYzM3OTU2OGM0NTk2YjBhNTdiZTBjYjNiNDRiOTIyNGZlYzg5MmRlMWVlMDhhYTlhOWUxODljNTBkZGFkN2EwNmIzNDVlNDFhNGYyMTgxODMwMWI1ODUzZjYxZmU2ZmU2MWM5NDllOTQ0OTQzZWQ4ZWY1OWZhZjc1YzE4N2I0In0.ViYLzDvwPysQsjZrWTeItwX74xVzVhdByIdTRdTU748";

	public static void main(String[] args) {
		handleAPIResponse(); 

	}
	
	
	//function to write actual score from the API. TTN
	public static void handleAPIResponse() {
	    
	    JSONObject apiResult = getDataFromAPI(urlString);
	    
		int currentRow = 12;
	    FileInputStream inp = null;
	    FileOutputStream output_file = null;
	    JSONObject jsonObject=null;
	    HSSFWorkbook workbook;
	    
	    try {            
	        inp = new FileInputStream("Insight_test.xls");
	        workbook = new HSSFWorkbook(inp);
			HSSFEvaluationWorkbook formulaParsingWorkbook = HSSFEvaluationWorkbook.create((HSSFWorkbook) workbook);
			SharedFormula sharedFormula = new SharedFormula(SpreadsheetVersion.EXCEL2007);
			FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();	        
	        // get all the feilds from excel file
            HSSFSheet sheet = workbook.getSheetAt(0);
	        Row row = sheet.getRow(8);
	        List tvl2CatsKeys = new ArrayList();
	        
	        if(row != null) {
	        	for(Cell cell : row) {
	        		if (cell.getCellType() == Cell.CELL_TYPE_STRING && cell.getRichStringCellValue().getString().trim() != "") {
	        			if(cell.getColumnIndex() != 0 && cell.getColumnIndex() != 1) {			
	        			tvl2CatsKeys.add(cell.getRichStringCellValue().getString().trim());
	                }
	        	}
	          }
	        }
	        JSONArray jsonarray =  (JSONArray)apiResult.get("series");        
	        
	        for(Object obj : jsonarray) {
	        	jsonObject = null;
	        	jsonObject = (JSONObject) obj;
	        	
				//handleRecord(sheet,currentRow,jsonObject,formulaParsingWorkbook, sharedFormula,evaluator,tvl2CatsKeys);
	        	handleAPIJSONData(sheet,currentRow,jsonObject,tvl2CatsKeys);
				currentRow++;
	        }

	        inp.close();
	    	output_file =new FileOutputStream("Insight_test.xls");  
	    	 //write changes
	    	workbook.write(output_file);
	    	//close the stream
	    	output_file.close();
	    	
	    } catch (IOException ex) {
	    	ex.printStackTrace();
	    }

	}
	public static void handleAPIJSONData(Sheet sheet,int currentRow,JSONObject jsonObject,List tvl2CatsKeys) {
		
		
		boolean flag = false;
		Object date = jsonObject.get("date");
		Object hourMs = jsonObject.get("hourMs");
		JSONObject pulse = (JSONObject)jsonObject.get("pulse");
		//System.out.println(pulse);
	    

	    for (Object key : pulse.keySet()) {
	        //based on you key types
	        String keyStr = (String)key;
	        Object keyvalue = pulse.get(keyStr);	    
	    }
	    
	    Iterator<Row> shetIter = sheet.iterator();
	    int rowCount = 0;
		while(shetIter.hasNext()){
			Row row = shetIter.next();
			
			if(rowCount++ < 12) {
				continue;
			}
			  Cell dateCell = row.getCell(0);
			
			  Date sDate1= dateCell.getDateCellValue();
			  DateFormat formatter = new SimpleDateFormat("dd-MM-yyyy");
			  formatter.setTimeZone(TimeZone.getTimeZone("UTC"));

			  
			  try {
				 Date todayWithZeroTime = formatter.parse(formatter.format(sDate1));
				  
				 long millis = todayWithZeroTime.getTime();
				 System.out.println(millis+ "   "+ todayWithZeroTime );
				 
			  } catch(ParseException ex) {
				 ex.printStackTrace();
			  }
			  

			  
			  
			  
//			  System.out.println(sDate1.);
			    /*Date date1;
				try {
					date1 = new SimpleDateFormat("dd-MMM-yyyy").parse(sDate1);
					System.out.println(sDate1+"\t"+date1);
				} catch (ParseException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}  */
			      
		}
			//= dateCell.getDateCellValue();
		    //SimpleDateFormat sdf = new SimpleDateFormat("yyyy/MM/dd HH:mm:ss");
		
	}
	
	//TTN
	@SuppressWarnings("finally")
	public static JSONObject getDataFromAPI(String urlString){
		StringBuilder result = new StringBuilder();	
		JSONObject object = null;
		try{
			URL url = new URL(urlString);
			URLConnection conn = url.openConnection();
		      
			BufferedReader rd = new BufferedReader(new InputStreamReader(conn.getInputStream()));
			String line;
			while ((line = rd.readLine()) != null) {
				result.append(line);
			}
			
			rd.close();
			
	        JSONParser jsonParser=new JSONParser();
	        object = (JSONObject)jsonParser.parse(result.toString());	         
			
		} catch(Exception ex) {
			ex.printStackTrace();
		} finally {
			return object;
		}
	}


}
