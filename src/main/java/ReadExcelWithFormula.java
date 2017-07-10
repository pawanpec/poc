import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Comparator;
import java.util.Date;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.TimeZone;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFEvaluationWorkbook;
import org.apache.poi.hssf.usermodel.HSSFFormulaEvaluator;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.formula.FormulaParser;
import org.apache.poi.ss.formula.FormulaRenderer;
import org.apache.poi.ss.formula.FormulaType;
import org.apache.poi.ss.formula.SharedFormula;
import org.apache.poi.ss.formula.ptg.Ptg;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.json.simple.JSONArray;
import org.json.simple.JSONObject;

import com.sun.image.codec.jpeg.TruncatedFileException;


public class ReadExcelWithFormula {
	static int startColumn=1,endColumn=130;
	public static final String urlString = "https://devapi.insight360.io/v3/data/companies/US30303M1027/series?start_date=2016-04-01&end_date=2016-04-30&metrics=allmetrics&score_type=pulse&token=eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJpc3MiOiJodHRwczovL2RldmF1dGguaW5zaWdodDM2MC5pbyIsInN1YiI6ImZha2VfaW50ZXJuYWxfdHZsX2FwaV91c2VyQHRydXZhbHVlbGFicy5jb20iLCJleHAiOiIyMDE4LTA1LTA5VDEwOjE3OjM1LjQwMFoiLCJpYXQiOiIyMDE3LTA1LTA5VDEwOjE3OjM1LjQwMFoiLCJuYW1lIjoiSW50ZXJuYWwgQXBpIFVzZXIiLCJlbmMiOiJlMDRmYjE5ZDQwZDc5OTMzZDBiMzdkNmZjNGEzYzAzN2Q1NDVlY2MyOTdjZTY5Y2VmNzZhMzk2NGQzN2FjMGI1YmYyMDYzNTUyMDY1MjA4MWRmYzRkOGY2ZDU2OGY3ZWQ1ODliMDIyODcwOGI2MDk1ODEyYmQ4Yzg2NDJmZDZiYjAwZWY5MDNlMzQ5MjZhMzM1MTRhZjRiZjBiMDY5NTMyYmM4ZmZiNjNjZWM1ZGEyZWRjMDgwMjZlMDhlOTBjZGNkOWU3NTE0NmJmOGNiNmE5NGRlYzIxOTgyMGU4ZDRlYTc2NjU2NmZjNDkxYmY3OGNhYjk1YjU0YmNmMjM0ZGJjOTAyMDhlMTBhODFjY2NjN2UyYjQ2ODhhZTYzMDM1OWIyYmRjMjViZTAxNDZhYzFiMDhkNTdiZWQ4MjZiYWQzMmRiNjVjNDk4MTRhZmI4MjhmM2UxYzQ1NTlhMzhjMzA2ZDI0MmY3NGRjYmM3OTgxOTE5N2ZkOWNlNjM1MjNjZWJmMmU2YWU1ZjkxNTQxN2I4MTIwNWViZjAzYmRjYzM3OTU2OGM0NTk2YjBhNTdiZTBjYjNiNDRiOTIyNGZlYzg5MmRlMWVlMDhhYTlhOWUxODljNTBkZGFkN2EwNmIzNDVlNDFhNGYyMTgxODMwMWI1ODUzZjYxZmU2ZmU2MWM5NDllOTQ0OTQzZWQ4ZWY1OWZhZjc1YzE4N2I0In0.ViYLzDvwPysQsjZrWTeItwX74xVzVhdByIdTRdTU748";

	public static void main(String args[]) {
		int currentRow = 12;
	    FileInputStream inp = null;
	    FileOutputStream output_file = null;
	    JSONObject jsonObject=null;
	    HSSFWorkbook workbook;
	    
	    try {            
	        inp = new FileInputStream("input_pulse.xls");
	        workbook = new HSSFWorkbook(inp);
			HSSFEvaluationWorkbook formulaParsingWorkbook = HSSFEvaluationWorkbook.create((HSSFWorkbook) workbook);
			SharedFormula sharedFormula = new SharedFormula(SpreadsheetVersion.EXCEL2007);
			FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
	        JSONArray jsonarray = ReadInputData.getInputData();
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
	        
	        for(Object obj : jsonarray) {
	        	jsonObject = null;
	        	jsonObject = (JSONObject) obj; 
				handleRecord(sheet,currentRow,jsonObject,formulaParsingWorkbook, sharedFormula,evaluator,tvl2CatsKeys);
				currentRow++;
	        }
	        int totalRows=12+jsonarray.size() ;
	        for (int i = 0; i < tvl2CatsKeys.size(); i++) {
				int cellNumber = searchKey(sheet, 8, (String) tvl2CatsKeys.get(i));
				// raw value will be found (or added) at row : currentRow and
				// column : cellNumber -2
				Cell c = sheet.getRow(8).getCell(cellNumber - 2, Row.CREATE_NULL_AS_BLANK);
				String formula = c.toString();
				if (formula.length() > 3) {
					String newFormula = formula.substring(0, formula.length() - 3) +totalRows + ")";
					c.setCellType(HSSFCell.CELL_TYPE_FORMULA);
					c.setCellFormula(newFormula);
				}

			}
	        Map inputData=new LinkedHashMap<Integer, Date>();
	        for (int i = 11; i <totalRows-1; i++) {
	        	Cell cell1 = sheet.getRow(i).getCell(0, Row.CREATE_NULL_AS_BLANK);
	        	Cell cell2 = sheet.getRow(i+1).getCell(0, Row.CREATE_NULL_AS_BLANK);
	        	Date d1 = cell1.getDateCellValue();
				Date d2 = cell2.getDateCellValue();
	        	if (d1!=null&&d2!=null) {
	        		d1=truncateToDay(d1);
	        		d2=truncateToDay(d2);
					if (d1.getTime()==d2.getTime()) {
						continue;
					} else {
						inputData.put(i, d1.getTime());
					} 
				}else{
					if (d1!=null) {
						d1=truncateToDay(d1);
						inputData.put(i, d1.getTime());
					}
				}
	        }
	        HSSFFormulaEvaluator.evaluateAllFormulaCells(workbook);
	        
	        //fetch api data
	        JSONArray  jsonArr = CommonUtility.getDataFromAPI(urlString);
	        
	        //iterate over map	        
	        Iterator<Map.Entry<Integer, Long>> mapIterator = inputData.entrySet().iterator();
	        
	        while (mapIterator.hasNext()) {
	            Map.Entry<Integer, Long> pairObj = mapIterator.next();
	            
	            //get value of date from map
	            JSONObject currentObj = CommonUtility.getValue(jsonArr, pairObj.getValue());

	            //get pulse score for that date
	            JSONObject pulseObj = (JSONObject)currentObj.get("pulse");
	        
	            //get row to update which is a key in map	        
		        //iterate over pulse score object and update the pulse score in actual field
	            CommonUtility.updateCurrentRowActualScore(pairObj.getKey(),pulseObj,sheet);
	        }
	        
	        	        
	        inp.close();
	    	output_file =new FileOutputStream("output.xls");  
	    	 //write changes
	    	workbook.write(output_file);
	    	//close the stream
	    	output_file.close();
	    	
	    } catch (IOException ex) {
	    	ex.printStackTrace();
	    }
	}
	 private static Date truncateToDay(Date date) {
         Calendar calendar = Calendar.getInstance();
         calendar.setTime(date);
         calendar.set(Calendar.HOUR_OF_DAY, 0);
         calendar.set(Calendar.MINUTE, 0);
         calendar.set(Calendar.SECOND, 0);
         calendar.set(Calendar.MILLISECOND, 0);
         return calendar.getTime();
     }

	private static void handleRecord(HSSFSheet sheet,int currentRow,JSONObject jsonObject,HSSFEvaluationWorkbook 
			formulaParsingWorkbook, SharedFormula sharedFormula,FormulaEvaluator evaluator,List tvl2CatsKeys) {
		    	
		// set date in 1st column : timestamp from JSON and add other values from JSON
		handleCellFromJson(sheet,currentRow-1,jsonObject,tvl2CatsKeys);
		
		// handle all the cell with formula
		for(int i=startColumn;i<=endColumn;i++) {
			handleFormulaCell(sheet,currentRow,i,formulaParsingWorkbook,sharedFormula,evaluator);	
		}
	}

	private static void handleCellFromJson(Sheet sheet,int currentRow,JSONObject jsonObject,List tvl2CatsKeys) {
		Long value = 0l;
		Double modelValue;
				
		// set date in 1st column : timestamp from JSON
		if(jsonObject.get("articlePubDateMs") instanceof Double) {
			value = new Double((Double) jsonObject.get("articlePubDateMs")).longValue();
		} else if(jsonObject.get("articlePubDateMs") instanceof Long) {
			value = (Long) jsonObject.get("articlePubDateMs");
		} 

		if(sheet.getRow(currentRow) == null) {
			sheet.createRow(currentRow);
    	 }
    	 sheet.getRow(currentRow).getCell(0, Row.CREATE_NULL_AS_BLANK).setCellValue(new Date(value)); // set A1=2
					
		//for tvl2Cats values
    	 JSONObject newValuesTvl2Cats = (JSONObject)jsonObject.get("tvl2Cats");
    	    
    	 for(int i = 0;i<tvl2CatsKeys.size();i++){
    		 if(newValuesTvl2Cats.get((String) tvl2CatsKeys.get(i)) instanceof Long){
    			modelValue = Double.longBitsToDouble((Long)newValuesTvl2Cats.get((String) tvl2CatsKeys.get(i)));
    		 } else {
    			 modelValue = (Double) newValuesTvl2Cats.get((String) tvl2CatsKeys.get(i));	 
    		 }
       		     // search this key in excel file
                 int cellNumber = searchKey(sheet,8,(String)tvl2CatsKeys.get(i));
                 //raw value will be found (or added) at row : currentRow and column : cellNumber -2
               	 setcellValue(sheet,currentRow, cellNumber-2,modelValue);
    		
    	 }
	}
	
	public static int searchKey(Sheet sheet, int rowNumber, String searchValue) {
		int value = -1;
		Row row = sheet.getRow(rowNumber);
		if(row !=null) {
			for(Cell cell : row) {
				if (cell.getCellType() == Cell.CELL_TYPE_STRING && cell.getRichStringCellValue().getString().trim().equals(searchValue)) {
					value = cell.getColumnIndex();
					break;
                }
			}
		}
		return value;
	}
	
	private static void handleFormulaCell(HSSFSheet sheet,int currentRow,int column,
			HSSFEvaluationWorkbook formulaParsingWorkbook, SharedFormula sharedFormula,FormulaEvaluator evaluator) {
			
		if(sheet.getRow(currentRow) == null) {
			sheet.createRow(currentRow);
		}	
		Cell sourceCell = sheet.getRow(currentRow-1).getCell(column,Row.CREATE_NULL_AS_BLANK);
		Cell currentCell = sheet.getRow(currentRow).getCell(column,Row.CREATE_NULL_AS_BLANK);
	
		// check if source cell has formula and current cell does not have formula, then, copy formula.
		if(sourceCell.getCellType() == HSSFCell.CELL_TYPE_FORMULA && currentCell.getCellType() != HSSFCell.CELL_TYPE_FORMULA){
			currentCell.setCellType(HSSFCell.CELL_TYPE_FORMULA);
			copyCellFormula(formulaParsingWorkbook,sharedFormula,sourceCell,currentCell);
		} 	
		handleCell(currentCell.getCellType(), currentCell,evaluator);
	}
	
	private static void setcellValue(Sheet sheet, int row, int column,Double value) {	
		if(sheet.getRow(row) == null) {
			sheet.createRow(row);
		}
		if(value!=null){
			sheet.getRow(row).getCell(column, Row.CREATE_NULL_AS_BLANK).setCellValue(value);
		}else if(row==11){
			sheet.getRow(row).getCell(column, Row.CREATE_NULL_AS_BLANK).setCellType(Cell.CELL_TYPE_BLANK);
		}
		 // set A1=2
	}
	
	private static void copyCellFormula(HSSFEvaluationWorkbook formulaParsingWorkbook,SharedFormula sharedFormula,
			Cell source, Cell destination){
	    Ptg[] sharedFormulaPtg = FormulaParser.parse(source.getCellFormula(), formulaParsingWorkbook, FormulaType.CELL, 0);
	    Ptg[] convertedFormulaPtg = sharedFormula.convertSharedFormulas(sharedFormulaPtg, 1, 0);
	    destination.setCellFormula(FormulaRenderer.toFormulaString(formulaParsingWorkbook, convertedFormulaPtg));
	}
	
	private static void handleCell(int type,Cell cell,FormulaEvaluator evaluator) {
	    if (type == HSSFCell.CELL_TYPE_STRING) {
	      System.out.println(cell.getStringCellValue());
	    } else if (type == HSSFCell.CELL_TYPE_NUMERIC) {
	       System.out.println(cell.getNumericCellValue());
	    } else if (type == HSSFCell.CELL_TYPE_BOOLEAN) {
	       System.out.println(cell.getBooleanCellValue());
	    } else if (type == HSSFCell.CELL_TYPE_FORMULA) {
	    	evaluator.evaluateFormulaCell(cell);
	        handleCell(cell.getCachedFormulaResultType(), cell, evaluator);
	    } else {
	       System.out.println("");
	    }
	}	
}