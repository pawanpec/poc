import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Date;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFEvaluationWorkbook;
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


public class ReadExcelWithFormula {
	static int startColumn=1,endColumn=130;

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
	        evaluator.clearAllCachedResultValues();
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
	
	private static int searchKey(Sheet sheet, int rowNumber, String searchValue) {
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
			sheet.getRow(row).getCell(column, Row.CREATE_NULL_AS_BLANK).setCellValue("");
			sheet.getRow(row).getCell(column+1, Row.CREATE_NULL_AS_BLANK).setCellValue(0);
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