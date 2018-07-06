package net.nullpunkt.exceljson.convert;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;

import net.nullpunkt.exceljson.pojo.ExcelWorkbook;
import net.nullpunkt.exceljson.pojo.ExcelWorksheet;

import net.nullpunkt.exceljson.pojo.HeaderObject;
import net.nullpunkt.exceljson.pojo.StepObject;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFEvaluationWorkbook;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

//"dd.MM.yy"
public class ExcelToJsonConverter {
	
	private ExcelToJsonConverterConfig config = null;
	
	public ExcelToJsonConverter(ExcelToJsonConverterConfig config) {
		this.config = config;
	}
	
	public static ExcelWorkbook convert(ExcelToJsonConverterConfig config) throws InvalidFormatException, IOException {
		return new ExcelToJsonConverter(config).convert();
	}

	Workbook wb = null;
	Sheet sheet = null;
    int rowLimit;
    int startRowOffset;
    int currentRowOffset;
    int totalRowsAdded;
    ExcelWorksheet excelWorksheet;
    StepObject stepObject;
    HeaderObject headerObject;
    int lastCell = 10;

	public ExcelWorkbook convert()
			throws InvalidFormatException, IOException {

        ExcelWorkbook book = new ExcelWorkbook();
		InputStream inp = new FileInputStream(config.getSourceFile());
		wb = WorkbookFactory.create(inp);

		book.setFileName(config.getSourceFile());
		int loopLimit =  wb.getNumberOfSheets();
		if (config.getNumberOfSheets() > 0 && loopLimit > config.getNumberOfSheets()) {
			loopLimit = config.getNumberOfSheets();
		}
		rowLimit 			= config.getRowLimit();
		startRowOffset 		= config.getRowOffset();
		currentRowOffset 	= -1;
		totalRowsAdded 		= 0;


		for (int i = 0; i < loopLimit; i++) {
			sheet = wb.getSheetAt(i);
			if (sheet == null) {
				continue;
			}
            excelWorksheet = new ExcelWorksheet();
            stepObject = new StepObject();
            headerObject = new HeaderObject();
            excelWorksheet.setName(sheet.getSheetName());
            /*
            * Get rows
            * */
        	for(int j=sheet.getFirstRowNum(); j<=sheet.getLastRowNum(); j++) {
	    		Row row = sheet.getRow(j+1);
	    		if(row==null) {
	    			continue;
	    		}
	    		boolean hasValues = false;
	    		ArrayList<Object> rowData = new ArrayList<Object>();
                /*
	    		* Get Columns
	    		* */
                for(int k=0; k<=lastCell; k++) {
                    Cell cell = row.getCell(k);
	    			if(cell!=null) {
                        int type = cell.getCellType();
                        if(type == Cell.CELL_TYPE_FORMULA) {
                            evaluateFormula(cell);
                        }
                        else{
                            cell.setCellType(Cell.CELL_TYPE_STRING);
                            Object value = cellToObject(cell);
                            hasValues = hasValues || value!=null;
                            rowData.add(value);
                        }
	    			} else {
                        rowData.add(null);
                    }
                }
                if(hasValues||!config.isOmitEmpty()) {
                    createStepObject(rowData);
                }
            }
        	if(config.isFillColumns()) {
                stepObject.fillColumns();
        	}
        	/*
        	* Add all the objects to sheet
        	* */
            book.addExcelWorksheet(excelWorksheet);
		}

		return book;
	}

    private void createStepObject(ArrayList<Object> rowData) {
        HeaderObject headerObjects = header(rowData);
        stepObject.addStepRow(headerObjects);
        totalRowsAdded++;
        excelWorksheet.setData(stepObject);
    }

    private void evaluateFormula(Cell cell) {
        System.out.println(cell);
        String formula = cell.getCellFormula();
        String Name = formula.split("!")[0];
        String SheetName = Name.replace("'","");
        String range = formula.split("!")[1]; //B3:J6
        String startRange = range.split(":")[0]; //B3
        String endRange = range.split(":")[1]; //J6
        CellReference Firstref = new CellReference(startRange);
        CellReference Lastref = new CellReference(endRange);
        int startRowNumber = Integer.parseInt(Firstref.getCellRefParts()[1]);
        int endRowNumber = Integer.parseInt(Lastref.getCellRefParts()[1]);
        Sheet workingsheet = wb.getSheet(SheetName); //SheetName
        for(int i=startRowNumber;i<=endRowNumber;i++) {
            Row row = workingsheet.getRow(i-1);
            if (row == null) {
                continue;
            }
            boolean hasValues = false;
            ArrayList<Object> referenceRowData = new ArrayList<Object>();
            for (int k = 0; k <= lastCell; k++) {
                Cell cell1 = row.getCell(k);
                if (cell1 != null) {
                    cell1.setCellType(Cell.CELL_TYPE_STRING);
                    Object value = cellToObject(cell1);
                    hasValues = hasValues || value != null;
                    referenceRowData.add(value);
                } else {
                    referenceRowData.add(null);
                }
            }
            if (hasValues || !config.isOmitEmpty()) {
                createStepObject(referenceRowData);
            }
        }
    }

    private HeaderObject header(ArrayList<Object> rowData) {
        HeaderObject headerObjects = new HeaderObject();
        headerObjects.setAction((String) rowData.get(1));
        headerObjects.setApi((String) rowData.get(2));
        headerObjects.setSkipstep(Boolean.valueOf((String) rowData.get(3)));
        headerObjects.setSelector((String) rowData.get(4));
        headerObjects.setSelectorName((String) rowData.get(5));
        headerObjects.setDescription((String) rowData.get(6));
        if(rowData.get(7) != null){
            headerObjects.setOptions(((String) rowData.get(7)).split(","));
        }
        else{
            headerObjects.setOptions(null);
        }
        headerObjects.setValue((String) rowData.get(8));
        headerObjects.setExpected((String) rowData.get(9));
        return headerObjects;
    }

    private Object cellToObject(Cell cell) {

		int type = cell.getCellType();
		
		if(type == Cell.CELL_TYPE_STRING) {
			return cleanString(cell.getStringCellValue());
		}
		
		if(type == Cell.CELL_TYPE_BOOLEAN) {
			return cell.getBooleanCellValue();
		}
		
		if(type == Cell.CELL_TYPE_NUMERIC) {

			if (cell.getCellStyle().getDataFormatString().contains("%")) {
		        return cell.getNumericCellValue() * 100;
		    }

            return numeric(cell);
		}

		return null;

	}
	
	private String cleanString(String str) {
		return str.replace("\n", "").replace("\r", "");
	}

	private Object numeric(Cell cell) {
		if(HSSFDateUtil.isCellDateFormatted(cell)) {
			if(config.getFormatDate()!=null) {
				return config.getFormatDate().format(cell.getDateCellValue());	
			}
			return cell.getDateCellValue();
		}
		return String.valueOf(cell.getNumericCellValue());
	}
}
