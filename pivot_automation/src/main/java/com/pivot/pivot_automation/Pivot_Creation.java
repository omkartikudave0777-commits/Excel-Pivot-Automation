/**
 * Pivot_Creation
 *
 * This class contains utility methods to create pivot tables in an Excel workbook using Apache POI.
 *
 * Functions:
 * - createPivotTable(XSSFSheet sheet, XSSFWorkbook workbook, List<Map<String, Object>> dataList):
 *   Creates pivot tables based on the provided configuration (dimensions, matrix operations, filters),
 *   and adds them as new sheets in the given workbook.
 *
 * - getColumnIndexByName(XSSFSheet sheet, String columnName):
 *   Returns the column index of the specified column name from the header row of the given sheet.
 */

package com.pivot.pivot_automation;

import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFPivotTable;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.*;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDataFields;

import java.util.List;
import java.util.Map;
import org.apache.poi.ss.usermodel.*;


public class Pivot_Creation {
	
	public static XSSFWorkbook createPivotTable(XSSFSheet sheet, XSSFWorkbook workbook, List<Map<String, Object>> dataList) {

		/*
		 Validate inputs to avoid null errors and unnecessary processing.
		 Returns early if sheet, workbook, or dataList is missing or empty.
		*/
		
        if (sheet == null || workbook == null || dataList == null || dataList.isEmpty()) 
        {
            System.out.println("Invalid input: sheet, workbook, or dataList is null/empty.");
            return workbook;
        }
        
        // Defines the data range for the pivot table using the top-left and bottom-right cell coordinates of the source sheet.
        AreaReference sourceData = new AreaReference(
                new CellReference(0, 0),
                new CellReference(sheet.getLastRowNum(), sheet.getRow(0).getLastCellNum() - 1),
                workbook.getSpreadsheetVersion());

        // Create pivot sheets based on each entry in dataList
        for (Map<String, Object> entry : dataList) 
        {
            String pivotSheetName = (String) entry.get("sheet_name");
            System.out.println();
            System.out.println("Pivot sheet name "+pivotSheetName);
            if (pivotSheetName == null || pivotSheetName.trim().isEmpty()) 
            {
                System.out.println("Skipping entry due to missing 'sheet_name'.");
                continue;
            }

            // Skip if sheet already exists
            if (workbook.getSheet(pivotSheetName) != null) 
            {
                System.out.println("  Sheet '" + pivotSheetName + "' already exists. Skipping creation.");
                continue;
            }

            // Create sheet and pivot table
            XSSFSheet pivotSheet = workbook.createSheet(pivotSheetName);
            Row headerRow = sheet.getRow(0);
            int totalCells = headerRow.getLastCellNum();
            String[] headers = new String[totalCells];
            
            XSSFPivotTable pivotTable = pivotSheet.createPivotTable(sourceData, new CellReference("A5"), sheet);
            int numberOfColumns = 13; // Adjust the number of columns to set the width for
            
            for (int columnIndex = 0; columnIndex < numberOfColumns; columnIndex++) 
            {
                int columnWidth = 15 * 256; // Here deciding the Width for the columns for an entire sheet where we will bw
                                            // creating pivot (15 characters in width)
                pivotSheet.setColumnWidth(columnIndex, columnWidth);
            }
            
            List<String> dimensions = (List<String>) entry.get("dimension");
            for (String dim : dimensions) 
            {
                for (int i = 0; i < totalCells; i++) 
                {
                    headers[i] = headerRow.getCell(i).getStringCellValue();
                    String headerName = headers[i];

                    if (dim.equals(headerName)) 
                    {
                        pivotTable.addRowLabel(i);
                        break;
                    }
                }
            }
            System.out.println("Adding row labels from 'dimension' block...");

            List<Map<String, String>> matrixList = (List<Map<String, String>>) entry.get("matrix");
            
            for (Map<String, String> matrix : matrixList) 
            {
                String columnName = matrix.get("column");
                String operation = matrix.get("operation").toLowerCase();

                int colIndex = getColumnIndexByName(sheet, columnName);

                if (colIndex == -1) 
                {
                    System.out.println("Column not found: " + columnName);
                    continue;
                }

                switch (operation) 
                {
                    case "count":
                        pivotTable.addColumnLabel(DataConsolidateFunction.COUNT, colIndex, "Count of " + columnName);
                        break;

                    case "sum":
                        pivotTable.addColumnLabel(DataConsolidateFunction.SUM, colIndex, "Sum of " + columnName);
                        break;

                    case "avg":
                        pivotTable.addColumnLabel(DataConsolidateFunction.AVERAGE, colIndex, "Avg of " + columnName);
                        break;

                    case "column":  // Custom tag for pivot column fields like "Date"
                        pivotTable.addColLabel(colIndex);
                        break;

                    case "row":     // Optional: support row labels too
                        pivotTable.addRowLabel(colIndex);
                        break;

                    default:
                        System.out.println("Unknown operation: " + operation);
                }
            }
            System.out.println("Adding data fields from 'matrix' block...");

            List<String> filters = (List<String>) entry.get("Filter");
            if(filters != null && !filters.isEmpty())
            {
	            for (String filterColumn : filters) 
	            {
	                int colIndex = getColumnIndexByName(sheet, filterColumn);
	                if (colIndex != -1) 
	                {
	                    pivotTable.addReportFilter(colIndex);
	                }
	            }
                System.out.println("Adding filters from 'Filter' block...");
            }
            else 
            {
            	System.out.println("Filter is not used !");
            }
            
        }

        return workbook;
    }
	
	public static int getColumnIndexByName(XSSFSheet sheet, String columnName) {
		
	    XSSFRow headerRow = sheet.getRow(0);
	    if (headerRow != null) 
	    {
	        for (int i = 0; i < headerRow.getLastCellNum(); i++) 
	        {
	            if (headerRow.getCell(i).getStringCellValue().equalsIgnoreCase(columnName)) 
	            {
	                return i;
	            }
	        }
	    }
	    return -1; // Not found
	}

}
