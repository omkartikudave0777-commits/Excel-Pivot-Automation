/**
 * Pivot Automation App
 *
 * Purpose:
 * This application reads an Excel workbook and a JSON configuration, then automatically
 * generates pivot tables for each sheet based on the provided configuration.
 *
 * Inputs:
 * - args[0]: Path to the input Excel (.xlsx) file
 * - args[1]: JSON string describing the pivot configuration
 * - args[2]: Output Filename with full path
 *            Format:
 *            '{
			    "house_data": 
			    [
			        {
			            "sheet_name": "pivot_house_data",
			            "dimension": ["Unit","Camp Name"],
			            "matrix": [{"column": "Date","operation": "column"},{"column": "House Id","operation": "Count"}],
			            "Filter": ["Sub Group"]
			        }
			    ]
			   }'
 *
 * Output:
 * - A new Excel file with all pivot tables added as new sheets
 *
 * Usage Example (Command line):
 * java -jar excel-pivot-automation.jar /path/to/data.xlsx '{"pivot_house_data":[{...}]}'
 * 
 * 
 * Command to execute java code through terminal : 
 * java -cp /home/yukta-13/Desktop/omkar/project/Malyalam_manorama/Pivot_project/mm-pivot-automation/mm_pivot/target/mm_pivot-0.0.1-SNAPSHOT.jar com.yukta.pivot.mm_pivot.App "/home/yukta-13/Desktop/omkar/project/Malyalam_manorama/mm_data_processing/promoter_usage_report.xlsx" '{
        "house_data": 
        [
            {
                "sheet_name": "summary",
                "dimension": ["Unit","Camp Name"],
                "matrix": [
                { "column": "Date", "operation": "column" },
        { "column": "House Id", "operation": "Count" }
    ],
        "Filter" : ["Sub Group"]
            }
        ]
    }' "/home/yukta-13/Desktop/omkar/project/Malyalam_manorama/mm_data_processing/Pivot_2_promoter_usage_report.xlsx"
 */

package com.pivot.pivot_automation;

import com.fasterxml.jackson.databind.ObjectMapper;
import java.util.*;
import java.util.Map;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFPivotTable;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.util.IOUtils;

import com.pivot.*;

public class App {
	public static void main(String[] args) throws Exception {
		
		System.out.println("=== Project execution started at " + java.time.LocalDateTime.now() + " ===");
		
		System.out.println();
		
		// This line increases the max byte array size in Apache POI to handle large Excel files with embedded data like images or charts.
		IOUtils.setByteArrayMaxOverride(600_000_000);
		
		// -------------------- Declaration Section --------------------
		// Initialize input/output file paths, JSON parser, Excel workbook, and list to track sheet names
        String input_filename = args[0];
        String jsonInput = args[1];
        String outputFilePath = args[2];
        
        XSSFWorkbook result_workbook=new XSSFWorkbook();
        ObjectMapper mapper = new ObjectMapper();
      	List<String> sheetNames = new ArrayList<>();
        Pivot_Creation pivot_object = new Pivot_Creation();
      	
        // Load JSON to jsonMap variable
        Map<String, Object> jsonMap = mapper.readValue(jsonInput, Map.class);
      	   
        // Read the input Excel file and load it into a workbook
  	    FileInputStream fis = new FileInputStream(input_filename);
  	    System.out.println("File read successfully completed");
  	    XSSFWorkbook workbook = new XSSFWorkbook(fis);
    	int numberOfSheets = workbook.getNumberOfSheets();
    	// For loop to store sheetnames into List
        for (int i = 0; i < numberOfSheets; i++) 
        {
            sheetNames.add(workbook.getSheetName(i));
        }
        
        /* 
         This block matches each sheet name from sheetNames with the sheets in the workbook and the JSON keys.
         If a match is found, it creates a pivot table using the matched sheet and JSON data.
        */
         for (String raw_sheet : sheetNames) 
         {
            for (int i = 0; i < workbook.getNumberOfSheets(); i++) 
            {
	            XSSFSheet sheet = workbook.getSheetAt(i);
	   	        String sheetName = sheet.getSheetName();
	   	        if(raw_sheet.equals(sheetName))
	   	        {
			   	     for (Map.Entry<String, Object> topEntry : jsonMap.entrySet()) 
			   	     {
				            String key = topEntry.getKey();
				            if(raw_sheet.equals(key))
				   	        {
				            	System.out.println("Creating pivot for sheet: " + raw_sheet);
					            List<Map<String, Object>> dataList = (List<Map<String, Object>>) topEntry.getValue();
					            System.out.println("Data for pivot creation: " + dataList);
				                result_workbook = pivot_object.createPivotTable(sheet, workbook, dataList);
				                continue;
				   	        }
			   	     }
			   	     continue;
	   	        }	     
            }
         }

         /*
          This block writes the result_workbook to the specified output file using a FileOutputStream.
		  If an IOException occurs during writing, it prints an error message.
         */
         try (FileOutputStream fos = new FileOutputStream(outputFilePath)) 
         {
             result_workbook.write(fos);
         }
         catch (IOException e)
         {
        	 System.out.println("An error occurred while writing the file. Please check the details: " + e);
         }
         
         /*
         Close the result workbook, input workbook, and file input stream.
         This ensures all resources are released properly.
         */
         
         result_workbook.close();
         workbook.close();
         fis.close();
         
         System.out.println();
         
         System.out.println("=== Project execution Ended at " + java.time.LocalDateTime.now() + " ===");
    }
}
