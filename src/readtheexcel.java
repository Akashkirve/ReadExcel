import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class readtheexcel {

	public static void main(String[] args) throws Exception {
	
String fileName = "excelfile\\excel1.xlsx"; // replace with your file name
        
        try (FileInputStream file = new FileInputStream(fileName);
             XSSFWorkbook workbook = new XSSFWorkbook(file)) {
            Sheet sheet1 = workbook.getSheetAt(0); // assuming sheet 1 is at index 0
            Sheet sheet2 = workbook.getSheetAt(1); // assuming sheet 2 is at index 1
            Sheet sheet3 = workbook.getSheetAt(2); // assuming sheet 3 is at index 2
            
            // read data from sheet 1 and store it in a map
            Map<String, String> sheet1Data = new HashMap<>();
            for (Row row : sheet1) {
                Cell keyCell = row.getCell(1); // assuming data is in column A
                Cell valueCell = row.getCell(2); // assuming data is in column B
                String key = keyCell.getStringCellValue();
                String value = valueCell.getStringCellValue();
                sheet1Data.put(key, value);
            }
            
            // compare data from sheet 1 with sheet 2 and sheet 3
            for (Row row : sheet2) {
                Cell keyCell = row.getCell(1); // assuming data is in column A
                String key = keyCell.getStringCellValue();
                if (sheet1Data.containsKey(key)) {
                    Cell valueCell = row.getCell(2); // assuming data is in column B
                    String value = valueCell.getStringCellValue();
                    String oldValue = sheet1Data.get(key);
                    if (value.equals(oldValue)) {
                        System.out.println("Data match for key " + key);
                        System.out.println("Sheet 2 value: " + value);
                        System.out.println("Sheet 1 value: " + oldValue);
                    }
                }
            }
            
            for (Row row : sheet3) {
                Cell keyCell = row.getCell(1); // assuming data is in column A
                String key = keyCell.getStringCellValue();
                if (sheet1Data.containsKey(key)) {
                    Cell valueCell = row.getCell(2); // assuming data is in column B
                    String value = valueCell.getStringCellValue();
                    String oldValue = sheet1Data.get(key);
                    if (value.equals(oldValue)) {
                        System.out.println("Data match for key " + key);
                        System.out.println("Sheet 3 value: " + value);
                        System.out.println("Sheet 1 value: " + oldValue);
                    }
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }	
		
	}
	
       }
