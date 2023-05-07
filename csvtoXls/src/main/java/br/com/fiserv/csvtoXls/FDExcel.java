package br.com.fiserv.csvtoXls;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.springframework.stereotype.Service;

@Service
public class FDExcel {
	

public static void fgen(String p_input_csv) {
    	
    	try {
    		
    		BufferedReader csvReader = new BufferedReader(new FileReader(p_input_csv+".csv"));
        	String rowCsv;
        	
        	FileInputStream inputStream = new FileInputStream(new File(p_input_csv));
        	Workbook workbook = WorkbookFactory.create(inputStream);
    		Sheet sheet = workbook.getSheetAt(0);
    		int rowCount = 5;
    		
			while ((rowCsv = csvReader.readLine()) != null) {
				
				String[] data = rowCsv.split(";");
			    
			    Row row = sheet.createRow(rowCount);
			    
			    int columnCount = -1;
			    Cell cell;
			    
				for (Object field : data) {
					cell = row.createCell(++columnCount);
					if (field instanceof String) {
						cell.setCellValue((String) field);
					} else if (field instanceof Integer) {
						cell.setCellValue((int) field);
					}
				}
				
				rowCount++;
			}
			
			inputStream.close();
			FileOutputStream outputStream = new FileOutputStream(p_input_csv);
			workbook.write(outputStream);
			workbook.close();
			outputStream.close();
			csvReader.close();
			
    	} catch (IOException | EncryptedDocumentException e) {
			e.printStackTrace();
		}
    	
	}
}
