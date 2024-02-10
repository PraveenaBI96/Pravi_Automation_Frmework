package genericUtility;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelUtility {
	
	public static Workbook wb;
	
	public void openWorkBook() throws EncryptedDocumentException, IOException {
		FileInputStream fis = new FileInputStream(IConstants.EXCELPATH);
		wb= WorkbookFactory.create(fis);
	}
	
	public void closeWorkBook() throws IOException {
		wb.close();
	}

	public Object getDataFromExcel(String sheetName,String Test_Id,String data) throws EncryptedDocumentException, IOException {
		Sheet sheet = wb.getSheet(sheetName);
//		DataFormatter format = new DataFormatter();
		Object value = null;
		for (int i = 0; i <= getLastRowNumber(sheetName); i++) {
        try {
		if(sheet.getRow(i).getCell(0).getStringCellValue().equals(Test_Id)){
				for (int j = 0; j < getLastCellNumber(sheetName,i-1); j++) {
					if(sheet.getRow(i-1).getCell(j).getStringCellValue().equals(data)) {
						Cell cell = sheet.getRow(i).getCell(j);
						if(cell.getCellType()==CellType.STRING) {
							value=sheet.getRow(i).getCell(j).getStringCellValue();
							break;
						}else if(cell.getCellType()==CellType.NUMERIC) {
							double numericValue = sheet.getRow(i).getCell(j).getNumericCellValue();
							long intValue = (long) numericValue;
							value= intValue;
							break;
						}
					}
				}
			}
        }catch(NullPointerException e) {
        	continue;
        }
		}
		return value;
	}
	
	public Object getStringData(String sheetName, int rowNum, int cellNum) throws EncryptedDocumentException, IOException {
		Sheet sheet = wb.getSheet(sheetName);
		Object value = null;
		Cell cell = sheet.getRow(rowNum).getCell(cellNum);
		if(cell.getCellType()==CellType.STRING) {
			value=sheet.getRow(rowNum).getCell(cellNum).getStringCellValue();
		}else if(cell.getCellType()==CellType.NUMERIC) {
			double numericValue = sheet.getRow(rowNum).getCell(cellNum).getNumericCellValue();
			long intValue = (long) numericValue;
			value= intValue;
		}
		return  value;
	}
	
	public int getLastRowNumber(String sheetName) throws EncryptedDocumentException, IOException {
		return wb.getSheet(sheetName).getLastRowNum();
	}
	
	public int getLastCellNumber(String sheetName, int rowNumber) throws EncryptedDocumentException, IOException {
		return wb.getSheet(sheetName).getRow(rowNumber).getLastCellNum();
	}
}
