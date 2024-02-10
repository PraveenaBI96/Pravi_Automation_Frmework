package genericUtility;

import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.testng.annotations.DataProvider;

public class DataProviderClass extends BaseClass{

	public Object[][] getData(int set, int data,String sheetName,int rowNum, int cellNum) throws EncryptedDocumentException, IOException{
		Object [][] objarr = new Object[set][data];
		int row=rowNum;
		for(int i=0; i<set; i++) {
			int cell =cellNum;
			for (int j = 0; j < data; j++) {
				objarr[i][j]= eLib.getStringData(sheetName, row, cell);
				cell++;
			}
			row++;
		}
		return objarr;
	}
	
	@DataProvider(name="Test_Id_03")
	public Object[][] getDataTest_Id_03() throws EncryptedDocumentException, IOException {
		return getData(3,3,"Sheet1",7,1);
	}
}
