package uiAutomation;

import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.testng.annotations.Test;

import genericUtility.BaseClass;
import genericUtility.DataProviderClass;


public class FetchDataFromExcel extends BaseClass {
	@Test
	public void Test_Id_01() throws EncryptedDocumentException, IOException {
		System.out.println(eLib.getDataFromExcel("Sheet1", "Test_Id_02", "name"));
		System.out.println(eLib.getDataFromExcel("Sheet1", "Test_Id_02", "id"));
		System.out.println(eLib.getDataFromExcel("Sheet1", "Test_Id_02", "phone"));
		System.out.println("=======================================");
		System.out.println(eLib.getDataFromExcel("Sheet1", "Test_Id_01", "name"));
		System.out.println(eLib.getDataFromExcel("Sheet1", "Test_Id_01", "id"));
		System.out.println(eLib.getDataFromExcel("Sheet1", "Test_Id_01", "phone"));
	}
	@Test(dataProvider ="Test_Id_03",dataProviderClass=DataProviderClass.class)
	public void Test_Id_03(String name, long id, long phone) {
		System.out.println(name+" "+id+" "+phone);
	}
}
