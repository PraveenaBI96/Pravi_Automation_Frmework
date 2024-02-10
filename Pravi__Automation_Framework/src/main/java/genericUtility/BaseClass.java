package genericUtility;

import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.testng.annotations.BeforeSuite;

public class BaseClass {
	
   public ExcelUtility eLib = new ExcelUtility();
	@BeforeSuite
	public void startSuite() throws EncryptedDocumentException, IOException {
		eLib.openWorkBook();
	}
	
	public void endSuite() throws IOException {
		eLib.closeWorkBook();
	}

}
