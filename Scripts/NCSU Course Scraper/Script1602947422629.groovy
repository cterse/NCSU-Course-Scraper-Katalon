import static com.kms.katalon.core.checkpoint.CheckpointFactory.findCheckpoint
import static com.kms.katalon.core.testcase.TestCaseFactory.findTestCase
import static com.kms.katalon.core.testdata.TestDataFactory.findTestData
import static com.kms.katalon.core.testobject.ObjectRepository.findTestObject
import static com.kms.katalon.core.testobject.ObjectRepository.findWindowsObject

import java.util.Map.Entry

import com.kms.katalon.core.checkpoint.Checkpoint as Checkpoint
import com.kms.katalon.core.cucumber.keyword.CucumberBuiltinKeywords as CucumberKW
import com.kms.katalon.core.mobile.keyword.MobileBuiltInKeywords as Mobile
import com.kms.katalon.core.model.FailureHandling as FailureHandling
import com.kms.katalon.core.testcase.TestCase as TestCase
import com.kms.katalon.core.testdata.TestData as TestData
import com.kms.katalon.core.testng.keyword.TestNGBuiltinKeywords as TestNGKW
import com.kms.katalon.core.testobject.TestObject as TestObject
import com.kms.katalon.core.webservice.keyword.WSBuiltInKeywords as WS
import com.kms.katalon.core.webui.driver.DriverFactory
import com.kms.katalon.core.webui.keyword.WebUiBuiltInKeywords as WebUI
import com.kms.katalon.core.windows.keyword.WindowsBuiltinKeywords as Windows
import internal.GlobalVariable as GlobalVariable

import org.apache.poi.xssf.usermodel.XSSFRow
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.openqa.selenium.Keys as Keys
import org.openqa.selenium.WebDriver
import org.openqa.selenium.WebElement
import org.openqa.selenium.By

XSSFWorkbook workbook = new XSSFWorkbook();
XSSFSheet sheet = workbook.createSheet();

WebUI.openBrowser('')

WebUI.navigateToUrl('https://webappprd.acs.ncsu.edu/php/coursecat/directory.php#course-search-results')

WebUI.selectOptionByValue(findTestObject('Page_Course Catalog NCSU/select_Undergraduate                       _c39826'), 
    'GRAD', true)

WebUI.click(findTestObject('Page_Course Catalog NCSU/button_Browse'))

WebUI.click(findTestObject('Page_Course Catalog NCSU/a_CSC - Computer Science'))

WebUI.click(findTestObject('Page_Course Catalog NCSU/button_Search'))

WebDriver webDriver = DriverFactory.getWebDriver();
// List<WebElement> courses = webDriver.findElements(By.xpath('//*[@id="course-search-results"]/ul/li/a'))
List<WebElement> courses = webDriver.findElements(By.cssSelector('#course-search-results > ul > li> a'))

for (int i=0; i<courses.size(); i++) {
	try {
		
	 
	TestObject courseObject = WebUI.convertWebElementToTestObject(courses[i])
	WebUI.scrollToElement(courseObject, 30)
	WebUI.click(courseObject)
	
	String courseCodeName = WebUI.getText(findTestObject('Page_Course Catalog NCSU/h_modalTitle'))
	String courseCode = courseCodeName
	String courseName = ""
	if (courseCodeName.contains(':')) {
		courseCode = courseCodeName.split(':')[0].trim();
		courseName = courseCodeName.split(':')[1].trim();
	}
	String courseUnits = WebUI.getText(findTestObject('Page_Course Catalog NCSU/span_course_units'))
	String courseDesc = WebUI.getText(findTestObject('Page_Course Catalog NCSU/span_course_descr'))
	String coursePrereq = WebUI.getText(findTestObject('Page_Course Catalog NCSU/p_Prerequisite'))
	if (coursePrereq != null && coursePrereq.startsWith("Prerequisite:")) {
		coursePrereq = coursePrereq.split(":", 2)[1].trim()
	}
	
	String courseAttributes = ""
	List<WebElement> attributes = webDriver.findElements(By.cssSelector('#course-attr > p'))
	for (int j = 0; j<attributes.size(); j++) {
		courseAttributes = attributes[j].getText() + " "
	}
	courseAttributes = courseAttributes.trim()
	
	Map<String, Set<String>> semProfMap = new HashMap<>();
	List<WebElement> semLinks = webDriver.findElements(By.cssSelector('#course-sem > a.sem-link'))
	for (int j=0; j<semLinks.size(); j++) {
		String semName = semLinks[j].getText()
		
		TestObject semLinkObj = WebUI.convertWebElementToTestObject(semLinks[j])
		WebUI.scrollToElement(semLinkObj, 30)
		WebUI.click(semLinkObj)
		
		List<WebElement> profLinks = webDriver.findElements(By.cssSelector('#search-results > table > tbody > tr > td:nth-child(7) > a'))
		Set<String> tempSet = new HashSet<>();
		for (WebElement profLink : profLinks) {
			tempSet.add(profLink.getText())
		}
		semProfMap.put(semName, tempSet)
	}
	
	println(courseCode+"~~"+courseName+"~~"+courseUnits+"~~"+courseDesc+"~~"+coursePrereq+"~~"+courseAttributes+"~~"+semProfMap)
	XSSFRow row = sheet.createRow(i)
	row.createCell(0).setCellValue(courseCode)
	row.createCell(1).setCellValue(courseName)
	row.createCell(2).setCellValue(courseUnits)
	row.createCell(3).setCellValue(courseDesc)
	row.createCell(4).setCellValue(coursePrereq)
	row.createCell(5).setCellValue(courseAttributes)
	String semProfString = ""
	for (Entry<String, Set<String>> e : semProfMap.entrySet()) {
		semProfString += e.getKey()+":"+e.getValue()+"  "
	}
	row.createCell(6).setCellValue(semProfString.trim())
	
	if (i % 10 == 0) {
		FileOutputStream outFile =new FileOutputStream(new File("E:\\Testdata.xlsx"));
		workbook.write(outFile);
		outFile.close();
	}
	
	WebUI.click(findTestObject('Page_Course Catalog NCSU/btn_Close_modal'))
	} catch (Exception e) {
	e.printStackTrace()
	}
}

FileOutputStream outFile =new FileOutputStream(new File("E:\\Testdata.xlsx"));
workbook.write(outFile);
outFile.close();