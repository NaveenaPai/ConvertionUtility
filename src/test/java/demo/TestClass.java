package demo;

import utility.ConvertDocToHTMUtility;

//import utility.SaveImagesToDocUtility;

public class TestClass {

	public static void main(String[] args) throws Exception {
		String fileName = "TestExecutionReport";
		String parentPath = "/src/main/resources/";
		String systemFolder = System.getProperty("user.dir");
		//String imageFolderPath = systemFolder + parentPath + "images/";
		String reportFolderPath = systemFolder + parentPath + "docReport/";

		// SaveImagesToDocUtility.SaveImagesToDoc(imageFolderPath, reportFolderPath, fileName);

		String htmlFolderPath = systemFolder + parentPath + "htmlReport/";
		ConvertDocToHTMUtility.DocxToHtml(reportFolderPath, fileName + ".docx", htmlFolderPath, fileName + ".htm");
	}

}
