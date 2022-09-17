package utility;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

public class SaveImagesToDocUtility {
	public static void SaveImagesToDoc(String sourcePath, String destnPath, String reportName) {

		List<String> filePaths = GetImages(sourcePath);

		XWPFDocument doc = new XWPFDocument();
		// Timestamp timestamp = new Timestamp(System.currentTimeMillis());
		// String time = timestamp.toString().replace(":", "_").replace(" ", "_");
		XWPFParagraph title = doc.createParagraph();
		XWPFRun run = title.createRun();

		try {
			for (String imgFile : filePaths) {

				run.setText(imgFile.substring(imgFile.lastIndexOf("/")+1, imgFile.length()));
				run.setBold(true);
				title.setAlignment(ParagraphAlignment.LEFT);
				File file = new File(imgFile);
				FileInputStream inputStream = new FileInputStream(file);
				run.addBreak();
				run.addPicture(inputStream, XWPFDocument.PICTURE_TYPE_PNG, imgFile, Units.toEMU(200), Units.toEMU(200)); // 200x200
				run.addBreak();
				inputStream.close();
			}
			destnPath = destnPath + reportName + ".docx";
			FileOutputStream fos = new FileOutputStream(destnPath);
			doc.write(fos);
			fos.close();

		} catch (Exception ex) {
			ex.printStackTrace();
		}
	}

	private static List<String> GetImages(String imageFilePath) {
		List<String> imageFiles = new ArrayList<String>();
		File folder = new File(imageFilePath);
		File[] listOfFiles = folder.listFiles();

		for (File file : listOfFiles) {
			if (file.isFile()) {
				imageFiles.add(imageFilePath + file.getName());
			}
		}
		return imageFiles;
	}
}
