package utility;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;

import org.apache.poi.xwpf.converter.core.FileImageExtractor;
import org.apache.poi.xwpf.converter.core.FileURIResolver;
import org.apache.poi.xwpf.converter.xhtml.XHTMLConverter;
import org.apache.poi.xwpf.converter.xhtml.XHTMLOptions;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

public class ConvertDocToHTMUtility {

	public static void DocxToHtml(String reportFolderPath, String reportName, String htmlFolderPath, String htmlName)
			throws Exception {
		
		String sourcefile = reportFolderPath + reportName;
		File file = new File(sourcefile);

		// Loading word Document generation XWPFDocument Object
		InputStream inStream = new FileInputStream(file);
		XWPFDocument document = new XWPFDocument(inStream);

		File imageFolderFile = new File(reportFolderPath);
		XHTMLOptions options = XHTMLOptions.create().URIResolver(new FileURIResolver(imageFolderFile));
		// XHTMLOptions options = XHTMLOptions.create().setImageManager(new ImageManager(new File("./"), "images"));
		options.setExtractor(new FileImageExtractor(imageFolderFile));
		options.setIgnoreStylesIfUnused(false);
		options.setFragment(true);

		// Will XWPFDocument Convert to XHTML
		OutputStream out = new FileOutputStream(new File(htmlFolderPath + htmlName));
		XHTMLConverter.getInstance().convert(document, out, options);
	}

}
