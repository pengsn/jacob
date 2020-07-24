package officeutil.util;

import java.io.File;

import com.jacob.com.JacobException;

import officeutil.config.JacobConfig;
import officeutil.consts.JacobMessage;
import officeutil.consts.OfficeFileSuffix;

/**
 * jacob office转pdf
 */
public class Office2PdfUtil {
	
	/**
	 * office文件转pdf
	 * 
	 * @param inputFilepath
	 * @param outputFilePath
	 * @return
	 */
	public static int office2pdf(String inputFilepath, String outputFilePath) {
		JacobConfig.check();
		return convert2PDF(inputFilepath, outputFilePath);
	}

	/***
	 * 判断需要转化文件的类型（Excel、Word、ppt）  
	 * 
	 * @param inputFile
	 * @param pdfFile
	 */
	private static int convert2PDF(String inputFile, String pdfFile) {
		String kind = OfficeFileUtil.getFileSuffix(inputFile);
		File file = new File(inputFile);
		if (!file.exists()) {
			throw new JacobException(JacobMessage.FILE_NOT_EXISTS);
		}
		File outfile = new File(pdfFile);
		if(!outfile.getParentFile().exists()) {
			outfile.getParentFile().mkdirs();
		}
		if (OfficeFileSuffix.pdf.name().equals(kind)) {
			throw new JacobException(JacobMessage.ALREADY_PDF);
		}
		if (OfficeFileSuffix.doc.name().equals(kind) || OfficeFileSuffix.docx.name().equals(kind) || OfficeFileSuffix.txt.name().equals(kind)) {
			return Word2PdfUtil.exec(inputFile, pdfFile);
		}
		if (OfficeFileSuffix.xls.name().equals(kind) || OfficeFileSuffix.xlsx.name().equals(kind)) {
			return Excel2PdfUtil.exec(inputFile, pdfFile);
		}
		if (OfficeFileSuffix.ppt.name().equals(kind) || OfficeFileSuffix.pptx.name().equals(kind)) {
			return PPT2PdfUtil.exec(inputFile, pdfFile);
		}
		throw new JacobException(JacobMessage.TYPE_NOT_SUPPORT);
	}

}
