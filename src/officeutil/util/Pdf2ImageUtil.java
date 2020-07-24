package officeutil.util;

import java.awt.image.BufferedImage;
import java.io.File;
import java.io.IOException;

import javax.imageio.ImageIO;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.rendering.PDFRenderer;

import com.itextpdf.text.pdf.PdfReader;

/**
 * pdf转图片
 */
public class Pdf2ImageUtil {

	public static void main(String[] args) {
		String outPutPath = "D:/dddd/根分区在线扩容.pdf";
		String dstImgFolder = "D:/dddd/";
		long start = System.currentTimeMillis();
		pdf2ImageDemo(outPutPath, dstImgFolder, 300);
		long end = System.currentTimeMillis();
		System.out.println("总共耗时：" + (end - start));
	}

	/***
	 * PDF文件转PNG/JPEG图片
	 * 
	 * @param PdfFilePathpdf完整路径
	 * @param imgFilePath图片存放的文件夹
	 * @param dpi越大转换后越清晰，相对转换速度越慢,一般电脑默认96dpi
	 */
	public static void pdf2ImageDemo(String PdfFilePath, String dstImgFolder, int dpi) {
		File file = new File(PdfFilePath);
		PDDocument pdDocument;
		try {
			String imgPDFPath = file.getParent();
			int dot = file.getName().lastIndexOf('.');
			// 获取图片文件名
			String imagePDFName = file.getName().substring(0, dot);
			String imgFolderPath = null;
			if (dstImgFolder.equals("")) {
				// 获取图片存放的文件夹路径
				imgFolderPath = imgPDFPath + File.separator + imagePDFName;
			} else {
				imgFolderPath = dstImgFolder + File.separator + imagePDFName;
			}

			if (createDirectory(imgFolderPath)) {
				pdDocument = PDDocument.load(file);
				PDFRenderer renderer = new PDFRenderer(pdDocument);
				PdfReader reader = new PdfReader(PdfFilePath);
				int pages = reader.getNumberOfPages();// 获取PDF页数
				System.out.println("PDF page number is:" + pages);
				StringBuffer imgFilePath = null;
				for (int i = 0; i < pages; i++) {
					String imgFilePathPrefix = imgFolderPath + File.separator + imagePDFName;
					imgFilePath = new StringBuffer();
					imgFilePath.append(imgFilePathPrefix);
					imgFilePath.append("_");
					imgFilePath.append(String.valueOf(i + 1));
					imgFilePath.append(".png");// PNG
					File dstFile = new File(imgFilePath.toString());
					BufferedImage image = renderer.renderImageWithDPI(i, dpi);
					ImageIO.write(image, "png", dstFile);// PNG
				}
				System.out.println("PDF文档转PNG图片成功！");
			} else {
				System.out.println("PDF文档转PNG图片失败：" + "创建" + imgFolderPath + "失败");
			}
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	private static boolean createDirectory(String folder) {
		File dir = new File(folder);
		if (dir.exists()) {
			return true;
		} else {
			return dir.mkdirs();
		}
	}
}
