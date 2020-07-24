package officeutil.util;

import java.util.Date;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.JacobException;
import com.jacob.com.Variant;

import officeutil.config.JacobConfig;

/**
 * word转pdf
 */
public class Word2PdfUtil {

	/**
	 *  word保存为pdf格式宏，值为17
	 */
	private static final int wdFormatPDF = 17;
	
	/***
	 *   Word转PDF  
	 * 
	 * @param inputFile
	 * @param pdfFile
	 * @return
	 */

	protected static int exec(String inputFile, String pdfFile) {
		long date = new Date().getTime();
		try {
			ComThread.InitSTA(true);
			// 打开Word应用程序
			ActiveXComponent app =  null;
			if(JacobConfig.isMsOffice) {
				app = new ActiveXComponent("Word.Application");
			}else{
				app= new ActiveXComponent("KWPS.Application");
			}
			// 设置Word不可见
			app.setProperty("Visible", new Variant(false));
			// 禁用宏
			app.setProperty("AutomationSecurity", new Variant(3));
			// 获得Word中所有打开的文档，返回documents对象
			Dispatch docs = app.getProperty("Documents").toDispatch();
			// 调用Documents对象中Open方法打开文档，并返回打开的文档对象Document
			Dispatch doc = Dispatch.call(docs, "Open", inputFile, false, true).toDispatch();
			/***
			 *   调用Document对象的SaveAs方法，将文档保存为pdf格式   Dispatch.call(doc,
			 * "SaveAs", pdfFile, wdFormatPDF word保存为pdf格式宏，值为17 )  
			 */
			Dispatch.call(doc, "ExportAsFixedFormat", pdfFile, wdFormatPDF);
			System.out.println(doc);
			// 关闭文档
			long date2 = new Date().getTime();
			int time = (int) ((date2 - date) / 1000);
			Dispatch.call(doc, "Close", false);
			// 关闭Word应用程序
			app.invoke("Quit", 0);
			return time;
		} catch (Exception e) {
			e.printStackTrace();
			throw new JacobException(e.getCause().toString() + e.getMessage());
		}finally{
			ComThread.Release();
		}
	}

}
