package officeutil.util;

import java.util.Date;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.JacobException;
import com.jacob.com.Variant;

import officeutil.config.JacobConfig;

public class Excel2PdfUtil {
	
	private static final int xlTypePDF = 0;
	
	/***
	 *   Excel转化成PDF  
	 * 
	 * @param inputFile
	 * @param pdfFile
	 * @return
	 */
	protected static int exec(String inputFile, String pdfFile) {
		try {
			ComThread.InitSTA(true);
			ActiveXComponent ax = null;
			if(JacobConfig.isMsOffice) {
				ax = new ActiveXComponent("Excel.Application");
			}else{
				ax = new ActiveXComponent("KET.Application");
			}
			System.out.println("开始转化Excel为PDF...");
			long date = new Date().getTime();
			ax.setProperty("Visible", false);
			ax.setProperty("AutomationSecurity", new Variant(3)); // 禁用宏
			Dispatch excels = ax.getProperty("Workbooks").toDispatch();

			Dispatch excel = Dispatch
					.invoke(excels, "Open", Dispatch.Method,
							new Object[] { inputFile, new Variant(false), new Variant(false) }, new int[9])
					.toDispatch();
			// 转换格式
			Dispatch.invoke(excel, "ExportAsFixedFormat", Dispatch.Method, new Object[] { new Variant(0), // PDF格式=0
					pdfFile, new Variant(xlTypePDF) // 0=标准 (生成的PDF图片不会变模糊)// 1=最小文件 // (生成的PDF图片糊的一塌糊涂)
			}, new int[1]);
			// 这里放弃使用SaveAs
			/*
			 * Dispatch.invoke(excel,"SaveAs",Dispatch.Method,new Object[]{
			 * outFile, new Variant(57), new Variant(false), new Variant(57),
			 * new Variant(57), new Variant(false), new Variant(true), new
			 * Variant(57), new Variant(true), new Variant(true), new
			 * Variant(true) },new int[1]);
			 */
			long date2 = new Date().getTime();
			int time = (int) ((date2 - date) / 1000);
			Dispatch.call(excel, "Close", new Variant(false));

			if (ax != null) {
				ax.invoke("Quit", new Variant[] {});
				ax = null;
			}
			return time;
		} catch (Exception e) {
			throw new JacobException(e.getCause().toString() + e.getMessage());
		}finally{
			ComThread.Release();
		}
	}

}
