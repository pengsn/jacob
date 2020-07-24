package officeutil.util;

import java.util.Date;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.JacobException;
import com.jacob.com.Variant;

import officeutil.config.JacobConfig;

public class PPT2PdfUtil {

	private static final int ppSaveAsPDF = 32;
	
	/***
	 * ppt转化成PDF  
	 * 
	 * @param inputFile
	 * @param pdfFile
	 * @return
	 */
	protected static int exec(String inputFile, String pdfFile) {
		System.out.println("开始转化PPT为PDF...");
		try {
			ComThread.InitSTA(true);
			ActiveXComponent app = null;
			if(JacobConfig.isMsOffice) {
				app = new ActiveXComponent("PowerPoint.Application");
			}else{
				app = new ActiveXComponent("KWPP.Application");
			}
			//            app.setProperty("Visible", false);
			long date = new Date().getTime();
			Dispatch ppts = app.getProperty("Presentations").toDispatch();
			Dispatch ppt = Dispatch.call(ppts, "Open", inputFile, true, // ReadOnly
					// false, // Untitled指定文件是否有标题
					false// WithWindow指定文件是否可见
			).toDispatch();
			Dispatch.invoke(ppt, "SaveAs", Dispatch.Method, new Object[] { pdfFile, new Variant(ppSaveAsPDF) },
					new int[1]);
			System.out.println("PPT");
			Dispatch.call(ppt, "Close");
			long date2 = new Date().getTime();
			int time = (int) ((date2 - date) / 1000);
			app.invoke("Quit");
			return time;
		} catch (Exception e) {
			throw new JacobException(e.getCause().toString() + e.getMessage());
		}finally{
			ComThread.Release();
		}
	}

	// 删除多余的页，并转换为PDF
	protected static void interceptPPT(String inputFile, String pdfFile) {
		ActiveXComponent app = null;
		try {
			ComThread.InitSTA(true);
			app = new ActiveXComponent("KWPP.Application");
			ActiveXComponent presentations = app.getPropertyAsComponent("Presentations");
			ActiveXComponent presentation = presentations.invokeGetComponent("Open", new Variant(inputFile),
					new Variant(false));
			int count = Dispatch.get(presentations, "Count").getInt();
			System.out.println("打开文档数:" + count);
			ActiveXComponent slides = presentation.getPropertyAsComponent("Slides");
			int slidePages = Dispatch.get(slides, "Count").getInt();
			System.out.println("ppt幻灯片总页数:" + slidePages);

			// 总页数的20%取整+1 最多不超过5页
			int target = (int) (slidePages * 0.5) + 1 > 5 ? 5 : (int) (slidePages * 0.5) + 1;
			// 删除指定页数
			while (slidePages > target) {
				// 选中指定页幻灯片
				Dispatch slide = Dispatch.call(presentation, "Slides", slidePages).toDispatch();
				Dispatch.call(slide, "Select");
				Dispatch.call(slide, "Delete");
				slidePages--;
				System.out.println("当前ppt总页数:" + slidePages);
			}
			Dispatch.invoke(presentation, "SaveAs", Dispatch.Method, new Object[] { pdfFile, new Variant(32) },
					new int[1]);
			Dispatch.call(presentation, "Save");
			Dispatch.call(presentation, "Close");
			presentation = null;
			app.invoke("Quit");
			app = null;
			ComThread.Release();
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
		}
	}

}
