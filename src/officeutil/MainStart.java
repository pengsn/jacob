package officeutil;

import officeutil.config.JacobConfig;
import officeutil.util.Office2PdfUtil;

public class MainStart {

	public static void main(String[] ars) {
		long time = System.currentTimeMillis();
		long endTime = 0;
		String s = System.getProperty("user.dir");
		endTime = System.currentTimeMillis();
		System.out.println(endTime - time);
		JacobConfig.configDll(s + "\\jacob-1.19-x64.dll", true);
		endTime = System.currentTimeMillis();
		System.out.println(endTime - time);
		Office2PdfUtil.office2pdf("D:\\nblh2020\\部署常用资料\\根分区在线扩容.doc", "D:\\dddd\\根分区在线扩容.pdf");
		Office2PdfUtil.office2pdf("D:\\nblh2020\\部署常用资料\\根分区在线扩容.doc", "D:\\dddd\\根分区在线扩容2.pdf");
		endTime = System.currentTimeMillis();
		System.out.println(endTime - time);
	}

}
