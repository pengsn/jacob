package officeutil.util;

/**
 * 获取文件后缀
 */
public class OfficeFileUtil {

	/***
	 * 判断文件类型
	 * 
	 * @param fileName
	 * @return
	 */
	public static String getFileSuffix(String fileName) {
		int splitIndex = fileName.lastIndexOf(".");
		return fileName.substring(splitIndex + 1);
	}
	
}
