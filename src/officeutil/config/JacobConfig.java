package officeutil.config;

import com.jacob.com.JacobException;
import com.jacob.com.LibraryLoader;

import officeutil.consts.JacobMessage;

/**
 * 配置文件
 */
public class JacobConfig {
	
	/**
	 * 是否是ms office
	 */
	public static boolean isMsOffice = false;
	
	public static String JACOB_DLL_PATH = "";
	
	/**
	 * 配置jcob的dll路径
	 * @param jacobDllPath
	 */
	public static void configDll(String jacobDllPath, boolean isMsOffice) {
		JacobConfig.JACOB_DLL_PATH = jacobDllPath;
		JacobConfig.isMsOffice = isMsOffice;
		System.setProperty(LibraryLoader.JACOB_DLL_PATH, jacobDllPath);
	}
	
	/**
	 * 检查配置
	 */
	public static void check() {
		if("".equals(JACOB_DLL_PATH.trim())) {
			throw new JacobException(JacobMessage.NO_DLL);
		}
	}
	
}
