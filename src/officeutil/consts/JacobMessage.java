package officeutil.consts;

/**
 * jacob转化异常消息
 */
public interface JacobMessage {

	public String NO_DLL = "未配置dll目录";
	
	public String UNKNOWN = "转化失败，未知错误...";
	
	public String TYPE_NOT_SUPPORT = "文件类型不支持转换";
	
	public String ALREADY_PDF = "原文件就是PDF文件,无需转化...";
	
	public String FILE_NOT_EXISTS = "转化失败，文件不存在...";
	
	public String TRY_AGAIN = "转化失败，请重新尝试...";

}
