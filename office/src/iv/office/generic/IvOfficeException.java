package iv.office.generic;

public class IvOfficeException extends Exception {

	/**
	 * 
	 */
	private static final long serialVersionUID = 486974151673204385L;

	public IvOfficeException(String message) {
		super(message);
	}
	public static void reportErr(String pth) {
			System.err.println("Выражение \""+pth+"\" не может быть вычислено.");
	}
}