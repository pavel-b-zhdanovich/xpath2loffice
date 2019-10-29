package iv.office.generic;

import com.sun.star.uno.Exception;


public interface IOfficeUnoRuntime {
	<T> T queryInterface(Class<T> arg0, Object arg1);
	<XComponentContext> XComponentContext getContext(boolean isHidden) throws Exception; 
}
