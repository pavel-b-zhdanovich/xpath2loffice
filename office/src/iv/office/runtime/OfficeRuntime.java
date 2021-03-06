package iv.office.runtime;

import java.util.List;

import ooo.connector.BootstrapSocketConnector;
import ooo.connector.server.OOoServer;

import com.sun.star.comp.helper.BootstrapException;
import com.sun.star.uno.Exception;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;

import iv.office.generic.IOfficeUnoRuntime;
import iv.settings.OfficeSettings;

public class OfficeRuntime implements IOfficeUnoRuntime {

	@Override
	public <T> T queryInterface(Class<T> zInterface,
            Object object){
		return UnoRuntime.queryInterface(zInterface,object);
	}

	@SuppressWarnings("unchecked")
	@Override
	public XComponentContext getContext(boolean isHidden) throws Exception {
		OfficeSettings settings = new OfficeSettings();
		String ooofolder= settings.OFFICE_EXEC_PATH;
		System.out.println("OFFICE_EXEC_PATH="+settings.OFFICE_EXEC_PATH);
		System.out.println("ooofolder="+ooofolder);
        List<String> options = OOoServer.getDefaultOOoOptions();
        options.add("--nofirststartwizard");
        if (isHidden) {options.add("--headless");}
        OOoServer oooServer = new OOoServer(ooofolder, options);

        BootstrapSocketConnector bootstrapSocketConnector = new BootstrapSocketConnector(oooServer);
        XComponentContext xRemoteContext;
		try {
			xRemoteContext = bootstrapSocketConnector.connect();
		} catch (BootstrapException e){ 
			e.printStackTrace();
			throw new Exception(); 
		}
		if (xRemoteContext == null) {
	         System.err.println("ERROR: Could not bootstrap default Office.");}
	    return xRemoteContext;
	}

}
