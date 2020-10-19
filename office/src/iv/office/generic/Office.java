package iv.office.generic;

import com.sun.star.beans.Property;
import com.sun.star.beans.PropertyValue;
import com.sun.star.beans.UnknownPropertyException;
import com.sun.star.beans.XPropertySet;
import com.sun.star.beans.XPropertySetInfo;
import com.sun.star.frame.XComponentLoader;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.sheet.XSpreadsheetDocument;
import com.sun.star.text.XTextDocument;
import com.sun.star.uno.XComponentContext;

public class Office {
	
	public static int HIDDEN = 1;
	public static int VISIBLE = 0;
	private Object desktop;
	protected IOfficeUnoRuntime rtime;
	protected boolean isHidden = false;
	
	public Office(IOfficeUnoRuntime prtime) {
		rtime=prtime;
	}
	
	public Office(IOfficeUnoRuntime prtime, int flags) {
		isHidden = (flags & HIDDEN) == 1; 
		rtime=prtime;
	}
		 
	protected  XComponentContext getContext()  throws Exception{
		return rtime.getContext(isHidden);
	}
	
	private XMultiComponentFactory getServiceManager()  throws Exception{
		return getContext().getServiceManager();
	}
		
	private Object getDesktop() throws Exception{
		if (desktop == null){
			desktop = getServiceManager().createInstanceWithContext(
			         "com.sun.star.frame.Desktop", getContext());}
		return desktop;
		}
		
	private XComponentLoader getLoader() throws Exception{
		return (XComponentLoader)
        rtime.queryInterface(XComponentLoader.class, getDesktop());
	}
	
	protected XSpreadsheetDocument _openCalcDocument(String url, Boolean hidden) throws Exception{
		PropertyValue[] loadProps = new PropertyValue[2];
		loadProps[0] = new com.sun.star.beans.PropertyValue();
		loadProps[0]. Name="FilterName";
		loadProps[0].Value="*.*";
		loadProps[1] = new com.sun.star.beans.PropertyValue();
		loadProps[1].Name = "Hidden";
		loadProps[1].Value = hidden; 
		
		XComponent component = getLoader().loadComponentFromURL(url, 
				                                                            "_blank", 
				                                                            0, 
				                                                            loadProps
				                                                            );
		return (XSpreadsheetDocument)rtime.queryInterface(	XSpreadsheetDocument.class,
                                  							component
                                  							);
	}
	
	protected XTextDocument _openTxtDocument(String url, Boolean hidden) throws Exception{
		PropertyValue[] loadProps = new PropertyValue[1];
		loadProps[0] = new com.sun.star.beans.PropertyValue();
		loadProps[0].Name = "Hidden";
		loadProps[0].Value = hidden; 
		
		XComponent component = getLoader().loadComponentFromURL(url, 
				                                                            "_blank", 
				                                                            0, 
				                                                            loadProps
				                                                            );
		return (XTextDocument)rtime.queryInterface(	XTextDocument.class,
													component
													);
	}
	
	public CalcDocument openCalcDocument(String url, Boolean hidden) throws Exception{
		return new CalcDocument(_openCalcDocument(url, hidden),rtime);
	}
	
	public TextDocument openTextDocument(String url, Boolean hidden) throws Exception{
		return new TextDocument(_openTxtDocument(url, hidden),rtime);
	}

	public static void printProperties(Object obj, IOfficeUnoRuntime rtime) throws WrappedTargetException{
		XPropertySet propSet = ( XPropertySet ) rtime.queryInterface(XPropertySet.class, obj );
		XPropertySetInfo propinfo = propSet.getPropertySetInfo();
		Property[] props = propinfo.getProperties();
		for (Property property : props) {
			try {
				System.out.println(property.Name+"="+propSet.getPropertyValue(property.Name));
			} catch (UnknownPropertyException e) {
				System.err.println(property.Name+"is UNKNOWN!");
			}
		}
	}


}
