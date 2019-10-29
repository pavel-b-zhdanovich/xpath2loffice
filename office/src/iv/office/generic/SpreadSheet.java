package iv.office.generic;

import com.sun.star.beans.PropertyVetoException;
import com.sun.star.beans.UnknownPropertyException;
import com.sun.star.beans.XPropertySet;
import com.sun.star.container.NoSuchElementException;
import com.sun.star.lang.IllegalArgumentException;
import com.sun.star.lang.IndexOutOfBoundsException;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.sheet.XNamedRanges;
import com.sun.star.sheet.XSpreadsheet;
import com.sun.star.table.XCell;
import com.sun.star.table.XCellRange;



public class SpreadSheet {
	
	
	private XSpreadsheet sheet;
	protected IOfficeUnoRuntime rtime;
	protected XNamedRanges xNamedRanges;
	
	public SpreadSheet(XSpreadsheet xsheet, IOfficeUnoRuntime prtime) throws NoSuchElementException, WrappedTargetException {
		sheet= xsheet;
		rtime=prtime;
		
	}
	public String getName(){
		return "todo";
	}
	
	
	public void printtoCell(String cellname, String text) throws Exception{
		XCellRange range;
		try {
			range = sheet.getCellRangeByName(cellname);
		} catch (Exception e) {
			throw new IndexOutOfBoundsException();
		}
		XCell cell = range.getCellByPosition(0, 0);
		cell. setFormula(text);
	}	
	
	public XSpreadsheet getxSpreadSheet(){
		return sheet;
	}
	
	public String [] getRanges() throws UnknownPropertyException, WrappedTargetException{
		XPropertySet sheetProps = ( XPropertySet ) rtime.queryInterface(XPropertySet.class, sheet );
		Object obj = sheetProps.getPropertyValue("NamedRanges");
		xNamedRanges =(XNamedRanges)rtime.queryInterface(XNamedRanges.class, obj);
		return xNamedRanges.getElementNames();
	}
	
	public void setIsVisible(Boolean b) throws UnknownPropertyException, WrappedTargetException, PropertyVetoException, IllegalArgumentException {
		 XPropertySet sheetProps = ( XPropertySet ) rtime.queryInterface(XPropertySet.class, sheet );
		sheetProps.setPropertyValue("IsVisible",b);
	}
	
}
