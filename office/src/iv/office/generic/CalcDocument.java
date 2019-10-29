package iv.office.generic;

import javax.xml.parsers.ParserConfigurationException;

import org.xml.sax.SAXException;

import com.sun.star.beans.PropertyValue;
import com.sun.star.beans.UnknownPropertyException;
import com.sun.star.beans.XPropertySet;
import com.sun.star.container.NoSuchElementException;
import com.sun.star.container.XIndexAccess;
import com.sun.star.container.XNameAccess;
import com.sun.star.frame.XStorable;
import com.sun.star.io.IOException;
import com.sun.star.lang.IllegalArgumentException;
import com.sun.star.lang.IndexOutOfBoundsException;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.sheet.XCalculatable;
import com.sun.star.sheet.XCellRangeReferrer;
import com.sun.star.sheet.XNamedRanges;
import com.sun.star.sheet.XSpreadsheet;
import com.sun.star.sheet.XSpreadsheetDocument;
import com.sun.star.sheet.XSpreadsheets;
import com.sun.star.table.XTableCharts;
import com.sun.star.table.XTableChartsSupplier;
import com.sun.star.view.XPrintable;
import iv.utils.xml.IXMLSource;

public class CalcDocument {

	public XSpreadsheetDocument document;
	protected IOfficeUnoRuntime rtime;
	/**
	 * @return the rtime
	 */
	public IOfficeUnoRuntime getRtime() {
		return rtime;
	}
	protected String[] rangeNames;
	protected XNamedRanges xNamedRanges;

	public CalcDocument(XSpreadsheetDocument pdocument, IOfficeUnoRuntime prtime) {
		document = pdocument;
		rtime = prtime;
	}

	public SpreadSheet getSheet(String sheetName) throws NoSuchElementException, WrappedTargetException {
		XSpreadsheets xSpreadsheets = document.getSheets();
		Object sheet = xSpreadsheets.getByName(sheetName);
		XSpreadsheet xsheet = (XSpreadsheet) rtime.queryInterface(
				XSpreadsheet.class, sheet);
		return new SpreadSheet(xsheet,rtime);
	}
	
	public SpreadSheet getSheet(short num) 
			throws NoSuchElementException, WrappedTargetException, IndexOutOfBoundsException {
		XSpreadsheets xSpreadsheets = document.getSheets();
		XIndexAccess ia = rtime.queryInterface(XIndexAccess.class, xSpreadsheets);
		return new SpreadSheet((XSpreadsheet) ia.getByIndex(num),rtime);
	}

	public SheetCellRange getRange(String sheetName, String rangeName) 
			throws NoSuchElementException, WrappedTargetException
			{
		XSpreadsheets xSpreadsheets = document.getSheets();
		Object sheet = xSpreadsheets.getByName(sheetName);
		XSpreadsheet xsheet = (XSpreadsheet) rtime.queryInterface(
				XSpreadsheet.class, sheet);
		return new SheetCellRange(xsheet, rangeName, rtime);
	}

	public SheetCellRange getRange(String rangeName) 
			throws UnknownPropertyException, WrappedTargetException, NoSuchElementException
			{
		return new SheetCellRange(getReferrer(rangeName).getReferredCells(), rtime);
	}


	
	public void save_xls(String storeUrl) throws IvOfficeException {
		XStorable xStorable = (XStorable) rtime.queryInterface(XStorable.class,
				document);
		PropertyValue[] storeProps = new PropertyValue[1];
		storeProps[0] = new PropertyValue();
		storeProps[0].Name = "FilterName";
		storeProps[0].Value = "MS Excel 97";
		try {
			xStorable.storeAsURL(storeUrl, storeProps);
		} catch (IOException e) {
			e.printStackTrace();
			throw new IvOfficeException("Не удалось сохранить документ в виде xls");
		}
	}

	
	public void save_pdf(String storeUrl) throws IvOfficeException {
		XStorable xStorable = (XStorable) rtime.queryInterface(XStorable.class,
				document);
		PropertyValue[] storeProps = new PropertyValue[2];
		storeProps[0] = new PropertyValue();
		storeProps[0].Name = "Overwrite";
		storeProps[0].Value = new Boolean(true);
		storeProps[1] = new PropertyValue();
		storeProps[1].Name = "FilterName";
		storeProps[1].Value = "calc_pdf_Export";
		try {
			xStorable.storeToURL(storeUrl, storeProps);
		} catch (IOException e) {
			e.printStackTrace();
			throw new IvOfficeException("Не удалось сохранить документ в виде pdf");
		}
	}

	public void save() throws IOException  {
		XStorable xStorable = (XStorable) rtime.queryInterface(XStorable.class,
				document);
		xStorable.store();
	}
	
	public void save(String storeUrl) throws IOException  {
		XStorable xStorable = (XStorable) rtime.queryInterface(XStorable.class,
				document);
		PropertyValue[] storeProps = new PropertyValue[0];
		xStorable.storeAsURL(storeUrl, storeProps);
	}

	public void printDocument(String printername) throws IllegalArgumentException  {

		XPrintable xPrintable = (XPrintable) rtime.queryInterface(
				XPrintable.class, document);

		PropertyValue[] printOpts = new PropertyValue[1];
		printOpts[0] = new com.sun.star.beans.PropertyValue();
		printOpts[0].Name = "Name";
		printOpts[0].Value = printername;

		xPrintable.setPrinter(printOpts);

		xPrintable.print(printOpts);
	}

	public void xml2Cell(IXMLSource xmlSource, String pth, String sheetName,
			String cellName) throws ParserConfigurationException,
			SAXException, java.io.IOException, Exception {
		new XML2Ods(xmlSource, this).xPathExpr2Cell(pth, sheetName, cellName);
	}
	
	
	protected void getRanges() throws UnknownPropertyException, WrappedTargetException{
		XPropertySet docProps = ( XPropertySet ) rtime.queryInterface(XPropertySet.class, document );
		Object obj = docProps.getPropertyValue("NamedRanges");
		xNamedRanges =(XNamedRanges)rtime.queryInterface(XNamedRanges.class, obj);
		rangeNames = xNamedRanges.getElementNames();
	}
	
	public String[] getRangeNames() throws UnknownPropertyException, WrappedTargetException{
		if (rangeNames == null) getRanges();
		return rangeNames;
	}
	
	protected XCellRangeReferrer getReferrer(String rName) throws UnknownPropertyException, WrappedTargetException, NoSuchElementException{
		if (xNamedRanges == null) getRanges();
		Object r = xNamedRanges.getByName(rName);
		XCellRangeReferrer rref = rtime.queryInterface(XCellRangeReferrer.class, r);
		return rref;
	}
	
	public void copySheet(String source, String dest, short ind){
		XSpreadsheets xSpreadsheets = document.getSheets();
		xSpreadsheets.copyByName(source, dest, ind);
	}

	public void copySheetToTheEnd(String source, String dest){
		XSpreadsheets xSpreadsheets = document.getSheets();
		short ind=(short)xSpreadsheets.getElementNames().length;
		xSpreadsheets.copyByName(source, dest, ind);
	}


	public void recalculate(){
		XCalculatable x = rtime.queryInterface(XCalculatable.class, document);
		x.calculateAll();
	}
	
	
	public String[] getChartNames(String sheetName) throws NoSuchElementException, WrappedTargetException{
		XTableChartsSupplier supplier;
		supplier = rtime.queryInterface(
	            XTableChartsSupplier.class,getSheet(sheetName).getxSpreadSheet());
		XTableCharts chartCollection = supplier.getCharts();
        XNameAccess  chartNameAccess = rtime.queryInterface(
            XNameAccess.class, chartCollection );
        return chartNameAccess.getElementNames();

	}
}