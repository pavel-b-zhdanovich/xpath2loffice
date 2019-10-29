package iv.office.generic;

import java.util.ArrayList;

import com.sun.star.beans.PropertyValue;
import com.sun.star.container.NoSuchElementException;
import com.sun.star.container.XNameAccess;
import com.sun.star.frame.XStorable;
import com.sun.star.lang.IndexOutOfBoundsException;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.table.XCellRange;
import com.sun.star.table.XTableColumns;
import com.sun.star.table.XTableRows;
import com.sun.star.text.XBookmarksSupplier;
import com.sun.star.text.XText;
import com.sun.star.text.XTextContent;
import com.sun.star.text.XTextDocument;
import com.sun.star.text.XTextRange;
import com.sun.star.text.XTextTable;
import com.sun.star.text.XTextTablesSupplier;

public class TextDocument {

	public XTextDocument txtDoc;
	protected IOfficeUnoRuntime rtime;
	private Character[] letters;
			
	public TextDocument(XTextDocument doc, IOfficeUnoRuntime prtime) {
		txtDoc = doc;
		rtime=prtime;
		int i=0;
		letters = new Character[27];
		for (char j = 'A'; j <= 'Z'; j++) {
			letters[i++] = j;
			}
	}
	
	public void save(String storeUrl) throws Exception{
		XStorable xStorable = (XStorable)rtime.queryInterface(XStorable.class, txtDoc);
	    PropertyValue[] storeProps = new PropertyValue[1];
	    storeProps[0] = new PropertyValue();
	    storeProps[0].Name = "FilterName";
	    storeProps[0].Value = "MS Word 97"; 
	    xStorable.storeAsURL(storeUrl, storeProps); 
	}
	 
protected XNameAccess getTables(){
	XTextTablesSupplier xTablesSupplier = (XTextTablesSupplier) rtime.queryInterface(
            XTextTablesSupplier.class, txtDoc);
	return xTablesSupplier.getTextTables();
	}
	
protected XNameAccess getBookmarks(){
	XBookmarksSupplier bmSupplier = (XBookmarksSupplier) rtime.queryInterface(
			XBookmarksSupplier.class, txtDoc);
	return bmSupplier.getBookmarks();
	}

public XTextTable getTableByName(String tabName) throws IvOfficeTableException {
	XNameAccess xTables = getTables();
	XTextTable table;
	try {
		table = (XTextTable)rtime.queryInterface(XTextTable.class, xTables.getByName(tabName));
	} catch (NoSuchElementException | WrappedTargetException e) {
		throw new IvOfficeTableException(e.getMessage());
	}
	return table;
	}

public void setRowsForXTable(XTextTable table, Integer num){
	if (num<=0)
		{return;}
	XTableRows rows = table.getRows();
	Integer count = rows.getCount();
	if (count <= num) 
		rows.insertByIndex(count, num-count);
	else rows.removeByIndex(num, count-num);
	}

public void setColsForXTable(XTextTable table, Integer num){
	if (num<=0)
		{return;}
	XTableColumns cols = rtime.queryInterface(XTableColumns.class,table.getColumns());
	Integer count = cols.getCount();
	if (count <= num) 
		cols.insertByIndex(count, num-count);
	else cols.removeByIndex(num, count-num);
	
	}

public void addRowsToXTable(XTextTable table, Integer num){
	if (num<=0)
		{return;}
	XTableRows rows = table.getRows();
	rows.insertByIndex(rows.getCount(), num);
	}

public void addRowsToTable(String tablename, Integer num) throws Exception{
	if (num<=0)
		{return;}
	XTextTable table = getTableByName(tablename);
	XTableRows rows = table.getRows();
	rows.insertByIndex(rows.getCount(), num);
	}

public void setCellValue(XTextTable table, String cellName, String value){
	XText cellText = (XText)rtime.queryInterface(XText.class,table.getCellByName(cellName));
	cellText.setString(value);
	}

public void setCellValue(XCellRange rng, int row, int col, String value) throws IndexOutOfBoundsException{
	XText cellText = (XText)rtime.queryInterface(XText.class,rng.getCellByPosition(col,row));
	cellText.setString(value);
	}

public String getCellValue(XTextTable table, int row, int col)throws IndexOutOfBoundsException{
	XText cellText = (XText)rtime.queryInterface(XText.class,table.getCellByName(getCellName(row, col)));
	return cellText.getString();
	}

public String getCellValue(XCellRange rng, int row, int col)throws IndexOutOfBoundsException{
	XText cellText = (XText)rtime.queryInterface(XText.class,rng.getCellByPosition(col,row));
	return cellText.getString();
	}

public void setBookMarkValue(String bmName, String value) throws IvOfficeException {
	
	XTextContent xtc;
	try {
		xtc = (XTextContent)rtime.queryInterface(	XTextContent.class, 
																	this.getBookmarks().getByName(bmName)
																	);
	} catch (NoSuchElementException | WrappedTargetException e) {
		throw new IvOfficeException(e.getMessage());
	}
	XTextRange xtr = (XTextRange) rtime.queryInterface(XTextRange.class,xtc.getAnchor());
	xtr.setString(value);
	}

public String getStringAtBookMark(String bmName) throws IvOfficeException{
	XTextContent xtc;
	try {
		xtc = (XTextContent)rtime.queryInterface(	XTextContent.class, 
				this.getBookmarks().getByName(bmName));
	} catch (NoSuchElementException | WrappedTargetException e) {
		throw new IvOfficeException(e.getMessage());
	}
	XTextRange xtr = (XTextRange) rtime.queryInterface(XTextRange.class,xtc.getAnchor());
	return xtr.getString();
	}

public ArrayList<String> getXPathTables(){
	ArrayList<String> result = new ArrayList<String>();
	for (String tabName : getTables().getElementNames()) {
		if ((tabName.charAt(0) == '}')||(tabName.charAt(0) == '/')){
			result.add(tabName);
			}
		}
	return result;
	}
	

public ArrayList<String> getXPathBookMarks(){
	ArrayList<String> result = new ArrayList<String>();
	for (String bmName : getBookmarks().getElementNames()) {
		result.add(bmName);
		}
	return result;
	}


public String getCellName(int row, int col){
	String result="";
	if (col > 26)
		{result=result+letters[col / 26];}
	result=result+ letters[(col) % 26]+(row+1);
	return result;
	}

}


