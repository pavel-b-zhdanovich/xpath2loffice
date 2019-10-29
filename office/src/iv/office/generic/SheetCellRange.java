package iv.office.generic;

import com.sun.star.beans.PropertyVetoException;
import com.sun.star.beans.UnknownPropertyException;
import com.sun.star.beans.XPropertySet;
import com.sun.star.container.NoSuchElementException;
import com.sun.star.lang.IllegalArgumentException;
import com.sun.star.lang.IndexOutOfBoundsException;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.sheet.XCellRangeAddressable;
import com.sun.star.sheet.XSheetCellCursor;
import com.sun.star.sheet.XSheetCellRange;
import com.sun.star.sheet.XSpreadsheet;
import com.sun.star.table.CellRangeAddress;
import com.sun.star.table.XCell;
import com.sun.star.table.XCellRange;
import com.sun.star.table.XColumnRowRange;
import com.sun.star.table.XTableColumns;
import com.sun.star.table.XTableRows;
import com.sun.star.util.XMergeable;

public class SheetCellRange {
	private XColumnRowRange range;
	private IOfficeUnoRuntime rtime;
	private String rangeName;
	private XSpreadsheet xsheet;
	private XCellRange as_XCellRange;
	private XPropertySet propSet;

	public class ColumnMergeInfo {
		int startMergeCol;
		int endMergeCol;
		int startMergeRow;
		int endMergeRow;
		String text;
		boolean isMerged;
		public boolean getMerged(){
			return isMerged;
		}
		public int rowCount(){
			return endMergeRow - startMergeRow+1;
		}
		public int colCount(){
			return endMergeCol - startMergeCol+1;
		}
	}


	public SheetCellRange(XSpreadsheet pxsheet, String  pRangeName, IOfficeUnoRuntime prtime) {
		this.rtime = prtime;
		this.xsheet = pxsheet;
		this.rangeName = pRangeName;
		initMe();
	}
	
	public SheetCellRange(XCellRange  cellRange, IOfficeUnoRuntime prtime) {
		as_XCellRange = cellRange;
		this.rtime = prtime;
		range = (XColumnRowRange) 
	            rtime.queryInterface(XColumnRowRange.class, cellRange);

	}
	
	protected void initMe(){
		if (xsheet == null) return;
		range = (XColumnRowRange) 
	            rtime.queryInterface(XColumnRowRange.class, xsheet.getCellRangeByName(rangeName));

	}
	
	public XCellRange asXCellRange(){
		if (this.as_XCellRange == null){
			this.as_XCellRange = rtime.queryInterface(XCellRange.class, range);
		}
		return as_XCellRange;
	}
	
	public void setRows(Integer num){
		if (num <= 0){return;}
		XTableRows rows = (XTableRows) range.getRows();
		
		Integer count = rows.getCount();
		if (count == num) return;
		if (count <= num) {
			rows.insertByIndex(1, num-count);
		}
		else {
			rows.removeByIndex(num, count-num);
		}
		initMe();
	}
	
	public int rowCount(){
		return range.getRows().getCount();
	}
	
	public int colCount(){
		return range.getColumns().getCount();
	}
	
	
	public boolean isACell(){
		return ((range.getColumns().getCount() == 1) && (range.getRows().getCount() == 1));
	}
	
	
	public void setCols(Integer num){
		if (num <= 0){return;}
		XTableColumns cols = range.getColumns();
		Integer count = cols.getCount();
		if (count <= num) 
			cols.insertByIndex(1, num-count);
		else {
			cols.removeByIndex(num, count-num);
		}
		initMe();
	}
	
	public void setText(Integer row, Integer col, String theText) throws IndexOutOfBoundsException{
		asXCellRange().getCellByPosition(col, row).setFormula(theText);
	}
	

	
	public void merge(Integer row, Integer col, Integer rowsToMerge, Integer colsToMerge) throws IndexOutOfBoundsException{
		XCellRange subRange = asXCellRange().getCellRangeByPosition(col, row, col+colsToMerge-1, row+rowsToMerge-1);
		XMergeable m = rtime.queryInterface(XMergeable.class, subRange);
		m.merge(true);
	}
	
	public void mergeAndSetText(Integer row, Integer col, Integer rowsToMerge, Integer colsToMerge, String theText) throws IndexOutOfBoundsException{
		XCellRange subRange = asXCellRange().getCellRangeByPosition(col, row, col+colsToMerge-1, row+rowsToMerge-1);
		XCell cell = asXCellRange().getCellByPosition(col, row);
		XMergeable m = rtime.queryInterface(XMergeable.class, subRange);
		m.merge(true);
		cell.setFormula(theText);
	}
	
	public String getCellXPath() throws IndexOutOfBoundsException{
		return asXCellRange().getCellByPosition(0,0).getFormula();
	}
	
	public String getRangeXPath() throws IndexOutOfBoundsException{
		return asXCellRange().getCellByPosition(0,1).getFormula();
	}
	public void clearRangeXPath() throws IndexOutOfBoundsException{
		asXCellRange().getCellByPosition(0,1).setFormula("");
	}
	
	public SpreadSheet getSpreadSheet() throws NoSuchElementException, WrappedTargetException{
		XSheetCellRange sheetCellRange = rtime.queryInterface(XSheetCellRange.class, range);
		return new SpreadSheet(sheetCellRange.getSpreadsheet(),rtime);
	}
	
	public ColumnMergeInfo getMergedInfo(Integer row, Integer col) throws IndexOutOfBoundsException, NoSuchElementException, WrappedTargetException{
		ColumnMergeInfo result =  new ColumnMergeInfo();
		XCell c = asXCellRange().getCellByPosition(col,row);
		result.isMerged = rtime.queryInterface(	XMergeable.class,c).getIsMerged();
		XSheetCellRange mergedSubRange = rtime.queryInterface(XSheetCellRange.class, c);
		XSheetCellCursor xCursor = getSpreadSheet().getxSpreadSheet().createCursorByRange(mergedSubRange);
		xCursor.collapseToMergedArea();
		XCellRangeAddressable xCellRangeAddressable = rtime.queryInterface(XCellRangeAddressable.class, xCursor);
		CellRangeAddress addr = xCellRangeAddressable.getRangeAddress();
		result.startMergeRow = addr.StartRow;
		result.endMergeRow = addr.EndRow;
		result.startMergeCol = addr.StartColumn;
		result.endMergeCol = addr.EndColumn;
		result.text = c.getFormula();
		return result;
	}
	
	public String getFormulaAt(Integer row, Integer col) throws IndexOutOfBoundsException{
		return asXCellRange().getCellByPosition(col,row).getFormula();
	}

	public void setStyle(Integer row, Integer col, String theStyle) throws IndexOutOfBoundsException, UnknownPropertyException, PropertyVetoException, IllegalArgumentException, WrappedTargetException{
		XCell xcell = asXCellRange().getCellByPosition(col, row);
		XPropertySet xset = rtime.queryInterface(XPropertySet.class, xcell);
		xset.setPropertyValue("CellStyle", theStyle);
	}
	
	protected Object getPropertyByName(String propName) throws UnknownPropertyException, WrappedTargetException{
		if (propSet == null){
			propSet = (XPropertySet) rtime.queryInterface(XPropertySet.class, as_XCellRange);
		}
		return propSet.getPropertyValue(propName);
	}
	
	public com.sun.star.awt.Point getPosition() throws UnknownPropertyException, WrappedTargetException{
		return (com.sun.star.awt.Point)getPropertyByName("Position");
	}
	
}
