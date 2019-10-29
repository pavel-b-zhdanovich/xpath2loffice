package iv.office.generic;

import java.io.IOException;
import java.util.ArrayList;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.xpath.XPath;
import javax.xml.xpath.XPathConstants;
import javax.xml.xpath.XPathExpressionException;
import javax.xml.xpath.XPathFactory;

import org.w3c.dom.NodeList;
import org.xml.sax.SAXException;

import com.sun.star.beans.UnknownPropertyException;
import com.sun.star.container.NoSuchElementException;
import com.sun.star.lang.IndexOutOfBoundsException;
import com.sun.star.lang.WrappedTargetException;

import iv.utils.xml.IXMLSource;

public class XML2Ods {
	
	private String preserveRowsSuffix = "__iv_preserve_rows";

	IXMLSource xmldoc;
	CalcDocument oodoc;
	org.w3c.dom.Document xdoc;
	XPath xPath;

	public XML2Ods(IXMLSource p_xmldoc, CalcDocument pdoc)
			throws ParserConfigurationException, SAXException, IOException {
		xmldoc = p_xmldoc;
		oodoc = pdoc;
		DocumentBuilderFactory builderFactory = DocumentBuilderFactory
				.newInstance();
		DocumentBuilder builder = builderFactory.newDocumentBuilder();
		xdoc = builder.parse(xmldoc.toStream());
		xPath = XPathFactory.newInstance().newXPath();
	}
	
	public void xPath2RangeMerged(SheetCellRange range) 
			throws IndexOutOfBoundsException, NoSuchElementException, WrappedTargetException{
		this.xPath2RangeMerged(range,false);
	}
	
	public void xPath2RangeMerged(SheetCellRange range, boolean preserveRows) 
			throws IndexOutOfBoundsException, NoSuchElementException, WrappedTargetException{
		String pth = range.getRangeXPath();
		try {
			NodeList nodeList = (NodeList) xPath.evaluate(pth, xdoc,XPathConstants.NODESET);
			range.clearRangeXPath();
			xPath2RangeMerged(nodeList, range, preserveRows);
			
		} catch (XPathExpressionException e) {
			IvOfficeException.reportErr(pth);
		}
	}
	
	public void xPath2RangeMerged(NodeList nodeList, SheetCellRange range, boolean preserveRows) 
			throws IndexOutOfBoundsException, NoSuchElementException, WrappedTargetException{

		Integer elementCount = nodeList.getLength();
		ArrayList<SheetCellRange.ColumnMergeInfo> rangeCols = 
				new ArrayList<SheetCellRange.ColumnMergeInfo>(1);
		int colCount = range.colCount();
		int colNum = 0;
		int rowNum = 0;
		do {
			SheetCellRange.ColumnMergeInfo mergeInfo = range.getMergedInfo(
					rowNum, colNum);
			try {
				xPath.compile(mergeInfo.text+"");
				range.setText(rowNum, colNum, "");
			} catch (XPathExpressionException e) {
				mergeInfo.text = "";
			}
			rangeCols.add(mergeInfo);
			colNum = mergeInfo.endMergeCol - rangeCols.get(0).startMergeCol + 1;
		} while (colNum < colCount);
		if (! preserveRows) range.setRows(elementCount);

		for (int i = 0; i < elementCount; i++) {
			org.w3c.dom.Node node = nodeList.item(i);
			colNum = 0; 
			int firstColDelta=rangeCols.get(0).startMergeCol - 1;
			for (int j = 0; j < rangeCols.size(); j++) {
				SheetCellRange.ColumnMergeInfo mi = rangeCols.get(j);
				String relativeExpr = mi.text+"";
				if ((relativeExpr == null) || (relativeExpr.isEmpty()))
					range.merge(i, colNum, 1, mi.colCount());
				else
					try {
						String relativeExprValue = xPath.evaluate(relativeExpr, node);
						range.mergeAndSetText(i, colNum, 1, mi.colCount(),
								relativeExprValue+"");
					} catch (NullPointerException | XPathExpressionException e) {
						IvOfficeException.reportErr(relativeExpr);
					}
				colNum = mi.endMergeCol - firstColDelta;
			}
		}

	}
	
	public void xPathExpr2Cell(String pth, String sheetName,
			String cellName) throws NoSuchElementException, WrappedTargetException, Exception {
		oodoc.getSheet(sheetName).printtoCell(cellName, xPath.evaluate(pth, xdoc));

	}
	public void xPathExpr2Cell(SheetCellRange range)
			throws IndexOutOfBoundsException {
		String attrPath = range.getCellXPath();
		try {
			range.setText(0, 0, xPath.evaluate(attrPath, xdoc));
		} catch (XPathExpressionException e) {
			IvOfficeException.reportErr(attrPath);
		}

	}

	public void xml2Ranges() {
		String[] names;
		try {
			names = oodoc.getRangeNames();
			for (int i = 0; i < names.length; i++) {
				SheetCellRange range = new SheetCellRange(oodoc.getReferrer(
						names[i]).getReferredCells(), oodoc.rtime);
				if (range.isACell()) {
					xPathExpr2Cell(range);
				} else {
					xPath2RangeMerged(range, names[i].endsWith(preserveRowsSuffix));
				}
			}
		} catch (	UnknownPropertyException | 
					WrappedTargetException | 
					NoSuchElementException | 
					IndexOutOfBoundsException e) {
			e.printStackTrace();
		}
	}

	public void xml2Ranges(String sheetName) throws NoSuchElementException, WrappedTargetException {
		String[] names;
		SpreadSheet sheet = oodoc.getSheet(sheetName);
		try {
			names = sheet.getRanges();
			for (int i = 0; i < names.length; i++) {
				SheetCellRange range = oodoc.getRange(sheetName, names[i]);
				try {
					if (range.isACell()) {
						xPathExpr2Cell(range);
					} else {
						xPath2RangeMerged(range);
					}
				} catch (Exception e) {System.err.println("Ошибка заполнения диапазона "+names[i]);}
			}
		} catch (	UnknownPropertyException |
					WrappedTargetException |
					NoSuchElementException e) {
			e.printStackTrace();
		} 
	}

}
