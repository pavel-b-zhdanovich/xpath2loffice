package iv.office.generic;

import java.io.IOException;
import java.io.InputStream;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.stream.XMLEventReader;
import javax.xml.stream.XMLInputFactory;
import javax.xml.stream.XMLStreamException;
import javax.xml.stream.events.XMLEvent;
import javax.xml.xpath.XPath;
import javax.xml.xpath.XPathConstants;
import javax.xml.xpath.XPathExpressionException;
import javax.xml.xpath.XPathFactory;

import org.w3c.dom.DOMException;
import org.w3c.dom.NamedNodeMap;
import org.w3c.dom.NodeList;
import org.xml.sax.SAXException;

import com.sun.star.lang.IndexOutOfBoundsException;
import com.sun.star.table.XCellRange;
import com.sun.star.text.XTextTable;

import iv.utils.xml.IXMLSource;

public class XML2Doc {

	IXMLSource xmldoc;
	TextDocument oodoc;
	XMLInputFactory inputFactory;
	InputStream inStream;
	XMLEventReader eventReader;
	XMLEvent event;
	XPath xPath;
	org.w3c.dom.Document xdoc;

	public XML2Doc(IXMLSource p_xmldoc, TextDocument p_oodoc)
			throws XMLStreamException {
		xmldoc = p_xmldoc;
		oodoc = p_oodoc;
		inputFactory = XMLInputFactory.newInstance();
		inStream = xmldoc.toStream();
		eventReader = inputFactory.createXMLEventReader(inStream);

	}

	protected String toXPath(String st) {
		String rst=st.substring(st.indexOf('`')+1);
		return rst.replace('}', '/').replace('$', '@').replace('~', '*')
				.replace('!', ':').replace('^', '&');
	}
	
	void fillBookMarks(){
		for (String bookmark : oodoc.getXPathBookMarks()) {
			try {
				String bmValue = (String) xPath.evaluate(toXPath(bookmark), xdoc);
				oodoc.setBookMarkValue(	bookmark, bmValue);
			} catch (XPathExpressionException | IvOfficeException e) {
				e.printStackTrace();
				continue;
			}
		}
	}
	
	void xPath2Table(XTextTable table, String pth, Integer startRow, String[] xPathRefs) 
									throws IndexOutOfBoundsException, DOMException{
		NodeList nodeList;
		try {
			nodeList = (NodeList) xPath.evaluate(pth, xdoc,
					XPathConstants.NODESET);
		} catch (XPathExpressionException e) {
			IvOfficeException.reportErr(pth);
			return;
		}
		int elementCount = nodeList.getLength();
		oodoc.addRowsToXTable(table, elementCount - 1);
		XCellRange tableAsCellRange = oodoc.rtime.queryInterface(XCellRange.class, table);
		for (int i = 0; i < elementCount; i++) {
			for (int j = 0; j < xPathRefs.length; j++) {
				if (! (xPathRefs[j]+"").isEmpty()) {
					try {
						oodoc.setCellValue(	tableAsCellRange,i+ startRow, 
											j, 
											xPath.evaluate( xPathRefs[j], 
															nodeList.item(i)
															)
											);
					} catch (XPathExpressionException e) {
						IvOfficeException.reportErr(xPathRefs[j]);
						continue;
					}
				}
			}
		}
	}
	
	void xPathAttrs2Table(	XTextTable table, String pth, Integer startRow) 
						{
		NodeList nodeList;
		try {
			nodeList = (NodeList) xPath.evaluate(pth, xdoc,
					XPathConstants.NODESET);
		int elementCount = nodeList.getLength();
		oodoc.addRowsToXTable(table, elementCount - 1);
		} catch (XPathExpressionException e) {
			IvOfficeException.reportErr(pth);
			return;
		}
		int elementCount = nodeList.getLength();
		Integer j=0;
		Integer attrCount = nodeList.item(0).getAttributes().getLength();
		oodoc.setColsForXTable(table, attrCount+1);//extra column for element text()
		for (int i = 0; i < elementCount; i++) {
			NamedNodeMap atrs = nodeList.item(i).getAttributes();
			for (j = 0; j < attrCount; j++) { 
				org.w3c.dom.Node node = atrs.item(j); 
				if (node.getNodeType() == org.w3c.dom.Node.ATTRIBUTE_NODE){
					oodoc.setCellValue(	table, oodoc.getCellName(i+startRow, j),
										node.getNodeValue());} 
					else {
						oodoc.setCellValue(	table,
											oodoc.getCellName(i+startRow, j), 
											nodeList.item(i).getTextContent()
											);
					}
				} 
			String textContent=nodeList.item(i).getTextContent(); 
			if (textContent != null && ! textContent.isEmpty())
				{	oodoc.setCellValue(	table,
										oodoc.getCellName(i+startRow,j),
										textContent);} 
			}
	}
	

	public void execute() throws ParserConfigurationException, SAXException, IOException, IvOfficeTableException, IndexOutOfBoundsException  {
		DocumentBuilderFactory builderFactory = DocumentBuilderFactory
				.newInstance();
		DocumentBuilder builder = builderFactory.newDocumentBuilder();
		xdoc = builder.parse(xmldoc.toStream());
		xPath = XPathFactory.newInstance().newXPath();
		fillBookMarks();
	for (String tblName : oodoc.getXPathTables()) {
			XTextTable table = oodoc.getTableByName(tblName);
			Integer lastRow = table.getRows().getCount()-1;
			Integer colCount = table.getColumns().getCount();
			String[] xPathRefs = new String[colCount];
			boolean hasXpaths=false;
			for (int j = 0; j < colCount; j++) {
				try {
					String xPathRef = oodoc.getCellValue(table, lastRow, j);
					if (! (xPathRef+"").isEmpty() ) {
						xPath.compile(xPathRef);
						hasXpaths=true;
					}
					xPathRefs[j] = xPathRef;
				} catch (XPathExpressionException e) {
				}
			}
			if (hasXpaths)
				{xPath2Table(table, toXPath(tblName), lastRow, xPathRefs);}
				else {	xPathAttrs2Table(table, toXPath(tblName), lastRow);}
			} 
		}
	
}