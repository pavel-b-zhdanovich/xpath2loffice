package iv.office.test;



import java.io.File;
import java.util.Date;

import iv.office.generic.CalcDocument;
import iv.office.generic.Office;
import iv.office.generic.TextDocument;
import iv.office.generic.XML2Doc;
import iv.office.generic.XML2Ods;
import iv.office.runtime.OfficeRuntime;
import iv.utils.xml.XMLFileSource;

public class OofficeTest {

	/**
	 * @param args
	 * @throws java.lang.Exception 
	 */
	
	static Runtime rt;

	protected static void t4_2() throws java.lang.Exception{
		Office office = new Office(new OfficeRuntime());
		String userDir=System.getProperty("user.dir");
		java.net.URI userDirAsURI = new File(userDir).toURI();
		TextDocument doc = office.openTextDocument
				(userDirAsURI+"/examples/testResults.odt",Boolean.FALSE);
		XML2Doc c = new XML2Doc(new XMLFileSource(
				new File(userDir+"/examples/testresults.xml")), doc);		
		c.execute(); 
		doc.save(userDirAsURI+"/tmp/modifiedDoc.odt");
		}
	
	protected static void t14()throws java.lang.Exception{
		Office office = new Office(new OfficeRuntime(),Office.VISIBLE);
		String userDir=System.getProperty("user.dir");
		java.net.URI userDirAsURI = new File(userDir).toURI();
		
		CalcDocument doc = office.openCalcDocument
				(userDirAsURI+"/examples/2_ndfl.xls",Boolean.FALSE);
		doc.save(userDirAsURI+"/tmp/modified.ods");
		XML2Ods c = new XML2Ods(new XMLFileSource(
				new File(userDir+"/examples/ndfl.xml")),doc);
		c.xml2Ranges();
		doc.save();
	}
	
	public static void main(String[] args) throws java.lang.Exception {
		System.out.println("Start at "+new Date().toString());
		t4_2();
		System.out.println("Finish at "+new Date().toString());
	}

}
