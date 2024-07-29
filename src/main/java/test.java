import java.util.Calendar;

import org.odftoolkit.odfdom.doc.OdfSpreadsheetDocument;

public class test {

	public static void main(String[] args) throws Exception {
		System.out.println("Hello!!!");
		
		OdfSpreadsheetDocument ods = OdfSpreadsheetDocument.loadDocument("doc/template.ots");
		
		ods.getSpreadsheetTables().get(0).appendRow();
		ods.getSpreadsheetTables().get(0).getRowByIndex(5).getCellByIndex(0).setStringValue("Row3");
		ods.getSpreadsheetTables().get(0).getRowByIndex(5).getCellByIndex(1).setDateValue(Calendar.getInstance());
		ods.getSpreadsheetTables().get(0).getRowByIndex(5).getCellByIndex(2).setDoubleValue(21.59);
		ods.getSpreadsheetTables().get(0).getRowByIndex(5).getCellByIndex(3).setDoubleValue(12.0);
		ods.getSpreadsheetTables().get(0).getRowByIndex(5).getCellByIndex(4).setDoubleValue(0.000005);
		ods.getSpreadsheetTables().get(0).getRowByIndex(5).getCellByIndex(5).setStringValue("Comment3");
		
		ods.getContentRoot().insertBefore(ods.getContentRoot().getChildNodes().item(0).cloneNode(true), null);
		ods.getSpreadsheetTables().get(1).setTableName("newSheet");
		
		ods.save("doc/documet.ods");
		
		System.out.println("Document saved.");
	}

}
