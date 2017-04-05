package com.landim;

import org.apache.poi.xssf.usermodel.*;
import org.w3c.dom.*;
import java.io.*;
import javax.xml.parsers.*;
import javax.xml.transform.*;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import java.io.File;
import java.io.FileInputStream;
import java.util.Iterator;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;


public class ExcelToXML {
	
	static XSSFRow row;
	
   public void generateXML(File excelFile)throws Exception {
      try { 
    	  
    	  
         DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
         DocumentBuilder builder = factory.newDocumentBuilder();
         Document doc = builder.newDocument();
        
         Element rootElement = doc.createElement("PACKAGE");
 		 doc.appendChild(rootElement);
 		 Attr attrRootEl = doc.createAttribute("name");
 		 attrRootEl.setValue("Default name");
 		 rootElement.setAttributeNode(attrRootEl);

 		Element postobj = doc.createElement("POSTOBJECT");
 		rootElement.appendChild(postobj);
 		postobj.setAttribute("description", "Прайс-лист поставщика");
 		postobj.setAttribute("action", "normal");

 		Element id = doc.createElement("Id");
 		id.appendChild(doc.createTextNode("PP0000000007"));
 		postobj.appendChild(id);

 		Element pp = doc.createElement("PP");
 		Element smdoc = doc.createElement("SMDOCUMENTS");
 		pp.appendChild(smdoc);
 		
 			Element id_smdoc = doc.createElement("ID");
 			id_smdoc.appendChild(doc.createTextNode("0000000007"));
 			smdoc.appendChild(id_smdoc);

 			Element doctype = doc.createElement("DOCTYPE");
 			doctype.appendChild(doc.createTextNode("PP"));
 			smdoc.appendChild(doctype);

 			Element bornin = doc.createElement("BORNIN");
 			bornin.appendChild(doc.createTextNode("UoZgZcsuT0O50LCmsuQ2cg=="));
 			smdoc.appendChild(bornin);

 			Element clientind = doc.createElement("CLIENTINDEX");
 			clientind.appendChild(doc.createTextNode("4"));
 			smdoc.appendChild(clientind);
 			
 			Element createdat = doc.createElement("CREATEDAT");
 			createdat.appendChild(doc.createTextNode("2017-03-29T00:00:00"));
 			smdoc.appendChild(createdat);
 			
 			Element CURRENCYMULTORDER = doc.createElement("CURRENCYMULTORDER");
 			CURRENCYMULTORDER.appendChild(doc.createTextNode("0"));
 			smdoc.appendChild(CURRENCYMULTORDER);
 			
 			Element CURRENCYRATE = doc.createElement("CURRENCYRATE");
 			CURRENCYRATE.appendChild(doc.createTextNode("1"));
 			smdoc.appendChild(CURRENCYRATE);
 			
 			Element CURRENCYTYPE = doc.createElement("CURRENCYTYPE");
 			CURRENCYTYPE.appendChild(doc.createTextNode("1"));
 			smdoc.appendChild(CURRENCYTYPE);
 			
 			Element DOCSTATE = doc.createElement("DOCSTATE");
 			DOCSTATE.appendChild(doc.createTextNode("1"));
 			smdoc.appendChild(DOCSTATE);

 			Element ISROUBLES = doc.createElement("ISROUBLES");
 			ISROUBLES.appendChild(doc.createTextNode("1"));
 			smdoc.appendChild(ISROUBLES);
 			
 			Element OPCODE = doc.createElement("OPCODE");
 			OPCODE.appendChild(doc.createTextNode("-1"));
 			smdoc.appendChild(OPCODE);
 			
 			Element PRICEROUNDMODE = doc.createElement("PRICEROUNDMODE");
 			PRICEROUNDMODE.appendChild(doc.createTextNode("1"));
 			smdoc.appendChild(PRICEROUNDMODE);
 			
 			Element TOTALSUM = doc.createElement("TOTALSUM");
 			TOTALSUM.appendChild(doc.createTextNode("0"));
 			smdoc.appendChild(TOTALSUM);
 			
 			Element TOTALSUMCUR = doc.createElement("TOTALSUMCUR");
 			TOTALSUMCUR.appendChild(doc.createTextNode("0"));
 			smdoc.appendChild(TOTALSUMCUR);
 			
 		Element smcombas = doc.createElement("SMCOMMONBASES");
 	 	pp.appendChild(smcombas);
 	 		
 	 		Element id_smcombas = doc.createElement("ID");
 	 		id_smcombas.appendChild(doc.createTextNode("0000000007"));
 	 		smcombas.appendChild(id_smcombas);
			
			Element DOCTYPE_smcombas = doc.createElement("DOCTYPE");
			DOCTYPE_smcombas.appendChild(doc.createTextNode("PP"));
			smcombas.appendChild(DOCTYPE_smcombas);
			
			Element BASEDOCTYPE = doc.createElement("BASEDOCTYPE");
			BASEDOCTYPE.appendChild(doc.createTextNode("CO"));
			smcombas.appendChild(BASEDOCTYPE);
			
			Element BASEID = doc.createElement("BASEID");
			BASEID.appendChild(doc.createTextNode("0000000002"));
			smcombas.appendChild(BASEID);
 		
 			
		Element smprovprice = doc.createElement("SMPROVIDERPRICE");
	 	pp.appendChild(smprovprice);
	 	 	
	 		Element id_smprovprice = doc.createElement("ID");
	 		id_smprovprice.appendChild(doc.createTextNode("0000000007"));
	 		smprovprice.appendChild(id_smprovprice);
	 		
	 		Element DOCTYPE_smprovprice = doc.createElement("DOCTYPE");
	 		DOCTYPE_smprovprice.appendChild(doc.createTextNode("PP"));
	 		smprovprice.appendChild(DOCTYPE_smprovprice);
	 	
	 		
	 	/***************CONTINUE HERE*****************************/	
 		postobj.appendChild(pp);
 		
 		FileInputStream fis = new FileInputStream(excelFile);
 			      XSSFWorkbook workbook = new XSSFWorkbook(fis);
 			      XSSFSheet spreadsheet = workbook.getSheetAt(0);
 			      Iterator <Row> rowIterator = spreadsheet.iterator();
 			      while (rowIterator.hasNext()) 
 			      {
 			         row = (XSSFRow) rowIterator.next();
 			         Iterator < Cell > cellIterator = row.cellIterator();
 			         while ( cellIterator.hasNext()) 
 			         {
 			            Cell cell = cellIterator.next();
 			            switch (cell.getCellType()) 
 			            {
 			               case Cell.CELL_TYPE_NUMERIC:
 			               System.out.print( 
 			               cell.getNumericCellValue() + " \t\t " );
 			               break;
 			               case Cell.CELL_TYPE_STRING:
 			               System.out.print(
 			               cell.getStringCellValue() + " \t\t " );
 			               break;
 			            }
 			         }
 			         System.out.println();
 			      }
 			      workbook.close();
 			      fis.close();

         
        TransformerFactory transformerFactory = TransformerFactory.newInstance();
 		Transformer transformer = transformerFactory.newTransformer();
 		transformer.setOutputProperty
        (OutputKeys.INDENT, "yes");
        transformer.setOutputProperty(
           "{http://xml.apache.org/xslt}indent-amount", "2");
 		
 		DOMSource source = new DOMSource(doc);
 		StreamResult result = new StreamResult(new File("Output.xml"));
 		transformer.transform(source, result);

 		System.out.println("Completed!");
      } catch (IOException e) {
         System.out.println("IOException " + e.getMessage());
      } catch (ParserConfigurationException e) {
         System.out
            .println("ParserConfigurationException " + e.getMessage());
      } catch (TransformerConfigurationException e) {
         System.out.println("TransformerConfigurationException "+ e.getMessage());
      } catch (TransformerException e) {
         System.out.println("TransformerException " + e.getMessage());
      }
   }

   public static void main(String[] argv) throws Exception {
      ExcelToXML excel = new ExcelToXML();
      File input = new File("test.xls");
      excel.generateXML(input);	
   }
}