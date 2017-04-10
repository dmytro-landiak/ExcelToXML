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
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;



public class ExcelToXML {
	
	static XSSFRow row;
	static Cell cell;
	static Iterator < Cell > cellIterator;
	private int count = 0;
	Element 
		smspecpp,
		doc_id,
		DOCTYPE_smspecpp,
		specitem,
		article,
		barcode,
		displayitem,
		itemprice,
		minquantity,
		packsize,
		vatrate,
		name;
	Long tmp;
	Double price;
	String idDoc;

@SuppressWarnings("deprecation")
public void generateXML(File excelFile)throws Exception {
      try {
    	  
    	  DateFormat dateFormat1 = new SimpleDateFormat("yyyy-MM-dd");
    	  DateFormat dateFormat2 = new SimpleDateFormat("HH:mm:ss");
    	  Date today = Calendar.getInstance().getTime();        
    	  String reportDate1 = dateFormat1.format(today);
    	  String reportDate2 = dateFormat2.format(today);

    	 
    	  FileInputStream fis = new FileInputStream(excelFile);
	      XSSFWorkbook workbook = new XSSFWorkbook(fis);
	      XSSFSheet spreadsheet = workbook.getSheetAt(0);
	      
	      for (int i = 1; i <= spreadsheet.getLastRowNum(); i++) {
	    	  XSSFRow row = spreadsheet.getRow(i);
	    	  Iterator < Cell > cellIterator = row.cellIterator();
	    	  	 for (int j = 0; cellIterator.hasNext(); j++){
	    	  		 Cell cell = cellIterator.next();
	    		 	 if (j == 10){
	    		 		idDoc = cell.getStringCellValue();
	    		 	 }
	    	  	 }
	      }
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
 		id.appendChild(doc.createTextNode("PP"+idDoc));
 		postobj.appendChild(id);

 		Element pp = doc.createElement("PP");
 		Element smdoc = doc.createElement("SMDOCUMENTS");
 		pp.appendChild(smdoc);
 		
 			Element id_smdoc = doc.createElement("ID");
 			id_smdoc.appendChild(doc.createTextNode(idDoc));
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
 			createdat.appendChild(doc.createTextNode(reportDate1 + "T" + reportDate2));
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
 	 		id_smcombas.appendChild(doc.createTextNode(idDoc));
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
	 		id_smprovprice.appendChild(doc.createTextNode(idDoc));
	 		smprovprice.appendChild(id_smprovprice);
	 		
	 		Element DOCTYPE_smprovprice = doc.createElement("DOCTYPE");
	 		DOCTYPE_smprovprice.appendChild(doc.createTextNode("PP"));
	 		smprovprice.appendChild(DOCTYPE_smprovprice);

 		postobj.appendChild(pp);
 			      
 				  for (int i = 1; i <= spreadsheet.getLastRowNum(); i++) {
 			    	  count = 0;
 			    	  row = spreadsheet.getRow(i);
 			    	  cellIterator = row.cellIterator();
 			    	  while (cellIterator.hasNext()){
 			    	    cell = cellIterator.next();
 			    	  	if (cell == null || cell.getCellType() == Cell.CELL_TYPE_BLANK)
 			    	  		count ++;
 			    	  }
 			    	     row = spreadsheet.getRow(i);
 			    	     cellIterator = row.cellIterator();
 			    	     if (count < 5) {
 			    	    	 smspecpp = doc.createElement("SMSPECPP");
 		 					 pp.appendChild(smspecpp);
	 			    	  	 for (int j = 0; cellIterator.hasNext(); j++){
	 			    	  		 cell = cellIterator.next();
		 			    		 switch (j) {
		  			               case 0:
		  			            	 doc_id = doc.createElement("DOCID");
		  			            	 doc_id.appendChild(doc.createTextNode(idDoc));
		  			            	 smspecpp.appendChild(doc_id);
		  			            	 
		  			            	 DOCTYPE_smspecpp = doc.createElement("DOCTYPE");
		  			            	 DOCTYPE_smspecpp.appendChild(doc.createTextNode("PP"));
		  			            	 smspecpp.appendChild(DOCTYPE_smspecpp);
		  			            	 
		  			            	 specitem = doc.createElement("SPECITEM");
		  			            	 tmp = (long)cell.getNumericCellValue();
		  			            	 specitem.appendChild(doc.createTextNode(tmp.toString()));
		  			            	 smspecpp.appendChild(specitem);
		  			            	 
		  			            	 displayitem = doc.createElement("DISPLAYITEM");
		  			            	 tmp = (long)cell.getNumericCellValue();
		  			            	 displayitem.appendChild(doc.createTextNode(tmp.toString()));
		  			            	 smspecpp.appendChild(displayitem);
		  			            	 
		  			               break;
		  			               case 1:
		  			            	 article = doc.createElement("ARTICLE");
		  			            	 tmp = (long)cell.getNumericCellValue();
		  			            	 article.appendChild(doc.createTextNode(tmp.toString()));
		  			            	 smspecpp.appendChild(article);
		  			               break;
		  			               case 2:
		  			            	 barcode = doc.createElement("BARCODE");
		  			            	 tmp = (long)cell.getNumericCellValue();
		  			            	 barcode.appendChild(doc.createTextNode(tmp.toString()));
		  			            	 smspecpp.appendChild(barcode);
		  			               break;
		  			               case 3:
		  			            	 name = doc.createElement("NAME");
		  			            	 name.appendChild(doc.createTextNode (cell.getStringCellValue()));
		  			            	 smspecpp.appendChild(name);
		  			               break;
		  			               case 4:
		  			            	 
		  			               break;
		  			               case 5:
		  			            	 packsize = doc.createElement("PACKSIZE");
					            	 tmp = (long)cell.getNumericCellValue();
					            	 packsize.appendChild(doc.createTextNode(tmp.toString()));
					            	 smspecpp.appendChild(packsize);
		  			               break;
		  			               case 6:
		  			            	 minquantity = doc.createElement("MINQUANTITY");
					            	 tmp = (long)cell.getNumericCellValue();
					            	 minquantity.appendChild(doc.createTextNode(tmp.toString()));
					            	 smspecpp.appendChild(minquantity);
					               break;
		  			               case 7:
		  			            	 vatrate = doc.createElement("VATRATE");
		  			            	 vatrate.appendChild(doc.createTextNode("20"));
		  			            	 smspecpp.appendChild(vatrate);
					               break;
		  			               case 8:
	
					               break;
					               
		  			               case 9:
		  			            	 itemprice = doc.createElement("ITEMPRICE");
					            	 price = cell.getNumericCellValue();
					            	 itemprice.appendChild(doc.createTextNode(price.toString()));
					            	 smspecpp.appendChild(itemprice);
					               break;
		  			             }
	 			    	  	 }
 			    	     }
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
}