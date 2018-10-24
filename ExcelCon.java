import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.DocumentBuilder;
import org.w3c.dom.Document;
import org.w3c.dom.NodeList;
import org.w3c.dom.Node;
import org.w3c.dom.Element;
import java.io.File;
import java.util.*;
import java.io.FileOutputStream;


import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.usermodel.*;




public class ExcelCon{

	private static Workbook workbook;
	private static int rowNum;

  public static void main(String argv[])throws Exception {
    try {
    
    initXls();

    Sheet sheet = workbook.getSheetAt(0);
    File fXmlFile = new File("status.xml");
    DocumentBuilderFactory dbFactory = DocumentBuilderFactory.newInstance();
    DocumentBuilder dBuilder = dbFactory.newDocumentBuilder();
    Document doc = dBuilder.parse(fXmlFile);
    doc.getDocumentElement().normalize();

    System.out.println("Root element :" + doc.getDocumentElement().getNodeName());
    NodeList nList = doc.getElementsByTagName("samp");
    System.out.println("----------------------------");
	List<Sample> lis = new ArrayList<Sample>();
 
    for (int temp = 0; temp < nList.getLength(); temp++) {
        Node nNode = nList.item(temp);
        System.out.println("\nCurrent Element :" + nNode.getNodeName());
        if (nNode.getNodeType() == Node.ELEMENT_NODE) {
         Element eElement = (Element) nNode;
         String source= eElement.getElementsByTagName("source").item(0).getTextContent();
         String destination = eElement.getElementsByTagName("destination").item(0).getTextContent();
         String status = eElement.getElementsByTagName("status").item(0).getTextContent();
         Sample s = new Sample(eElement.getAttribute("status"),eElement.getAttribute("source"),eElement.getAttribute("destination"));
         lis.add(s);


	Row row =sheet.createRow(rowNum++);
        Cell cell = row.createCell(0);
        cell.setCellValue(source);

        cell = row.createCell(1);
        cell.setCellValue(destination);

        cell = row.createCell(2);
        cell.setCellValue(status);
     
        }
    }
    } catch (Exception e) {
    e.printStackTrace();
    }

 FileOutputStream fileOut = new FileOutputStream("C://Temp//Excel-Out.xls");
 workbook.write(fileOut);
 workbook.close();
 fileOut.close();

       
  }


 private static void initXls() {
        workbook = new HSSFWorkbook();

        CellStyle style = workbook.createCellStyle();
        Font boldFont = workbook.createFont();
        boldFont.setBold(true);
        style.setFont(boldFont);
        //style.setAlignment(CellStyle.ALIGN_CENTER);

        Sheet sheet = workbook.createSheet();
        rowNum = 0;
        Row row = sheet.createRow(rowNum++);
        Cell cell = row.createCell(0);
        cell.setCellValue("Source");
        cell.setCellStyle(style);

        cell = row.createCell(1);
        cell.setCellValue("destination");
        cell.setCellStyle(style);

        cell = row.createCell(2);
        cell.setCellValue("Status");
        cell.setCellStyle(style);
        
    }

}







class Sample{
private String status;
private String source;
private String destination;

Sample(String status,String source,String destination){
this.status = status;
this.source = source;
this.destination = destination;
}


public String getStatus(){
return this.status;
}

}