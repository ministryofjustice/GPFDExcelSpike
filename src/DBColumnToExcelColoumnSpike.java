import java.io.IOException;
import java.util.HashMap;
import java.util.Map;


import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;

import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;
import org.xml.sax.SAXException;

public class DBColumnToExcelColoumnSpike {

	public static void main(String[] args) {
		try {
			new DBColumnToExcelColoumnSpike().parseToMap("./ColumnMappingConfig.xml");
		} catch (ParserConfigurationException | SAXException | IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

	}
	
	Map<String, MappingClass> parseToMap(String columMpaaingFile) throws ParserConfigurationException, SAXException, IOException {
		
		
		Map<String, MappingClass> mappingFields = new HashMap<String, MappingClass>();
 
    


      DocumentBuilderFactory dbFactory = DocumentBuilderFactory.newInstance();
      DocumentBuilder dBuilder = dbFactory.newDocumentBuilder();
      org.w3c.dom.Document doc = dBuilder.parse(columMpaaingFile);

      //optional, but recommended
      //read this - http://stackoverflow.com/questions/13786607/normalization-in-dom-parsing-with-java-how-does-it-work
      doc.getDocumentElement().normalize();


      NodeList nList = doc.getElementsByTagName("field");

      System.out.println("----------------------------");

      for (int temp = 0; temp < nList.getLength(); temp++) {

         Node nNode = nList.item(temp);


         if (nNode.getNodeType() == Node.ELEMENT_NODE) {

            Element eElement = (Element) nNode;


            MappingClass mappingClass = new MappingClass();
            mappingClass.setSourceName(eElement.getAttribute("sourceName"));
            mappingClass.setMappedName(eElement.getAttribute("mappedName"));
            if(eElement.getAttribute("dateFormat")!=null)
            	mappingClass.setDateFormat(eElement.getAttribute("dateFormat"));
            else
            	mappingClass.setDateFormat(null);
            if(eElement.getAttribute("numberFormat")!=null)
            	mappingClass.setNumberFormat(eElement.getAttribute("numberFormat"));
            else
            	mappingClass.setNumberFormat(null);;
           
            mappingFields.put(mappingClass.getSourceName(), mappingClass);

         }
      }
      System.out.println( "Done");
         mappingFields.forEach((K, V) -> {
         System.out.println( "Done K = " +  K + " V = " + V.toString());
      });

         return mappingFields;
	}


}
