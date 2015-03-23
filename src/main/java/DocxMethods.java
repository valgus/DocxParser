import org.docx4j.jaxb.XPathBinderAssociationIsPartialException;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.wml.*;
import javax.xml.bind.JAXBElement;
import javax.xml.bind.JAXBException;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.util.ArrayList;
import java.util.List;

public final class DocxMethods {

    public  static WordprocessingMLPackage getTemplate(String name) throws Docx4JException, FileNotFoundException {

        	  WordprocessingMLPackage template = WordprocessingMLPackage.load(new FileInputStream(new File(name)));
        	  return template;
        	 }

    public static List<Object> getAllElementFromObject(Object obj, Class<?> toSearch) {
        List<Object> result = new ArrayList<Object>();
        if (obj instanceof JAXBElement) obj = ((JAXBElement<?>) obj).getValue();

        if (obj.getClass().equals(toSearch))
            result.add(obj);
        else if (obj instanceof ContentAccessor) {
            List<?> children = ((ContentAccessor) obj).getContent();
            for (Object child : children) {
                result.addAll(getAllElementFromObject(child, toSearch));
            }

        }
        return result;
    }


    public static List<Object> createParagraphJAXBNodes(WordprocessingMLPackage template) {
        final String XPATH_TO_SELECT_TEXT_NODES = "//w:p";
         List<Object> jaxbNodes = null;
        try {
            jaxbNodes = template.getMainDocumentPart().getJAXBNodesViaXPath(XPATH_TO_SELECT_TEXT_NODES, true);
        } catch (JAXBException e) {
            e.printStackTrace();
        } catch (XPathBinderAssociationIsPartialException e) {
            e.printStackTrace();
        }
        return jaxbNodes;
    }

    public static List<Object> createTableJAXBNodes(WordprocessingMLPackage template) {
        final String XPATH_TO_SELECT_TEXT_NODES = "//w:tbl";
        List<Object> jaxbNodes = null;
        try {
            jaxbNodes = template.getMainDocumentPart().getJAXBNodesViaXPath(XPATH_TO_SELECT_TEXT_NODES, true);
        } catch (JAXBException e) {
            e.printStackTrace();
        } catch (XPathBinderAssociationIsPartialException e) {
            e.printStackTrace();
        }
        return jaxbNodes;
    }


}
