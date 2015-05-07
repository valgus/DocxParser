package Model;

import org.docx4j.Docx4J;
import org.docx4j.convert.out.FOSettings;
import org.docx4j.fonts.IdentityPlusMapper;
import org.docx4j.fonts.Mapper;
import org.docx4j.fonts.PhysicalFont;
import org.docx4j.fonts.PhysicalFonts;
import org.docx4j.jaxb.Context;
import org.docx4j.jaxb.XPathBinderAssociationIsPartialException;
import org.docx4j.model.fields.FieldUpdater;
import org.docx4j.model.structure.PageDimensions;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.*;

import javax.xml.bind.JAXBElement;
import javax.xml.bind.JAXBException;
import java.awt.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.OutputStream;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.List;

public final class DocxMethods {

    public  static WordprocessingMLPackage getTemplate(String name) throws Docx4JException, FileNotFoundException {

        	  WordprocessingMLPackage template = WordprocessingMLPackage.load(new FileInputStream(new File(name)));
        	  return template;
        	 }

    public static List<Object> getAllElementFromObject(Object obj, Class<?> toSearch) {
        List<Object> result = new ArrayList<>();
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


    public static List<Object> createSdtJAXBNodes(WordprocessingMLPackage template) {
        final String XPATH_TO_SELECT_TEXT_NODES = "//w:sdt";
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

    public static int getIndexOfParagraph (MainDocumentPart mainPart, P p) {
        return mainPart.getContent().indexOf(p);
    }

    public static P getParagraphFromIndex (MainDocumentPart mainPart,int i) {
        return (P)mainPart.getContent().get(i);
    }

    public void setPageMargins(WordprocessingMLPackage word) {
        ObjectFactory factory = Context.getWmlObjectFactory();
        try {
            Body body = word.getMainDocumentPart().getContents().getBody();
            PageDimensions page = new PageDimensions();
            SectPr.PgMar pgMar = page.getPgMar();
            pgMar.setBottom(BigInteger.valueOf(pixelsToDxa(50)));
            pgMar.setTop(BigInteger.valueOf(pixelsToDxa(50)));
            pgMar.setLeft(BigInteger.valueOf(pixelsToDxa(50)));
            pgMar.setRight(BigInteger.valueOf(pixelsToDxa(50)));
            SectPr sectPr = factory.createSectPr();
            body.setSectPr(sectPr);
            sectPr.setPgMar(pgMar);
        } catch (Docx4JException e) {
            e.printStackTrace();
        }
    }

    protected static int getDPI() {
        return GraphicsEnvironment.isHeadless() ? 96 :
                Toolkit.getDefaultToolkit().getScreenResolution();
    }

    private int pixelsToDxa(int pixels) {
        return  ( 1440 * pixels / getDPI() );
    }
}
