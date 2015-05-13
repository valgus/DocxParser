package Model;


import org.docx4j.TraversalUtil;
import org.docx4j.XmlUtils;
import org.docx4j.convert.out.flatOpcXml.FlatOpcXmlCreator;
import org.docx4j.dml.wordprocessingDrawing.Inline;
import org.docx4j.finders.SectPrFinder;
import org.docx4j.jaxb.Context;
import org.docx4j.jaxb.XPathBinderAssociationIsPartialException;
import org.docx4j.model.structure.PageDimensions;
import org.docx4j.model.structure.SectionWrapper;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.exceptions.InvalidFormatException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.Part;
import org.docx4j.openpackaging.parts.WordprocessingML.BinaryPartAbstractImage;
import org.docx4j.openpackaging.parts.WordprocessingML.FooterPart;
import org.docx4j.openpackaging.parts.WordprocessingML.HeaderPart;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.openpackaging.parts.relationships.Namespaces;
import org.docx4j.relationships.Relationship;
import org.docx4j.wml.*;

import javax.xml.bind.JAXBElement;
import javax.xml.bind.JAXBException;
import java.awt.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.List;

public final class DocxMethods {

    private static ObjectFactory factory = Context.getWmlObjectFactory();

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

    public static void setPageMargins(WordprocessingMLPackage word) {
        ObjectFactory factory = Context.getWmlObjectFactory();
        try {
            Body body = word.getMainDocumentPart().getContents().getBody();
            PageDimensions page = new PageDimensions();
            SectPr.PgMar pgMar = page.getPgMar();

            pgMar.setBottom(BigInteger.valueOf((1135)));
            pgMar.setTop(BigInteger.valueOf((1135)));
            pgMar.setLeft(BigInteger.valueOf((1700)));
            pgMar.setRight(BigInteger.valueOf((567)));
            SectPr.PgSz size = page.getPgSz();
            size.setH(BigInteger.valueOf(16838));
            size.setW(BigInteger.valueOf(11906));
            SectPr sectPr = factory.createSectPr();
            sectPr.setPgSz(size);
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

    private static int pixelsToDxa(int pixels) {
        return  ( 1440 * pixels / getDPI() );
    }

    public static WordprocessingMLPackage cleanHeaderFooter(WordprocessingMLPackage word)  {
        MainDocumentPart mdp = word.getMainDocumentPart();
        SectPrFinder finder = new SectPrFinder(mdp);
        new TraversalUtil(mdp.getContent(), finder);
        for (SectPr sectPr : finder.getOrderedSectPrList()) {
            sectPr.getEGHdrFtrReferences().clear();
        }

        // Remove rels
        List<Relationship> hfRels = new ArrayList<>();
        for (Relationship rel : mdp.getRelationshipsPart().getRelationships().getRelationship() ) {

            if (rel.getType().equals(Namespaces.HEADER)
                    || rel.getType().equals(Namespaces.FOOTER)) {
                hfRels.add(rel);
            }
        }
        for (Relationship rel : hfRels ) {
            mdp.getRelationshipsPart().removeRelationship(rel);
        }
        return word;
    }

    public  static P newImage(WordprocessingMLPackage doc, File file) {

        java.io.InputStream is = null;
        try {
            is = new java.io.FileInputStream(file );
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        long length = file.length();
        if (length > Integer.MAX_VALUE) {
            System.out.println("File too large!!");
        }
        byte[] bytes = new byte[(int)length];
        int offset = 0;
        int numRead;
        try {
            while (offset < bytes.length
                    && (numRead=is.read(bytes, offset, bytes.length-offset)) >= 0) {
                offset += numRead;
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
        if (offset < bytes.length) {
            System.out.println("Could not completely read file "+file.getName());
        }
        try {
            is.close();
        } catch (IOException e) {
            e.printStackTrace();
        }

        String filenameHint = null;
        String altText = null;
        int id1 = 0;
        int id2 = 1;
        BinaryPartAbstractImage imagePart = null;
        try {
            imagePart = BinaryPartAbstractImage.createImagePart(doc, bytes);
        } catch (Exception e) {
            e.printStackTrace();
        }

        Inline inline = null;
        try {
            inline = imagePart.createImageInline( filenameHint, altText,
                    id1, id2, true);
        } catch (Exception e) {
            e.printStackTrace();
        }
        P  p = factory.createP();
        p.setPPr(factory.createPPr());
        R  run = factory.createR();
        RPr rPr = factory.createRPr();
        BooleanDefaultTrue value = new BooleanDefaultTrue();
        rPr.setNoProof(value );
        p.getContent().add(run);
        Drawing drawing = factory.createDrawing();
        run.getContent().add(drawing);
        drawing.getAnchorOrInline().add(inline);
        return p;

    }

    public static double ngrammPossibility (String actual, String checked) {
        if (actual.length() < 3 || checked.length() < 3)
            return (actual.equals(checked))?1:0;
        actual = actual.toLowerCase().trim().replaceAll(" +", "");
        checked = checked.toLowerCase().trim().replaceAll(" +", "");
        String[] actualGramm  = new String[actual.length() - 2];
        String[] checkedGramm  = new String[checked.length() - 2];
        int index = 0;
        for (int i = 0; i < checked.length(); i++) {
            if (i ==checked.length() - 3) {
                checkedGramm[index] = checked.substring(i, i+3);
                break;
            }
            else checkedGramm[index] = checked.substring(i, i+3);
            index++;
        }
        index = 0;
        for (int i = 0; i < actual.length(); i++) {
            if (i ==actual.length() - 3) {
                actualGramm[index] = actual.substring(i, i+3);
                break;
            }
            else actualGramm[index] = actual.substring(i, i+3);
            index++;
        }

        double coincidence = 0.0;
        index = 0;
//        int max = (checkedGramm.length>actualGramm.length)? actualGramm.length :checkedGramm.length;
        int max = actualGramm.length;
        while (index != max) {
            if (checked.contains(actualGramm[index]))
                coincidence++;
            ++index;
        }
        index = 0;
        int max2 = checkedGramm.length;
        while (index != max2) {
            if (actual.contains(checkedGramm[index]))
                coincidence++;
            ++index;
        }
        return coincidence/(max+max2);
    }

}
