import org.docx4j.Docx4J;
import org.docx4j.convert.out.FOSettings;
import org.docx4j.fonts.IdentityPlusMapper;
import org.docx4j.fonts.Mapper;
import org.docx4j.fonts.PhysicalFont;
import org.docx4j.fonts.PhysicalFonts;
import org.docx4j.jaxb.XPathBinderAssociationIsPartialException;
import org.docx4j.model.fields.FieldUpdater;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.ContentAccessor;
import org.docx4j.wml.P;

import javax.xml.bind.JAXBElement;
import javax.xml.bind.JAXBException;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.OutputStream;
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

    public static void convertToPDF (String way, WordprocessingMLPackage wordMLPackage) throws Exception {


        // Font regex (optional)
        // Set regex if you want to restrict to some defined subset of fonts
        // Here we have to do this before calling createContent,
        // since that discovers fonts
        String regex = null;
        // Windows:
        // String
        // regex=".*(calibri|camb|cour|arial|symb|times|Times|zapf).*";
        regex=".*(calibri|camb|cour|arial|times|comic|georgia|impact|LSANS|pala|tahoma|trebuc|verdana|symbol|webdings|wingding).*";
        // Mac
        // String
        // regex=".*(Courier New|Arial|Times New Roman|Comic Sans|Georgia|Impact|Lucida Console|Lucida Sans Unicode|Palatino Linotype|Tahoma|Trebuchet|Verdana|Symbol|Webdings|Wingdings|MS Sans Serif|MS Serif).*";
        PhysicalFonts.setRegex(regex);


        if (way==null)
            System.out.println("No imput path passed, creating dummy document");



        // Refresh the values of DOCPROPERTY fields
        FieldUpdater updater = new FieldUpdater(wordMLPackage);
        updater.update(true);

        // Set up font mapper (optional)
        Mapper fontMapper = new IdentityPlusMapper();
        wordMLPackage.setFontMapper(fontMapper);

        // .. example of mapping font Times New Roman which doesn't have certain Arabic glyphs
        // eg Glyph "ي" (0x64a, afii57450) not available in font "TimesNewRomanPS-ItalicMT".
        // eg Glyph "ج" (0x62c, afii57420) not available in font "TimesNewRomanPS-ItalicMT".
        // to a font which does
        PhysicalFont font
                = PhysicalFonts.get("Arial Unicode MS");
        // make sure this is in your regex (if any)!!!
        if (font!=null) {
            fontMapper.put("Times New Roman", font);
            fontMapper.put("Arial", font);
        }
        fontMapper.put("Libian SC Regular", PhysicalFonts.get("SimSun"));

        // FO exporter setup (required)
        // .. the FOSettings object
        FOSettings foSettings = Docx4J.createFOSettings();
        foSettings.setFoDumpFile(new java.io.File(way + ".fo"));

        foSettings.setWmlPackage(wordMLPackage);

        // Document format:
        // The default implementation of the FORenderer that uses Apache Fop will output
        // a PDF document if nothing is passed via
        // foSettings.setApacheFopMime(apacheFopMime)
        // apacheFopMime can be any of the output formats defined in org.apache.fop.apps.MimeConstants eg org.apache.fop.apps.MimeConstants.MIME_FOP_IF or
        // FOSettings.INTERNAL_FO_MIME if you want the fo document as the result.
        //foSettings.setApacheFopMime(FOSettings.INTERNAL_FO_MIME);

        // exporter writes to an OutputStream.
        String outputfilepath = "new.pdf";
        OutputStream os = new java.io.FileOutputStream(outputfilepath);

        // Specify whether PDF export uses XSLT or not to create the FO
        // (XSLT takes longer, but is more complete).

        // Don't care what type of exporter you use
        Docx4J.toFO(foSettings, os, Docx4J.FLAG_EXPORT_PREFER_XSL);

        // Prefer the exporter, that uses a xsl transformation
        // Docx4J.toFO(foSettings, os, Docx4J.FLAG_EXPORT_PREFER_XSL);

        // Prefer the exporter, that doesn't use a xsl transformation (= uses a visitor)
        // .. faster, but not yet at feature parity
        // Docx4J.toFO(foSettings, os, Docx4J.FLAG_EXPORT_PREFER_NONXSL);

        System.out.println("Saved: " + outputfilepath);
    }

    public static void main(String[] args) {
        try {
            convertToPDF("3.docx", DocxMethods.getTemplate("3.docx"));
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

}
