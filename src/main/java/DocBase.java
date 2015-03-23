import org.docx4j.jaxb.Context;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.*;
import org.plutext.jaxb.xslfo.TextTransformType;

import javax.xml.bind.JAXBElement;
import java.io.File;
import java.io.FileNotFoundException;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.List;

public final class DocBase {


    public static void setStyle (P p, String fontSize, String font, String style, String level, String id) {

        String text = getText(p);
        if (text == null || text == "") return;
        p.getPPr().setRPr(new ParaRPr());

        if (style != null && !style.isEmpty()) {
            setStyle(p, style);
        }

        if (fontSize != null && !fontSize.isEmpty()) {
            setSize(p, fontSize);
        }

        if (font != null && font != "") {
            setFont(p, font);
        }
        if (level != null && level != "" && id!= null && id != "") {
            setLevel(p, level, id);
        }


    }

    private static void setSize(P p, String fontSize) {
        HpsMeasure size = new HpsMeasure();
        size.setVal(new BigInteger(fontSize));
        List<Object> contents = p.getContent();
        for (Object content : contents) {
            if (content instanceof PPr) {
               ((PPr)content).getRPr().setSz(size);
               ((PPr)content).getRPr().setSzCs(size);
            }
            if (content instanceof R) {
                if (((R)content).getRPr() == null) {((R)content).setRPr(new RPr());}
                ((R)content).getRPr().setSz(size);
                ((R)content).getRPr().setSzCs(size);
            }
        }
        try {
            p.getPPr().getRPr().setSz(size);
            p.getPPr().getRPr().setSzCs(size);
        }
        catch (NullPointerException ex) {}
    }

    private static void setStyle(P p, String number) {
        if (number == null) return;

        ObjectFactory factory = Context.getWmlObjectFactory();
        String text = getText(p);
        p.getContent().clear();
        Text t = factory.createText();
        t.setValue(text);
        R run = factory.createR();
        run.getContent().add(t);
        run.setRPr(factory.createRPr());
        p.getContent().add(run);
        PPr pPr = factory.createPPr();
        ParaRPr rpr = factory.createParaRPr();
        pPr.setRPr(rpr);
        p.setPPr(pPr);

        PPrBase.PStyle style = new PPrBase.PStyle();
        style.setVal(number);

        p.getPPr().setPStyle(style);
    }

    public static void setText (P p, String text) {
        if (text == null || text.equals("")) return;
        List<Object> contents = p.getContent();
        R first = null;
        Text t = new Text();
        t.setValue(text);
        for (int i = 0; i< contents.size(); i++) {
            if (contents.get(i) instanceof R) {
                if (first != null)
                    contents.remove(i);
                else
                    first = (R)contents.get(i);
            }
        }
        if (first!=null) {

            for (Object o : first.getContent()) {
                if (o instanceof JAXBElement) {
                try{
                    if (((JAXBElement) o).getValue() instanceof Text)
                      ((JAXBElement) o).setValue(t);

                }
                catch (ClassCastException ex) {}
                }
            }
        }
        else {
            first =new R();
            first.getContent().add(t);
            p.getContent().add(first);
        }
    }

    private static void setFont (P p, String font) {
        RFonts rf = new RFonts();
        rf.setAscii(font);
        rf.setCs(font);
        rf.setHAnsi(font);
        rf.setHint(STHint.DEFAULT);
        try {
            p.getPPr().getRPr().setRFonts(rf);
        }
        catch (NullPointerException ex) {
            p.setPPr(new PPr());
            p.getPPr().getRPr().setRFonts(rf);
        }
        List<Object> contents = p.getContent();
        for (Object content : contents) {
            if (content instanceof R) {
                R r = (R)content;
                if (r.getRPr() != null)
                    r.getRPr().setRFonts(rf);
                else {
                    r.setRPr(new RPr());
                    r.getRPr().setRFonts(rf);
                }
            }
        }

    }

    public static void setSpacing (P p, boolean isTitle) {
        if (p.getPPr()==null) p.setPPr(new PPr());
        PPrBase.Spacing spacing = new PPrBase.Spacing();
        spacing.setLine( new BigInteger((isTitle==true)?"480":"360"));
        try {
            p.getPPr().setSpacing(spacing);
        }
        catch ( NullPointerException ex) {
            System.out.println("ups..");}
    }

    public static P addParagraph(MainDocumentPart mdp, String simpleText, int index) {
        ObjectFactory factory = Context.getWmlObjectFactory();
        P para = factory.createP();
        if (simpleText !=null) {

            Text t = factory.createText();
            t.setValue(simpleText);
            R run = factory.createR();
            run.getContent().add(t);
            para.getContent().add(run);
            PPr pPr = factory.createPPr();
            para.setPPr(pPr);


        }
        mdp.getContent().add(index, para);
        DocBase.setSpacing(para, false);
        return para;
    }

    public static void setHighlight (P p, String color) {
        List<Object> contents = p.getContent();

        if (color == null) {
            if (p.getPPr() != null && p.getPPr().getRPr() != null)
                p.getPPr().getRPr().setHighlight(null);
            for (Object content : contents) {
                try{
                    R r = (R) content;
                    if (r.getRPr() != null)
                      r.getRPr().setHighlight(null);
            }
                catch (ClassCastException ex) {}
            return;

           }
        }
        Highlight highlight = new Highlight();
        highlight.setVal(color);
        p.getPPr().getRPr().setHighlight(highlight);
        for (Object content : contents) {
            try{
            R r = (R) content;
            r.getRPr().setHighlight(highlight);
            }
            catch (ClassCastException ex) {}
        }

    }

    private static void setLevel (P p, String level, String id ) {
        if (level == null || level.equals("-1")) return;
        if (id == null || id.equals("-1")) return;
        PPrBase.NumPr numPr = new PPrBase.NumPr();
        PPrBase.NumPr.Ilvl ilvl = new PPrBase.NumPr.Ilvl();
        ilvl.setVal(new BigInteger(level));
        numPr.setIlvl(ilvl);
        PPrBase.NumPr.NumId numID = new PPrBase.NumPr.NumId();
        numID.setVal(new BigInteger(id));
        numPr.setNumId(numID);
        try {
            p.getPPr().setNumPr(numPr);
        }
        catch (NullPointerException ex) {
            p.setPPr(new PPr());
            p.getPPr().setNumPr(numPr);
        }
    }

    public static void setList (P p) {
        ObjectFactory wmlObjectFactory = new org.docx4j.wml.ObjectFactory();
        PPrBase.NumPr pprbasenumpr = wmlObjectFactory.createPPrBaseNumPr();
        p.getPPr().setNumPr(pprbasenumpr);
        PPrBase.NumPr.Ilvl pprbasenumprilvl = wmlObjectFactory.createPPrBaseNumPrIlvl();
        pprbasenumpr.setIlvl(pprbasenumprilvl);
        pprbasenumprilvl.setVal( BigInteger.valueOf( 1) );
        PPrBase.NumPr.NumId pprbasenumprnumid = wmlObjectFactory.createPPrBaseNumPrNumId();
        pprbasenumpr.setNumId(pprbasenumprnumid);
        pprbasenumprnumid.setVal( BigInteger.valueOf( 22) );
        PPrBase.PStyle pprbasepstyle = wmlObjectFactory.createPPrBasePStyle();
        p.getPPr().setPStyle(pprbasepstyle);
        pprbasepstyle.setVal( "ListParagraph");
    }

    private static String getStyle(P p) {
        try{
            return p.getPPr().getPStyle().getVal();
        }
        catch (NullPointerException ex) {return null;}
    }


    public static String getHighlight (P p) {
        try{
            return p.getPPr().getRPr().getHighlight().getVal();
        }
        catch (NullPointerException ex) {return null;}

    }

    private static String getFont (P p) {

        List<Object> contents = p.getContent();
        for (Object content : contents) {
            if (content instanceof R) {
                try{
                return ((R)content).getRPr().getRFonts().getAscii().toString();
                }
                catch(NullPointerException ex) {return "Times New Roman";}
            }
        }
        return "Times New Roman";
    }

    public static BigInteger getLevelInList (P p) {
        try{
            BigInteger level = p.getPPr().getNumPr().getIlvl().getVal();
            return  level;
        }
        catch(NullPointerException ex) {return new BigInteger("-1");}
    }

    public static BigInteger getNumIDInList(P p) {
        try{
            BigInteger numID = p.getPPr().getNumPr().getNumId().getVal();
            return  numID;
        }
        catch(NullPointerException ex) {return new BigInteger("-1");}
    }

    public static String getText (P p) {
        List<Object> words = DocxMethods.getAllElementFromObject(p, Text.class);
        StringBuffer name = new StringBuffer("");
        for (Object word : words) {
            name.append(((Text)word).getValue());
        }
        return name.toString();
    }

    private static boolean isUpperCase(P p) {
        return DocBase.getText(p)==DocBase.getText(p).toUpperCase();
    }

    public static  String getAttributes (P p) {
        StringBuffer attributes = new StringBuffer();
        attributes.append(getStyle(p));//0
        attributes.append(';');
        attributes.append(getLevelInList(p));//1
        attributes.append(';');
        attributes.append(getNumIDInList(p));//2
        attributes.append(';');
        return attributes.toString();
    }

    public static String[] getAttributes(Title title) {
        String attributes = title.getAttributes();
        String splittedString[] = attributes.split(";");
        return splittedString;
    }


    public static void main(String[] args) {
        WordprocessingMLPackage word;
        try {
            word = DocxMethods.getTemplate("docx/document.docx");
            addParagraph(word.getMainDocumentPart(), "123 run me", 3);
            List<Object> jaxbNodes = DocxMethods.createParagraphJAXBNodes(word);
            for (Object jaxbNode : jaxbNodes) {
                P p = (P) jaxbNode;
               // setFont(p, "Arial");
                setSpacing(p, true);
                System.out.println(getText(p)+getFont(p));

                System.out.println(isUpperCase(p));
            }
            word.save(new File("docx/2.docx"));
        } catch (Docx4JException e) {
            e.printStackTrace();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }

    }
}
