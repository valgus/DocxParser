import org.docx4j.jaxb.Context;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.*;

import java.io.File;
import java.io.FileNotFoundException;
import java.math.BigInteger;
import java.util.List;

public final class DocBase {


    public static void setStyle (P p, boolean bold, boolean italic, String align, String fontSize,
                                 boolean uppercase, String font) {

        List<Object> textes = DocxMethods.getAllElementFromObject(p, Text.class);
        if (textes.size()==1 & textes.get(0)=="") return;
        if (p.getPPr().getRPr()==null) {
            p.getPPr().setRPr(new ParaRPr());
        }
        for (Object text : textes)
        {
            if (uppercase) {

               ((Text) text).setValue(((Text) text).getValue().toUpperCase());
            }

            setBold(p.getPPr(), bold);



            setItalic(p.getPPr(), italic);


            if (fontSize != null && !fontSize.isEmpty()) {
                setSize(p.getPPr(), fontSize);
            }

            if (align != null || align != "") {
                setAlign(p.getPPr(), align.toLowerCase());
            }

            if (font != null || font != "") {
                setFont(p.getPPr(), font);
            }

        }
    }

    private static void setAlign (PPr paragraphProperties, String align) {
        JcEnumeration jcEnumeration = JcEnumeration.fromValue(align);
        Jc jc = new Jc();
        jc.setVal(jcEnumeration);
        paragraphProperties.setJc(jc);
    }

    private static void setSize(PPr runProperties, String fontSize) {
        HpsMeasure size = new HpsMeasure();
        size.setVal(new BigInteger(fontSize));
        runProperties.getRPr().setSz(size);
        runProperties.getRPr().setSzCs(size);
    }

    public static void setFont (PPr p, String font) {
        RFonts rf = new RFonts();
        rf.setAscii(font);
        rf.setCs(font);
        rf.setHAnsi(font);
        rf.setHint(STHint.DEFAULT);
        p.getRPr().setRFonts(rf);
//        List<Object> contents = p.;
//        for (Object content : contents) {
//            try{
//                R r = (R) content;
//                r.getRPr().setRFonts(rf);
//            }
//            catch (ClassCastException ex) {}
//        }
    }

    public static void setList (P p, int numID, int ilvl) {
        ObjectFactory objectFactory = Context.getWmlObjectFactory();
        PPr ppr = objectFactory.createPPr();

        PPrBase.NumPr numpr = objectFactory.createPPrBaseNumPr();
        ppr.setNumPr(numpr);


        PPrBase.NumPr.Ilvl wilvl = objectFactory.createPPrBaseNumPrIlvl();

        numpr.setIlvl(wilvl);

        wilvl.setVal(BigInteger.valueOf(ilvl));

        PPrBase.NumPr.NumId wnumID = objectFactory.createPPrBaseNumPrNumId();

        numpr.setNumId(wnumID);

        wnumID.setVal(BigInteger.valueOf(numID));
        PPrBase.PStyle pstyle = objectFactory.createPPrBasePStyle();
        ppr.setPStyle(pstyle);
        p.setPPr(ppr);
    }

    public static void setSpacing (P p, boolean isTitle) {
        if (p.getPPr()==null) p.setPPr(new PPr());
        PPrBase.Spacing spacing = new PPrBase.Spacing();
        spacing.setLine( new BigInteger((isTitle==true)?"480":"360"));
        //   spacing.setLineRule(STLineSpacingRule.AUTO);
        try {
            p.getPPr().setSpacing(spacing);
        }
        catch ( NullPointerException ex) {
            System.out.println("ups..");}
    }


    private static void setBold(PPr runProperties, boolean bold) {
        BooleanDefaultTrue b = new BooleanDefaultTrue();
        b.setVal(bold);
        runProperties.getRPr().setB(b);
    }

    private static void setItalic(PPr runProperties, boolean italic) {
        BooleanDefaultTrue i = new BooleanDefaultTrue();
        i.setVal(italic);
        runProperties.getRPr().setI(i);
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
        Highlight highlight = new Highlight();
        highlight.setVal(color);
        p.getPPr().getRPr().setHighlight(highlight);

        List<Object> contents = p.getContent();
        for (Object content : contents) {
            try{
            R r = (R) content;
            r.getRPr().setHighlight(highlight);
            }
            catch (ClassCastException ex) {}
        }
    }






    public static String getHighlight (P p) {
        try{
            return p.getPPr().getRPr().getHighlight().getVal();
        }
        catch (NullPointerException ex) {return null;}

    }


    public static String getSize(P p) {
        String i=null;
        List<Object> contents = p.getContent();
        for (Object content : contents) {
            if (content instanceof R) {
                try{
                i= ((R)content).getRPr().getSz().getVal().toString();
                    return i;
                }
                catch (NullPointerException ex) {i = null;}
            }
        }
        if (i == null) {
            try{
                i= p.getPPr().getRPr().getSzCs().getVal().toString();
                return i;
            }
            catch (NullPointerException ex) {i = null;}
        }
        return "28";
    }

    public static String getFont (P p) {

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

    public static String getAlign (P p) {
        List<Object> contents = p.getContent();
        for (Object content : contents) {
            if (content instanceof R) {
                try {
                    return ((P)((R)content).getParent()).getPPr().getJc().getVal().toString();
                }
                catch (NullPointerException ex) {return "left";}
            }
        }
        return "left";
    }

    private static boolean isUpperCase(P p) {
        return DocBase.getText(p)==DocBase.getText(p).toUpperCase();
    }

    private static boolean isItalic(P p) {
        List<Object> contents = p.getContent();
        for (Object content : contents) {
            if (content instanceof R) {
                try {
                    return ((R)content).getRPr().getI().isVal();
                }
                catch (NullPointerException ex) {return false;}
            }
        }
        return false;
    }

    private static boolean isBold(P p) {
        List<Object> contents = p.getContent();
        for (Object content : contents) {
            if (content instanceof R) {
                try {
                    return ((R)content).getRPr().getB().isVal();
                }
                catch (NullPointerException ex) {return false;}
            }
        }
        return false;
    }

    public static  String getAttributes (P p) {

        StringBuffer attributes = new StringBuffer();
        attributes.append(getSize(p));//0
        attributes.append(";");
        attributes.append(getFont(p));//1
        attributes.append(";");
        attributes.append(getAlign(p));//2
        attributes.append(";");
        attributes.append(isBold(p));//3
        attributes.append(";");
        attributes.append(isItalic(p));//4
        attributes.append(";");
        attributes.append(isUpperCase(p));//5
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
            List<Object> jaxbNodes = DocxMethods.createJaxbNodes(word);
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
