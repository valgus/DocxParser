package Model;

import Model.DocBase;
import Model.DocxMethods;
import Model.TemplateParser;
import Model.Title;
import org.docx4j.jaxb.Context;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.Part;
import org.docx4j.openpackaging.parts.WordprocessingML.StyleDefinitionsPart;
import org.docx4j.openpackaging.parts.relationships.Namespaces;
import org.docx4j.wml.*;

import javax.print.Doc;
import javax.xml.bind.JAXBElement;
import javax.xml.namespace.QName;
import java.io.File;
import java.math.BigInteger;
import java.util.*;

public class MainPart {

    File template, compared;
    boolean exist = false;
    ObjectFactory factory = Context.getWmlObjectFactory();


    public void setTwoDocx(String template, String compared) {
        this.template = new File(template);
        this.compared = new File(compared);
        exist = true;
    }

    public WordprocessingMLPackage setAppropriateText() throws Exception {
        WordprocessingMLPackage document  = DocxMethods.getTemplate(compared.getAbsolutePath());
        if (!exist) throw new Exception();
        Part stylesPart = new org.docx4j.openpackaging.parts.WordprocessingML.StyleDefinitionsPart();
        ((org.docx4j.openpackaging.parts.WordprocessingML.StyleDefinitionsPart) stylesPart)
                    .unmarshalDefaultStyles();
       // document.addTargetPart(stylesPart);

//        List<Object> contents = document.getMainDocumentPart().getContent().subList(
//                0, 43);
//        document.getMainDocumentPart().getContent().removeAll(contents);
        List<Title> titles = TemplateParser.getListOfTitles(template.getAbsolutePath());
        List<Object> documentParagraphes = DocxMethods.getAllElementFromObject(document.getMainDocumentPart(), P.class);
        Map<Integer, P> corespondences = findCorespondences(document, documentParagraphes, titles);

        if (corespondences.size() == 0) {
            throw new Exception("Файл не соответствует шаблону!");
        }
        else {
            Collection<P> values = corespondences.values();
            Iterator it = values.iterator();
            P p;
            Title t;
            if (it.hasNext()) {
                do{
                    p = (P)it.next();
                  //  DocBase.setSpacing(p, 360);
                    t = findTitle(titles, p);
                    if (t!=null) {
                        DocBase.setText(p, t.getName(), true);
                        String[] atr = DocBase.getAttributes(t);
                        DocBase.setStyle(p, null, null, null, atr[1], atr[2], 0, "CENTER", true);
                        DocBase.setHighlight(p, "green");
                    }
                    else {
                        DocBase.setStyle(p, null, null, null, null, null, 0, "CENTER", true);
                        DocBase.setHighlight(p, "green");
                    }

                } while (it.hasNext());
            }
        }
        for (Object o : documentParagraphes) {   //setAttributes
            P p = (P) o;
   //         DocBase.setSpacing(p, 240);
            if (!DocBase.getText(p).trim().equals("") & DocBase.getHighlight(p)==null)
                DocBase.setStyle(p, "28","Times New Roman", null, null, null, 0, "BOTH", false);
        }

//        changeEnumeration(document, documentParagraphes);
//
        for (Object o : documentParagraphes) {
            P p = (P) o;
            if (DocBase.getHighlight(p)!=null) {
                DocBase.setHighlight(p, null);
            }
        }

        //set table of contents
        P paragraphForTOC = factory.createP();
        R r = factory.createR();

        FldChar fldchar = factory.createFldChar();
        fldchar.setFldCharType(STFldCharType.BEGIN);
        fldchar.setDirty(true);
        r.getContent().add(getWrappedFldChar(fldchar));
        paragraphForTOC.getContent().add(r);

        R r1 = factory.createR();
        Text txt = new Text();
        txt.setSpace("preserve");
        txt.setValue("TOC \\o \"1-3\" \\h \\z \\u \\h");
        r.getContent().add(factory.createRInstrText(txt) );
        paragraphForTOC.getContent().add(r1);

        FldChar fldcharend = factory.createFldChar();
        fldcharend.setFldCharType(STFldCharType.END);
        R r2 = factory.createR();
        r2.getContent().add(getWrappedFldChar(fldcharend));
        paragraphForTOC.getContent().add(r2);
        document.getMainDocumentPart().getContent().add(0,  paragraphForTOC);

        return document;

    }

    private Map<Integer, P> findCorespondences(WordprocessingMLPackage wordprocessingMLPackage,List<Object> document, List<Title> titles) {
        Map<Integer, P> check = new TreeMap<>();
        for (int i = 0; i< titles.size(); i++) {
            P p = findP(document, titles.get(i).getName());
            if (p != null) {
                check.put(DocxMethods.getIndexOfParagraph(wordprocessingMLPackage.getMainDocumentPart(), p), p);
            }
        }
        return check;
    }

    private P findP (List<Object> paragraphes, String s) throws ClassCastException{
        int i = 0;
        P paragraph = null;
        while ( paragraphes.size()> i) {
            String pTemplate = DocBase.getText((P) paragraphes.get(i)).toLowerCase();
            if (s.toLowerCase().equals(pTemplate.toLowerCase())) {
                paragraph = (P) paragraphes.get(i);
                break;
            }
            i++;
        }

        return paragraph;
    }

    private Title findTitle (List<Title> t, P p) {
        String name = DocBase.getText(p).toLowerCase();
        for (Title title : t) {
            if (title.getName().toLowerCase().equals(name)) {
                return title;
            }
        }
        return  null;
    }

    private void changeEnumeration (WordprocessingMLPackage document,List<Object> documentParagraphes)
    {
        ArrayList<Integer> indexes = new ArrayList<>();
        int number = 0;
        BigInteger not  = new BigInteger("-1");
        for (Object o : documentParagraphes) {
            P p = (P) o;
            DocBase.setSpacing(p, 240);
            if (!DocBase.getText(p).trim().equals("") & DocBase.getHighlight(p)==null) {
                if (!DocBase.getLevelInList(p).equals(not) ) {
                    int i = DocxMethods.getIndexOfParagraph(document.getMainDocumentPart(), p);
                    if (indexes.size()== 0)
                        number++;
                    if (indexes.size()!= 0 && (indexes.get(indexes.size()-1) +1 != i))
                        number++;
                    indexes.add(i);
                }
            }
        }
        String id = "22";

        int previous = indexes.get(0)-1;
        for (Integer index : indexes) {
            P p = DocxMethods.getParagraphFromIndex(document.getMainDocumentPart(), index);
            if (previous != index -1) {
                id = String.valueOf(Integer.decode(id)+1);
            }
            DocBase.setList(p);
            previous = index;
        }
    }

    private JAXBElement getWrappedFldChar(FldChar fldchar) {
        return new JAXBElement( new QName(Namespaces.NS_WORD12, "fldChar"),
                FldChar.class, fldchar);

    }
}
