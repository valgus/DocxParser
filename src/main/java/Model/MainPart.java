package Model;

import org.docx4j.jaxb.Context;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.Part;
import org.docx4j.openpackaging.parts.WordprocessingML.StyleDefinitionsPart;
import org.docx4j.openpackaging.parts.relationships.Namespaces;
import org.docx4j.wml.*;

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
        //set Pade margins and size
        DocxMethods.setPageMargins(document);
        //delete contents of header and footer
        document = DocxMethods.cleanHeaderFooter(document);
        if (!exist) throw new Exception();
        List<Title> titles = TemplateParser.getListOfTitles(template.getAbsolutePath());
        List<Object> documentParagraphes = DocxMethods.getAllElementFromObject(document.getMainDocumentPart(), P.class);
        Map<Integer, P> corespondences = findCorespondences(document, documentParagraphes, titles);
        if (corespondences.size() == 0)
            throw new Exception("Файл не соответствует шаблону!");
            int indexOfTableOfContent = -1;
            Collection<P> values = corespondences.values();
            Iterator it = values.iterator();
            P p;
            Title t;
            if (it.hasNext()) {
                do{
                    p = (P)it.next();
                  //  DocBase.setSpacing(p, 360);
                    if (indexOfTableOfContent==-1)
                        indexOfTableOfContent = DocxMethods.getIndexOfParagraph(document.getMainDocumentPart(), p);
                    t = findTitle(titles, p);
                    if (t!=null) {
                        DocBase.setText(p, t.getName(), true);
                        String[] atr = DocBase.getAttributes(t);
                   //     P previousP = DocBase.makePageBr();
                 //     document.getMainDocumentPart().getContent().add(i-1, previousP);
                        DocBase.setText(p, t.getName(), true);
                        String s = "LEFT";
                        String size = "28";
                        if (atr[0].equals("1")) {
                            DocBase.setText(p, DocBase.getText(p).toUpperCase(), true);
                            size = "32";
                           if (!DocBase.getText(p).toLowerCase().equals("приложения")) s = "CENTER";
                        }

                        DocBase.setStyle(p, size, "Times New Roman", null, atr[1], atr[2], 0, s, true);
                        DocBase.setHighlight(p, "green");
                    }
                    else {
                        DocBase.setStyle(p, null, null, null, null, null, 0, "LEFT", true);
                        DocBase.setHighlight(p, "green");
                    }

                } while (it.hasNext());
            }

        for (Object o : documentParagraphes) {   //setAttributes
            p = (P) o;
   //         DocBase.setSpacing(p, 240);
            if (!DocBase.getText(p).trim().equals("") & DocBase.getHighlight(p)==null) {
                DocBase.deleteTabsInParagraph(p);
                DocBase.addTab(p);
                DocBase.setStyle(p, "28","Times New Roman", null, null, null, 0, "LEFT", false);
            }
        }

        setEnumeration(documentParagraphes.subList(indexOfTableOfContent, documentParagraphes.size()-1));
        for (Object o : documentParagraphes) {
            p = (P) o;
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
        document.getMainDocumentPart().getContent().add(indexOfTableOfContent,  paragraphForTOC);
        document.getMainDocumentPart().getContent().add(indexOfTableOfContent+1, DocBase.makePageBr());

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

    private JAXBElement getWrappedFldChar(FldChar fldchar) {
        return new JAXBElement( new QName(Namespaces.NS_WORD12, "fldChar"),
                FldChar.class, fldchar);

    }


    private void setEnumeration(List<Object> documentParagraphes) {
        int currentNumId, currentLevel;
        LinkedHashMap<Integer, Integer> data = new LinkedHashMap<>();
        for (int i = 0; i < documentParagraphes.size(); i++) {
            P p = (P)documentParagraphes.get(i);
            currentLevel = DocBase.getLevelInList(p).intValue();
            currentNumId = DocBase.getNumIDInList(p).intValue();
            if (DocBase.getHighlight(p) == null && !DocBase.getText(p).isEmpty()){
                if (currentLevel!= -1 && currentNumId!=-1) {
                    if (data.size()==0) {
                        data.put(i, 1);
                    }
                    else {
                        Integer previous = data.get(i-1);
                        if (previous!= null) {
                            if (currentLevel == DocBase.getLevelInList((P)documentParagraphes.get(i-1)).intValue() &&
                                    currentNumId == DocBase.getNumIDInList((P)documentParagraphes.get(i-1)).intValue())
                                data.put(i, previous);
                            else if (currentLevel < DocBase.getLevelInList((P)documentParagraphes.get(i-1)).intValue()){
                                data.put(i, previous-1);
                            }
                            else {
                                data.put(i, previous+1);
                            }
                        }
                        else
                            data.put(i, 1);
                    }
                }
                else {
                    P[] around;
                    if (i !=0 && i!= documentParagraphes.size()-1) {
                        around = new P[2];
                        around[0] = (P)documentParagraphes.get(i-1);
                        around[1] = (P)documentParagraphes.get(i+1);
                    }
                    else {
                        around = new P[1];
                        around[0] = (i==0)? (P)documentParagraphes.get(i+1) :(P)documentParagraphes.get(i-1);
                    }
                    if (DocBase.isInList(p, around)) {
                        if (data.size()==0) {
                            data.put(i, 1);
                        }
                        else {
                            Integer previous = data.get(i-1);
                            if (previous!= null) {
                                if (DocBase.sameEnumeration((P)documentParagraphes.get(i-1), p))
                                    data.put(i, previous);
                                else {
                                        data.put(i, previous+1);
                                }
                            }
                            else {
                                data.put(i, 1);
                            }
                        }
                    }
                }
            }
        }
        String[] chars = {"а","б","в","д","е","ж","и","к","л","м","н","п","р","с","т","у","ф","х","ц","ч",
                "ш","щ","э","ю","я"};
        int num = 1;
        int charnum = 0;
        String character;
        Set<Integer> indexes = data.keySet();
        Iterator it = indexes.iterator();
        Integer previous = null;
        while (it.hasNext()) {
            int index = (int)it.next();
            P p = (P)documentParagraphes.get(index);
            DocBase.removeEnum(p);
            if (previous!= null) {
                if (data.get(previous) == data.get(index) && previous + 1 == index)
                    character = ";";
                else if (previous + 1 < index || data.get(previous) < data.get(index)){
                    character = ".";
                    num = 1;
                    charnum = 0;
                }
                    else
                        character = ":";
                String s = DocBase.getText((P)documentParagraphes.get(previous)).trim();
                if (s.endsWith(";")||s.endsWith(".")||s.endsWith(":")) {
                    s = s.substring(0, s.length()-1);
                    DocBase.setText((P)documentParagraphes.get(previous), s, true);
                }
                DocBase.setText((P)documentParagraphes.get(previous), character, false);
            }
            if (data.get(index) == 1) {
                DocBase.setText(p,"-  " + DocBase.getText(p), true);
            }
            else if (data.get(index) == 2) {
                DocBase.addTab(p);
                DocBase.setText(p, chars[charnum++] + ")  " +DocBase.getText(p), true);
            }
                else {
                DocBase.addTab(p);
                DocBase.addTab(p);
                DocBase.setText(p,String.valueOf(num++) + ")  " + DocBase.getText(p), true);
            }
            previous = index;
            if (!it.hasNext())
                DocBase.setText(p, ".", false);
        }
    }
}
