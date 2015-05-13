package Model;

import org.docx4j.XmlUtils;
import org.docx4j.dml.wordprocessingDrawing.Anchor;
import org.docx4j.jaxb.Context;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.openpackaging.parts.relationships.Namespaces;
import org.docx4j.wml.*;

import javax.xml.bind.JAXBElement;
import javax.xml.bind.JAXBException;
import javax.xml.namespace.QName;
import java.io.File;
import java.math.BigInteger;
import java.util.*;
public class MainPart {

    Document doc;
    public MainPart(Document doc){
        this.doc = doc;
    }

    File template, compared;
    ObjectFactory factory = Context.getWmlObjectFactory();
    ProcessDocument process;


    public void setTwoDocx(String template, File compared, ProcessDocument process) {
        this.template = new File(template);
        this.compared = compared;
        this.process = process;
    }

    public boolean sendIndfo(String s) {
        return process.sendInfo(s);
    }

    public WordprocessingMLPackage setAppropriateText(WordprocessingMLPackage document) {
        //set Pade margins and size
        DocxMethods.setPageMargins(document);
        //delete contents of header and footer
        List<Title> titles = TemplateParser.getListOfTitles(template.getAbsolutePath());
        List<Object> documentParagraphes = DocxMethods.getAllElementFromObject(document.getMainDocumentPart(), P.class);
        int from = 0;
        for (int i =0; i < documentParagraphes.size(); i++) {
            if (DocBase.getText((P)documentParagraphes.get(i)).matches(" *[«“\"]?[ПЭТОАБИ]{1}[12]?[»”\"]? *"))
                from = i;
            if (DocBase.getText((P)documentParagraphes.get(i)).matches(" *[0-9]{4} *") ||
                    DocBase.getText((P)documentParagraphes.get(i)).equals("ДАТА"))
                from = i;
        }
        documentParagraphes = documentParagraphes.subList(from, documentParagraphes.size()-1);
        TreeMap<Integer, P> corespondences = new TreeMap<>();
        boolean hasAnnotation = false;
        boolean empty = true;
        for (int index = 0; index < documentParagraphes.size(); index++) {
            P p = (P)documentParagraphes.get(index);
            List<Object> contents = DocxMethods.getAllElementFromObject(p, FldChar.class);
            List<Object> textOCntents = DocxMethods.getAllElementFromObject(p, Text.class);
            if (contents.size() >= 1 ) {
                boolean remove = true;
                for (Object text : textOCntents  ) {
                    if (((Text)text).getValue().contains("Рисунок") || ((Text)text).getValue().contains("Рис") ||
                            ((Text)text).getValue().contains("Иллюстрация")) {
                        remove = false;
                        break;
                    }
                }
                if (remove)document.getMainDocumentPart().getContent().remove(p);
            }
            if (!DocBase.getText(p).equals("")) {
                empty = false;
            }
            else {
                List<Object> drawingContent = DocxMethods.getAllElementFromObject(p, Drawing.class);
                if (drawingContent.size()==0)
                    document.getMainDocumentPart().getContent().remove(p);
            }
        }
        if (empty) {
            sendIndfo("Docx is empty");
            return null;
        }
        corespondences.putAll(findCorespondences(document, documentParagraphes, titles));
        if (corespondences.size() < titles.size()/2) {
            sendIndfo("Файл не соответствует шаблону!");
            return null;
        }
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
        //        document.getMainDocumentPart().getContent().add(indexOfTableOfContent, DocBase.makePageBr());
 //       document.getMainDocumentPart().getContent().add(corespondences.firstKey(), DocBase.makePageBr());
        document.getMainDocumentPart().getContent().add(corespondences.firstKey(),  paragraphForTOC);
        P content = factory.createP();
        DocBase.setText(content, "СОДЕРЖАНИЕ", false);
        DocBase.setStyle(content, "28", "Times New Roman", null, -1, -1, 0, "LEFT", true);
        document.getMainDocumentPart().getContent().add(corespondences.firstKey(), content);
        for (int index = 0; index < documentParagraphes.size(); index++) {
            P p = (P)documentParagraphes.get(index);
            String s = DocBase.getText(p);
            if (s.toLowerCase().equals("аннотация") || DocxMethods.ngrammPossibility("аннотация",s) > 0.8) {
                corespondences.put(index, p);
                hasAnnotation = true;
            }
            if (s.toLowerCase().equals("список сокращений") ||
                    DocxMethods.ngrammPossibility("список сокращений",s) > 0.8)
                corespondences.put(index, p);
            if (s.toLowerCase().contains("приложени") && s.length() < 15) {
                corespondences.put(index, p);
            }
        }
        int indexOfTableOfContent = -1;
        Collection<P> values = corespondences.values();
        Iterator it = values.iterator();
        P p;Title t;
        if (it.hasNext()) {
            do{
                p = (P)it.next();

                //       List<Object> content = p.getContent();
//                for (int i =0 ; i < content.size(); i++) {
//                    if (!(content.get(i) instanceof R))
//                        content.remove(content.get(i));
//                }
                //  DocBase.setSpacing(p, 360);
                if (indexOfTableOfContent==-1) {
                    document.getMainDocumentPart().getContent().add(DocxMethods.getIndexOfParagraph(document.getMainDocumentPart(),p),
                            DocBase.makePageBr());
                    if (DocxMethods.ngrammPossibility("аннотация",DocBase.getText(p)) < 0.4)
                        indexOfTableOfContent = DocxMethods.getIndexOfParagraph(document.getMainDocumentPart(), p);
                    if (doc.isAnnotation() & !hasAnnotation) {
                        P annotation = factory.createP();
                        DocBase.setText(annotation, "АННОТАЦИЯ", false);
                        DocBase.setStyle(annotation, "32", "Times New Roman", null, -1, -1, 0, "LEFT", true);
                        P explaination = factory.createP();
                        DocBase.setText(explaination, "Необходимо добавить раздел \"Аннотация\".", false);
                        document.getMainDocumentPart().getContent().add(corespondences.firstKey()-1, explaination);
                        document.getMainDocumentPart().getContent().add(corespondences.firstKey()-1, annotation);

                    //    document.getMainDocumentPart().getContent().add(indexOfTableOfContent, DocBase.makePageBr());
                    }
                }
                t = findTitle(titles, p);
                if (t!=null) {
                    //      DocBase.setText(p, t.getName(), true);
                    String[] atr = DocBase.getAttributes(t);
                    DocBase.setText(p, t.getName(), true);
                    String s = "LEFT";
                    String size = "28";
                    if (atr[0].equals("1")) {
                        DocBase.setText(p, DocBase.getText(p).toUpperCase(), true);
                        size = "32";
                        if (!DocBase.getText(p).toLowerCase().contains("приложен"))
                            DocBase.setStyle(p, size, "Times New Roman", null, 0, 1, 0, s, true);
                        else
                            DocBase.setStyle(p, size, "Times New Roman", null, -1, -1, 0, "RIGHT", true);
                    }
                    else {
                        DocBase.setStyle(p, size, "Times New Roman", null, 1, 1, 0, s, true);
                    }
                    DocBase.setHighlight(p, "green");
                }
//                else {
//                    DocBase.setStyle(p, "32", "Times New Roman", null, -1, -1, 0, "LEFT", true);
//                    DocBase.setHighlight(p, "green");
//                }


            } while (it.hasNext());
        }

        for (Object o : documentParagraphes) {   //setAttributes
            p = (P) o;
            if (!DocBase.getText(p).trim().equals("") & DocBase.getHighlight(p)==null) {
                DocBase.deleteTabsInParagraph(p);
                DocBase.addTab(p);
                DocBase.setStyle(p, "24","Times New Roman", null, -1, -1, 120, "LEFT", false);
            }
        }
        if (hasAnnotation)
            indexOfTableOfContent+=3;
        setEnumeration(documentParagraphes.subList(indexOfTableOfContent, documentParagraphes.size()-1));
        int numEmpty = 0;
        for (int i = 0; i < documentParagraphes.size(); i++) {
            p = (P) documentParagraphes.get(i);
            if (DocBase.getHighlight(p)!=null) {
                DocBase.setHighlight(p, null);
            }
            if (DocBase.getText(p).trim().equals("") ||DocBase.getText(p).trim().isEmpty()) {
                List<Object> drawingContent = DocxMethods.getAllElementFromObject(p, Drawing.class);
                if (drawingContent.size()==0) {
                    document.getMainDocumentPart().getContent().remove(p);
                if (indexOfTableOfContent > i)
                    numEmpty ++;
                }
            }
        }
  //      processImage(document.getMainDocumentPart());
        //set table of contents
        indexOfTableOfContent-=numEmpty;



        boolean findRegistrazuia_izmenenii = false;
        boolean findSoglasovano = false;
        boolean findSostavili = false;
        for (int i = documentParagraphes.size()-1; i > documentParagraphes.size()/2; i--) {
            P para  = (P) documentParagraphes.get(i);
            if (DocxMethods.ngrammPossibility(DocBase.getText(para).toLowerCase(), "согласовано") > 0.6)
                findSoglasovano = true;
            if (DocxMethods.ngrammPossibility(DocBase.getText(para).toLowerCase(), "составили") > 0.6)
                findSostavili = true;
            if (DocxMethods.ngrammPossibility(DocBase.getText(para).toLowerCase(), "лист регистрации изменений") > 0.6)
                findRegistrazuia_izmenenii = true;
        }
        if (!findSostavili && sendIndfo("добавить составили")) {
            try {
                document.getMainDocumentPart().addObject(DocBase.makePageBr());
                P para = factory.createP();
                DocBase.setText(para, "СОСТАВИЛИ", false);
                DocBase.setStyle(para, "28", "Times New Roman", null, -1, -1, 0, "CENTER", true);
                document.getMainDocumentPart().addObject(para);
                Tbl tbl = (Tbl)XmlUtils.unmarshalString(Models.sostaviliTbl);
                document.getMainDocumentPart().addObject(tbl);
            } catch (JAXBException e) {
                e.printStackTrace();
            }
        }
        if (!findSoglasovano & sendIndfo("добавить согласовано")) {
            try {
                P para = factory.createP();
                DocBase.setText(para, "СОГЛАСОВАНО", false);
                DocBase.setStyle(para, "28", "Times New Roman", null, -1, -1, 0, "CENTER", true);
                document.getMainDocumentPart().addObject(para);
                Tbl tbl = (Tbl)XmlUtils.unmarshalString(Models.soglasovanoTbl);
                document.getMainDocumentPart().addObject(tbl);
                document.getMainDocumentPart().addObject(DocBase.makePageBr());
            } catch (JAXBException e) {
                e.printStackTrace();
            }
        }
        if (!findRegistrazuia_izmenenii) {
            try {
                P para = factory.createP();
                DocBase.setText(para, "ЛИСТ РЕГИСТРАЦИИ ИЗМЕНЕНИЙ", false);
                DocBase.setStyle(para, "28", "Times New Roman", null, -1, -1, 0, "CENTER", true);
                document.getMainDocumentPart().addObject(para);
                Tbl tbl = (Tbl)XmlUtils.unmarshalString(Models.registrIzm+Models.s2+Models.s3+Models.s4+Models.s5+Models.s6);
                document.getMainDocumentPart().addObject(tbl);
            } catch (JAXBException e) {
                e.printStackTrace();
            }
        }

        return document;

    }

    private TreeMap<Integer, P> findCorespondences(WordprocessingMLPackage wordprocessingMLPackage,List<Object> document, List<Title> titles) {
        TreeMap<Integer, P> check = new TreeMap<>();
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
            if (s.toLowerCase().equals(pTemplate.toLowerCase()) ||
                    DocxMethods.ngrammPossibility(s.toLowerCase(), pTemplate.toLowerCase()) > 0.8) {
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
            if (title.getName().toLowerCase().equals(name)
                    || DocxMethods.ngrammPossibility(title.getName().toLowerCase(), name.toLowerCase()) >= 0.5) {
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
                        if (i!= 0) {
                            String s = DocBase.getText((P)documentParagraphes.get(i-1));
                            if (!s.endsWith(":") && (s.endsWith(".") || s.endsWith(";") || s.endsWith("!"))) {
                                DocBase.setText((P)documentParagraphes.get(i-1), s.substring(0, s.length()-1) + ":", true );
                            }
                        }

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
                        if (i!= 0) {
                            String s = DocBase.getText((P)documentParagraphes.get(i-1));
                            if (!s.endsWith(":") && (s.endsWith(".") || s.endsWith(";") || s.endsWith("!"))) {
                                DocBase.setText((P)documentParagraphes.get(i-1), s.substring(0, s.length()-1) + ":", true );
                            }
                        }
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
                            if (i!= 0) {
                                String s = DocBase.getText((P)documentParagraphes.get(i-1));
                                if (!s.endsWith(":") && (s.endsWith(".") || s.endsWith(";") || s.endsWith("!"))) {
                                    DocBase.setText((P)documentParagraphes.get(i-1), s.substring(0, s.length()-1) + ":", true );
                                }
                            }
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
                                String s = DocBase.getText((P)documentParagraphes.get(i-1));
                                if (!s.endsWith(":") && (s.endsWith(".") || s.endsWith(";") || s.endsWith("!"))) {
                                    DocBase.setText(p, s.substring(0, s.length()-1) + ":", true );
                                }
                            }
                        }
                    }
                }
            }
        }

        String character;
        Set<Integer> indexes = data.keySet();
        Iterator it = indexes.iterator();
        Integer previous = null;
        Integer next = null;
        while (it.hasNext()) {
            int index = (int)it.next();
            if (index!= indexes.size()-1)
                next = index+1;
            P p = (P)documentParagraphes.get(index);
            DocBase.removeEnum(p);
            if (previous!= null) {
                if (next!= null && next-1 != index && previous+1 != index) {
                    DocBase.setNumberedParagraph(p, 1, 2);
                }
                else {
                    if (data.get(previous) == data.get(index) && previous + 1 == index)
                        character = ";";
                    else if (previous + 1 < index || data.get(previous) < data.get(index)){
                        character = ".";
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
            }
            else {
                if (next!= null && next-1 != index) {
                    DocBase.setNumberedParagraph(p, 1, 2);
                }
            }
            if (data.get(index) == 1) {
                DocBase.setNumberedParagraph(p, 1, 3);
            }
            else {
                DocBase.setNumberedParagraph(p, 1, 4);
            }
            previous = index;
            if (!it.hasNext())
                DocBase.setText(p, ".", false);
        }
    }


    private void processImage (MainDocumentPart mdp) {
        int number = 0;
        List<Object> contens;
        Object o1 = null, o2 = null;
        for (int i = 0; i < mdp.getContent().size(); i++) {
            Object o = mdp.getContent().get(i);
            contens = DocxMethods.getAllElementFromObject(o, Drawing.class);
            if (contens.size() != 0) {
                for (Object content : contens) {
                    if (content instanceof Drawing) {
                        List<Object> anchors = ((Drawing) content).getAnchorOrInline();
                        for (Object anchor : anchors) {
                            if (anchor instanceof Anchor) {
                                ((Anchor) anchor).getEffectExtent().setB(0L);
                                ((Anchor) anchor).getEffectExtent().setT(0L);
                                ((Anchor) anchor).getEffectExtent().setL(0L);
                                ((Anchor) anchor).getEffectExtent().setR(0L);

                                ((Anchor) anchor).getExtent().setCx(0L);
                                ((Anchor) anchor).getExtent().setCy(0L);
                            }
                        }
                    }
                }
                mdp.getContent().add(i+1, createIt());
//                number++;
//                if (i!= 0) o1 = mdp.getContent().get(i-1);
//                if (i != mdp.getContent().size()-1) o2 = mdp.getContent().get(i-1);
//                boolean isName1 = false, isName2 = false;
//                String s = "";
//                if (o1!= null && o1 instanceof P) {
//                    s = DocBase.getText((P)o1);
//                    if (s.toLowerCase().contains("рис") || s.toLowerCase().contains("рисунок")) {
//                        if (!s.toLowerCase().contains("см.")&&!s.toLowerCase().contains("смотри")
//                                &&!s.toLowerCase().contains("котор")&& !s.toLowerCase().contains("(")&&
//                                s.indexOf("(")!= 0)
//                            isName1 = true;
//                    }
//                }
//                if (o2 != null && o2 instanceof P) {
//                    s = DocBase.getText((P)o1);
//                    if (s.toLowerCase().contains("рис") || s.toLowerCase().contains("рисунок") || s.toLowerCase().
//                            contains("Илюстрация") || s.toLowerCase().contains("Граф")) {
//                        if (!s.toLowerCase().contains("см.")&&!s.toLowerCase().contains("смотри")
//                                &&!s.toLowerCase().contains("котор")&& !s.toLowerCase().contains("(")&&
//                                s.indexOf("(")!= 0)
//                            isName2 = true;
//                    }
//                }
//                if (isName1) {
//                    s = DocBase.getText((P)o1);
//                    mdp.getContent().remove(o1);
//                }
//                if (isName1 || isName2) {
////                    s = s.replaceAll("Рисунок", "");
////                    s = s.replaceAll("Рис", "");
////                    s = s.replaceAll("Илюстрация", "");
////                    s = s.replaceAll("Граф","");
//                    Pattern p = Pattern.compile("\\d [: -\".,=]*");
//                    Matcher m = p.matcher(s);
//                    int lastDigit = m.regionEnd();
//                    String name = s.substring(lastDigit, s.length()-1);
//                    if (isName2) {
//                        if (DocBase.getText((P)o2).matches("Рисунок \\d -[\\s\\S]")) {
//                            break;
//                        }
//                    }
//                    P pWithName = factory.createP();
//                    DocBase.setText(pWithName, "Рисунок" + number +" - "+name, false);
//                    mdp.getContent().add(i+1, createIt());
//                }
            }
        }
    }
    public P createIt() {
        P p = factory.createP();
        // Create object for r
        R r = factory.createR();
        p.getContent().add(r);
        // Create object for t (wrapped in JAXBElement)
        Text text = factory.createText();
        JAXBElement<org.docx4j.wml.Text> textWrapped = factory
                .createRT(text);
        r.getContent().add(textWrapped);
        text.setValue("Рисунок ");
        text.setSpace("preserve");
        // Create object for fldSimple (wrapped in JAXBElement)
        CTSimpleField simplefield = factory.createCTSimpleField();
        JAXBElement<org.docx4j.wml.CTSimpleField> simplefieldWrapped = factory
                .createPFldSimple(simplefield);
        p.getContent().add(simplefieldWrapped);
        // Create object for r
        R r2 = factory.createR();
        simplefield.getContent().add(r2);
        // Create object for t (wrapped in JAXBElement)
        Text text2 = factory.createText();
        JAXBElement<org.docx4j.wml.Text> textWrapped2 = factory
                .createRT(text2);
        r2.getContent().add(textWrapped2);
        text2.setValue("1");
        // Create object for rPr
        RPr rpr = factory.createRPr();
        r2.setRPr(rpr);
        // Create object for noProof
        BooleanDefaultTrue booleandefaulttrue = factory
                .createBooleanDefaultTrue();
        rpr.setNoProof(booleandefaulttrue);
        simplefield.setInstr(" SEQ Figure \\* ARABIC ");
        // Create object for r
        R r3 = factory.createR();
        p.getContent().add(r3);
        // Create object for t (wrapped in JAXBElement)
        Text text3 = factory.createText();
        JAXBElement<org.docx4j.wml.Text> textWrapped3 = factory
                .createRT(text3);
        r3.getContent().add(textWrapped3);
        text3.setValue(" ");
        text3.setSpace("preserve");
        // Create object for r
        R r4 = factory.createR();
        p.getContent().add(r4);
        // Create object for t (wrapped in JAXBElement)
        Text text4 = factory.createText();
        JAXBElement<org.docx4j.wml.Text> textWrapped4 = factory
                .createRT(text4);
        r4.getContent().add(textWrapped4);
        text4.setValue("–");
        // Create object for r
        R r5 = factory.createR();
        p.getContent().add(r5);
        // Create object for t (wrapped in JAXBElement)
        Text text5 = factory.createText();
        JAXBElement<org.docx4j.wml.Text> textWrapped5 = factory
                .createRT(text5);
        r5.getContent().add(textWrapped5);
        text5.setValue("This is the caption of the figure");
        text5.setSpace("preserve");
        // Create object for bookmarkStart (wrapped in JAXBElement)
        CTBookmark bookmark = factory.createCTBookmark();
        JAXBElement<org.docx4j.wml.CTBookmark> bookmarkWrapped = factory
                .createPBookmarkStart(bookmark);
        p.getContent().add(bookmarkWrapped);
        bookmark.setName("_GoBack");
        bookmark.setId(BigInteger.valueOf(0));
        // Create object for bookmarkEnd (wrapped in JAXBElement)
        CTMarkupRange markuprange = factory.createCTMarkupRange();
        JAXBElement<org.docx4j.wml.CTMarkupRange> markuprangeWrapped = factory
                .createPBookmarkEnd(markuprange);
        p.getContent().add(markuprangeWrapped);
        markuprange.setId(BigInteger.valueOf(0));
        // Create object for pPr
        PPr ppr = factory.createPPr();
        p.setPPr(ppr);
        // Create object for pStyle
        PPrBase.PStyle pprbasepstyle = factory.createPPrBasePStyle();
        ppr.setPStyle(pprbasepstyle);
        pprbasepstyle.setVal("Caption");
        // Create object for jc
        Jc jc = factory.createJc();
        ppr.setJc(jc);
        jc.setVal(org.docx4j.wml.JcEnumeration.CENTER);

        return p;
    }

}
