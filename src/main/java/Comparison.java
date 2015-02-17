import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.wml.P;
import org.docx4j.wml.Text;

import java.io.File;
import java.util.*;

public class Comparison {

    File template, compared;
    boolean exist = false;


    public void setTwoDocx(String template, String compared) {
        this.template = new File(template);
        this.compared = new File(compared);
        exist = true;
    }

    public void setAppropriateText() throws Exception {
        if (!exist) throw new Exception();
        WordprocessingMLPackage document = DocxMethods.getTemplate(compared.getAbsolutePath());
        WordprocessingMLPackage template1 = DocxMethods.getTemplate(template.getAbsolutePath());
        List<Title> titles = TemplateParser.getListOfTitles(template.getAbsolutePath());
        List<Object> templateParagraphes = template1.getMainDocumentPart().getContent();
        List<Object> documentParagraphes = DocxMethods.createJaxbNodes(document);
        Map<Integer, P> corespondences = findCorespondences(document, documentParagraphes, titles);
        int notP = 0;
        if (corespondences.size() == 0) {
            throw new Exception("Файл не соответствует шаблону!");
        }
        else {
            Collection<P> values = corespondences.values();
            Iterator it = values.iterator();
            int indexInTemplate;
            int indexInDocument;
            String t;
            Object jaxbNode = it.hasNext() == true ? (P)it.next(): null;
            do {
                if (jaxbNode instanceof P) {
                    P p = (P) jaxbNode;
                    indexInDocument = getIndexOfParagraph(document, p);
                    indexInTemplate = getIndexOfParagraph(template1, findP(templateParagraphes, DocBase.getText(p)));
                    t = DocBase.getText((P)documentParagraphes.get(++indexInDocument));
                    p = (P)it.next();
                    String next = DocBase.getText(p);
                    while (!t.toLowerCase().equals(next.toLowerCase())) {
                        ++indexInTemplate;
                        DocBase.addParagraph(template1.getMainDocumentPart(), t,  indexInTemplate);
                        ++indexInDocument;
                        P pp;
                        while (true) {
                            pp = getParagraphFromIndex(document, indexInDocument);
                            if (pp != null)
                                break;
                            ++indexInDocument;
                            notP++;
                        }
                        t = DocBase.getText(pp);
                    }
                }
            } while (it.hasNext());
            indexInDocument = getIndexOfParagraph(document, jaxbNode);
            indexInTemplate = getIndexOfParagraph(template1, findP(templateParagraphes, DocBase.getText(jaxbNode)));
            t = DocBase.getText((P)documentParagraphes.get(++indexInDocument));
            while (indexInDocument+1 < documentParagraphes.size()-notP) {
                ++indexInTemplate;
                DocBase.addParagraph(template1.getMainDocumentPart(), t,  indexInTemplate);
                ++indexInDocument;
                t = DocBase.getText(getParagraphFromIndex(document, indexInDocument));
            }
        }
        setStyle(template1, titles);
        template1.save(new File("docx/2.docx"));
    }

    private void setStyle (WordprocessingMLPackage document,List<Title> titles ) {
        //set Spacing
        List<Object> contents =  document.getMainDocumentPart().getContent();
        HashMap<P, Title> paragraphes = new HashMap<>();
        for (Object content : contents) {
            paragraphes.put((P)content, null);
        }
        for (Title title : titles) {
            P p = findP(contents, title.getName());
            if (p!= null && paragraphes.containsKey(p)) {
                paragraphes.put(p, title);
            }
        }
        contents = null;
        P p;
        Title t;
        for (Map.Entry<P,Title> paragraph : paragraphes.entrySet()){
            p = paragraph.getKey();
            t = paragraph.getValue();
            if (paragraph.getValue() != null) {
                DocBase.setSpacing(p, true);
                String[] atr = DocBase.getAttributes(t);
                DocBase.setStyle(p, Boolean.getBoolean(atr[3]), Boolean.getBoolean(atr[4]), atr[2],
                        atr[0],Boolean.getBoolean(atr[5]), atr[1]);
            }
            else{
                DocBase.setSpacing(p, false);
               if (!DocBase.getText(p).trim().equals(""))
                    DocBase.setStyle(p, false, false, "left", "28",false,"Times New Roman");
            }
        }

    }

    private Map<Integer, P> findCorespondences(WordprocessingMLPackage wordprocessingMLPackage,List<Object> document, List<Title> titles) {
        Map<Integer, P> check = new TreeMap<>();
        int index;
        for (int i = 0; i< titles.size(); i++) {
            P p = findP(document, titles.get(i).getName());
            if (p != null) {
                check.put(getIndexOfParagraph(wordprocessingMLPackage, p), p);
            }
        }
        return check;
    }

    private P findP (List<Object> paragraphes, String s) throws ClassCastException{
        int i = 0;
        P paragraph = null;
        while ( paragraphes.size()> i) {
            String pTemplate = DocBase.getText((P) paragraphes.get(i)).toLowerCase();
            if (s.toLowerCase().equals(pTemplate)) {
                paragraph = (P) paragraphes.get(i);
                break;
            }
            i++;
        }

        return paragraph;
    }

    private int getIndexOfParagraph (WordprocessingMLPackage wordprocessingMLPackage, P p) {
        return wordprocessingMLPackage.getMainDocumentPart().getContent().indexOf(p);
    }

    private P getParagraphFromIndex (WordprocessingMLPackage wordprocessingMLPackage, int i) {
        try{
         return (P)wordprocessingMLPackage.getMainDocumentPart().getContent().get(i);
        }
        catch (ClassCastException ex) {
            return null;
        }
    }
}

//        int index, Tindex;
//for (int i =0; i < titles.size(); i++) {
//            Title t = titles.get(i);
//            P currentT = findP(templateParagraphes, t.getName());
//            P currentD = findP(documentParagraphes, t.getName());
//            DocBase.setSpacing(currentT, t.getLvl()==0 ? true: false);
//            if (currentD!= null) {
//                index = getIndexOfParagraph(document, currentD)+1;
//                Tindex = getIndexOfParagraph(template1, currentT);
//                String text = DocBase.getText(getParagraphFromIndex(document, index));
//                if (i + 1 < titles.size()) {
//                    while (!titles.get(i+1).getName().toLowerCase().equals(text.toLowerCase())) {
//                        ++Tindex;
//                        DocBase.addParagraph(template1.getMainDocumentPart(), text,  Tindex);
//                        ++index;
//                        text = DocBase.getText(getParagraphFromIndex(document, index));
//                    }
//                }
//                else {
//                    while (index + 1 < document.getMainDocumentPart().getContent().size()) {
//                        Tindex++;
//                        DocBase.addParagraph(template1.getMainDocumentPart(), text,  Tindex);
//                        index++;
//                        text = DocBase.getText(getParagraphFromIndex(document, index));
//                    }
//                }
//            }
//        }