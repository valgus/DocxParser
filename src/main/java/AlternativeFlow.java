import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.wml.NumFmt;
import org.docx4j.wml.P;

import java.io.File;
import java.math.BigInteger;
import java.util.*;

public class AlternativeFlow {

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
        List<Title> titles = TemplateParser.getListOfTitles(template.getAbsolutePath());
        List<Object> documentParagraphes = DocxMethods.createParagraphJAXBNodes(document);
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
                    DocBase.setSpacing(p, true);
                    t = findTitle(titles, p);
                    if (t!=null) {
                        DocBase.setText(p, t.getName());
                        String[] atr = DocBase.getAttributes(t);
                        DocBase.setStyle(p, null, null, atr[0], atr[1], atr[2]);
                        DocBase.setHighlight(p, "green");
                    }
                    else {
                        DocBase.setStyle(p, null, null, "1", null, null);
                        DocBase.setHighlight(p, "green");
                    }
                } while (it.hasNext());
            }
        }
        for (Object o : documentParagraphes) {   //setAttributes
            P p = (P) o;
            DocBase.setSpacing(p, false);
            if (!DocBase.getText(p).trim().equals("") & DocBase.getHighlight(p)==null)
                DocBase.setStyle(p, "28","Times New Roman", null, null, null);
        }

        changeEnumeration(document, documentParagraphes);

        for (Object o : documentParagraphes) {
            P p = (P) o;
            if (DocBase.getHighlight(p)!=null) {
                DocBase.setHighlight(p, null);
            }
        }


        document.save(new File("docx/2.docx"));

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

    private P getParagraphFromIndex (WordprocessingMLPackage wordprocessingMLPackage,int i) {
        return (P)wordprocessingMLPackage.getMainDocumentPart().getContent().get(i);
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
            DocBase.setSpacing(p, false);
            if (!DocBase.getText(p).trim().equals("") & DocBase.getHighlight(p)==null) {
                if (!DocBase.getLevelInList(p).equals(not) ) {

                    int i = getIndexOfParagraph(document, p);
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
            P p = getParagraphFromIndex(document, index);
            if (previous != index -1) {
                id = String.valueOf(Integer.decode(id)+1);
            }
            DocBase.setList(p);
            previous = index;
        }
    }
}