import org.docx4j.jaxb.Context;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.openpackaging.parts.relationships.Namespaces;
import org.docx4j.wml.*;

import javax.xml.bind.JAXBElement;
import javax.xml.namespace.QName;
import java.io.File;
import java.math.BigInteger;
import java.util.*;

public class AlternativeFlow {

    File template, compared;
    boolean exist = false;
    ObjectFactory factory = Context.getWmlObjectFactory();
    String firstCoincedence;
    public void setTwoDocx(String template, String compared) {
        this.template = new File(template);
        this.compared = new File(compared);
        exist = true;
    }

    public void setAppropriateText() throws Exception {
        if (!exist) throw new Exception();
        WordprocessingMLPackage document = DocxMethods.getTemplate(compared.getAbsolutePath());
        List<Title> titles = TemplateParser.getListOfTitles(template.getAbsolutePath());
        findCorespondences(document, null, titles, true);


        //delete information before first coincedence
//        for (int i = 0; i <firstCoincedence; i++ ) {
//            document.getMainDocumentPart().getContent().remove(i);
//        }
        if (firstCoincedence == null)
            throw new Exception("Does not match template");
        P p = findP(document.getMainDocumentPart().getContent(), firstCoincedence);
        int i = DocxMethods.getIndexOfParagraph(document.getMainDocumentPart(), p);
        WordprocessingMLPackage newDocument = WordprocessingMLPackage.createPackage();

        newDocument.getMainDocumentPart().getContent().addAll
                (document.getMainDocumentPart().getContent().subList(
                        i, document.getMainDocumentPart().getContent().size() - 1));
        System.out.println(newDocument.getMainDocumentPart().getContent().size());
        System.out.println(document.getMainDocumentPart().getContent().size());
        List<Object> documentParagraphes = DocxMethods.createParagraphJAXBNodes(newDocument);
        Map<Integer, P> corespondences = findCorespondences(newDocument, documentParagraphes, titles, false);

        if (corespondences.size() < 2 ) {
            throw new Exception("Does not match template");
        }

        Collection<P> values = corespondences.values();
        Iterator it = values.iterator();
        Title t;
        if (it.hasNext()) {
            do{
                p = (P)it.next();
                t = findTitle(titles, p);
                if (t!=null) {
                    DocBase.setText(p, t.getName(), true);
                    String[] atr = DocBase.getAttributes(t);
                    DocBase.setStyle(p, null, null, atr[0], atr[1], atr[2], 0);
                    DocBase.setHighlight(p, "green");
                }
                else {
                    DocBase.setStyle(p, null, null, "1", null, null, 0);
                    DocBase.setHighlight(p, "green");
                }
            } while (it.hasNext());
        }
//        for (Object o : documentParagraphes) {   //setAttributes
//            p = (P) o;
//            if (!DocBase.getText(p).trim().equals("") & DocBase.getHighlight(p)==null)
//                DocBase.setStyle(p, "28","Times New Roman", null, null, null, 360);
//        }

      //  changeEnumeration(newDocument, documentParagraphes);

        for (Object o : documentParagraphes) {
            p = (P) o;
            if (DocBase.getHighlight(p)!=null) {
                DocBase.setHighlight(p, null);
            }
        }

        //set table of contents
//        P paragraphForTOC = factory.createP();
//        R r = factory.createR();
//
//        FldChar fldchar = factory.createFldChar();
//        fldchar.setFldCharType(STFldCharType.BEGIN);
//        fldchar.setDirty(true);
//        r.getContent().add(getWrappedFldChar(fldchar));
//        paragraphForTOC.getContent().add(r);
//
//        R r1 = factory.createR();
//        Text txt = new Text();
//        txt.setSpace("preserve");
//        txt.setValue("TOC \\o \"1-3\" \\h \\z \\u \\h");
//        r.getContent().add(factory.createRInstrText(txt) );
//        paragraphForTOC.getContent().add(r1);
//
//        FldChar fldcharend = factory.createFldChar();
//        fldcharend.setFldCharType(STFldCharType.END);
//        R r2 = factory.createR();
//        r2.getContent().add(getWrappedFldChar(fldcharend));
//        paragraphForTOC.getContent().add(r2);
//        newDocument.getMainDocumentPart().getContent().add(0, paragraphForTOC);

        File docx = new File("2.docx") ;
        newDocument.save(docx);
        //  documentParagraphes = DocxMethods.createTableJAXBNodes(DocxMethods.getTemplate(docx.getAbsolutePath()));
//        if (true)
//
//            for (Object o : documentParagraphes){
//             //   SdtBlock sdtBlock = (SdtBlock) o;
//                System.out.println(o);
//
//            }
    }

    private Map<Integer, P> findCorespondences(WordprocessingMLPackage uml,
                                               List<Object> document, List<Title> titles, boolean all) {
        Map<Integer, P> check = new TreeMap<>();
        int index;
        for (int i = 0; i< titles.size(); i++) {
            P p = (!all) ? findP(document, titles.get(i).getName()) :
                    findP(uml.getMainDocumentPart().getContent(), titles.get(i).getName()) ;
            if (p != null) {
                if (firstCoincedence == null & all) {
                    firstCoincedence = DocBase.getText(p);
                    return null;
                }
                check.put(DocxMethods.getIndexOfParagraph(uml.getMainDocumentPart(), p), p);
            }
        }
        return check;
    }

    private P findP (List<Object> paragraphes, String s) throws ClassCastException{
        int i = 0;
        P paragraph = null;
        while ( paragraphes.size()> i) {
            if (paragraphes.get(i) instanceof P) {
                String pTemplate = DocBase.getText((P) paragraphes.get(i)).toLowerCase();
                if (s.toLowerCase().equals(pTemplate)) {
                    paragraph = (P) paragraphes.get(i);
                    break;
                }
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
