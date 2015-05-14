package Model;

import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.wml.P;

import java.io.FileNotFoundException;
import java.util.ArrayList;
import java.util.List;

public class TemplateParser {

    private static WordprocessingMLPackage template;

    private static boolean setDocument(String name) {
        try {
            template = DocxMethods.getTemplate(name);
        } catch (Docx4JException e) {
            e.printStackTrace();
            return false;
        } catch (FileNotFoundException e) {
            System.out.println("File not found!");
            return false;
        }
        return true;
    }

    private static boolean isTitle (P p){
        try {
            if (p.getPPr().getPStyle().getVal()!= null) {
                try{
                    int styleid = Integer.valueOf(p.getPPr().getPStyle().getVal());
                    return true;
                }
                catch (NumberFormatException ex) {
                    return false;
                }
            }
        }
        catch (NullPointerException ex) {return false;}
        return false;
    }

    public static List<Title> getListOfTitles (String name) {
        List<Title> titles = new ArrayList<>();
        P lastTitle = null;
        int lastTitleIndex = 1;
        setDocument(name);
        List<Object> jaxbNodes = DocxMethods.createParagraphJAXBNodes(template);
        for (int i = 0; i < jaxbNodes.size(); i++) {
            P p = (P)jaxbNodes.get(i);
            if (isTitle(p) && !DocBase.getText(p).equals("") ||
                    DocBase.getText(p).toLowerCase().equals("приложение") ||
                    DocBase.getText(p).toLowerCase().equals("приложения")  ) {
                Title currentTitle = new Title(0, DocBase.getText(p), DocBase.getAttributes(p));
                if (lastTitle!=null) {
                    titles.get(titles.size()-1).setDescription(jaxbNodes.subList(lastTitleIndex+1, i));
                }
                lastTitle = p;
                lastTitleIndex = i;
                titles.add(currentTitle);
            }
        }
        if (lastTitleIndex!= jaxbNodes.size()-1)
            titles.get(titles.size()-1).setDescription(jaxbNodes.subList(lastTitleIndex+1, jaxbNodes.size()-1));
        for (Title title : titles) {
            DocBase.deleteEmptyPara(title.getDescription());
        }
        return titles;
    }


}


