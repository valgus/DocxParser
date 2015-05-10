package Model;

import Model.DocBase;
import Model.DocxMethods;
import Model.Title;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.wml.P;

import java.io.*;
import java.util.*;

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
        if (p.getPPr().getPStyle().getVal().equals("1") && p.getPPr().getTabs() != null) {
            return true;
        }
        }
        catch (NullPointerException ex) {return false;}
        return false;
    }

    public static List<Title> getListOfTitles (String name) {
        List<Title> titles = new ArrayList<>();
        setDocument(name);
        List<Object> jaxbNodes = DocxMethods.createParagraphJAXBNodes(template);
        for (Object jaxbNode : jaxbNodes) {
            P p = (P)jaxbNode;
            int lvl = DocBase.getLevelInList(p).intValue();
            int id =  DocBase.getNumIDInList(p).intValue();
            if (id != -1) {
                Title currentTitle = new Title(lvl, DocBase.getText(p), DocBase.getAttributes(p));
                titles.add(currentTitle);
            }
            if (isTitle(p)) {
                Title currentTitle = new Title(0, DocBase.getText(p), DocBase.getAttributes(p));
                titles.add(currentTitle);
            }
            }
        return titles;
    }




//lvl == -1? lvl+1:
}


