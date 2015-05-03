import org.docx4j.dml.wordprocessingDrawing.Inline;
import org.docx4j.jaxb.Context;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.BinaryPartAbstractImage;
import org.docx4j.sharedtypes.STOnOff;
import org.docx4j.wml.*;
import java.io.*;
import java.math.BigInteger;

public class Main {
    public static void main(String[] args) {
  //    AlternativeFlow comparisonWithTemplate = new AlternativeFlow();
        TableOfContentsAdd comparisonWithTemplate = new TableOfContentsAdd();
        comparisonWithTemplate.setTwoDocx("docx/template.docx","docx/gost19_tehnicheskoe_zadanie.docx");
        try {
            comparisonWithTemplate.setAppropriateText();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

}