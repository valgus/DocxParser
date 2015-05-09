package Model;

import org.docx4j.convert.out.common.preprocess.CoverPageSectPrMover;
import org.docx4j.convert.out.common.preprocess.ParagraphStylesInTableFix;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;

import java.io.File;

public class Main {
    public static void main(String[] args) {

        MainPart comparisonWithTemplate = new MainPart();
        comparisonWithTemplate.setTwoDocx("docx/template.docx","docx/document.docx");
        try {
           WordprocessingMLPackage word =  comparisonWithTemplate.setAppropriateText();

            EditingFirstPages editingFirstPages = new EditingFirstPages(word,"Руководство оператора" ,19, "программа проверки на соответствие госту");
            word = editingFirstPages.processDoc();
            System.out.println(word.getMainDocumentPart().getContent().size());

            word.save(new File("2.docx"));
            CoverPageSectPrMover.process(word);
            ParagraphStylesInTableFix.process(word);
            DocxToPDFConverter.convert(word);
        } catch (Exception e) {
            e.printStackTrace();
        }

    }
//    private static ObjectFactory factory = Context.getWmlObjectFactory();
//    public static void main(String[] args) throws InvalidFormatException, JAXBException {
//        WordprocessingMLPackage word = WordprocessingMLPackage.createPackage();
//        Styles styles = (Styles)word.getMainDocumentPart().getStyleDefinitionsPart().unmarshalDefaultStyles();
//        StyleDefinitionsPart styleDefinitionsPart = new StyleDefinitionsPart();
//        styleDefinitionsPart.setPackage(word);
//        styleDefinitionsPart.setJaxbElement(styles);
//        word.getMainDocumentPart().addTargetPart(styleDefinitionsPart);
//        P p = factory.createP();
//        DocBase.setText(p, "GHbdtn", false);
//        DocBase.setStyle(p, null, "Arial", null, "0", "0", 0, "RIGHT");
//        DocBase.setBold(p, true);
//       // DocBase.setAlign(p, "RIGHT");
//        word.getMainDocumentPart().addObject(p);
//        try {
//            word.save(new File("4.docx"));
//        } catch (Docx4JException e) {
//            e.printStackTrace();
//        }
//
//    }

}