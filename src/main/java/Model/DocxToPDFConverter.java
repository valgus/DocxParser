package Model;

import org.docx4j.Docx4J;
import org.docx4j.convert.out.FOSettings;
import org.docx4j.convert.out.pdf.viaXSLFO.PdfSettings;
import org.docx4j.fonts.IdentityPlusMapper;
import org.docx4j.fonts.Mapper;
import org.docx4j.fonts.PhysicalFont;
import org.docx4j.fonts.PhysicalFonts;
import org.docx4j.model.fields.FieldUpdater;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.samples.AbstractSample;

import java.io.OutputStream;

public class DocxToPDFConverter extends AbstractSample {


    static {
        inputfilepath = System.getProperty("user.dir") + "/docx/document.docx";
        saveFO = true;
    }


    static boolean saveFO;

    public static void convert (WordprocessingMLPackage wordMLPackage) throws Exception {

        String regex = ".*(calibri|camb|cour|arial|times|comic|georgia|impact|LSANS|pala|tahoma|trebuc|verdana|symbol|webdings|wingding).*";
        PhysicalFonts.setRegex(regex);



        if (inputfilepath==null) {
            throw new Exception("No input path passed, creating dummy document");

        }
        boolean result = true;
        try {
            org.docx4j.convert.out.pdf.PdfConversion c = new org.docx4j.convert.out.pdf.viaXSLFO.Conversion(wordMLPackage);
            ((org.docx4j.convert.out.pdf.viaXSLFO.Conversion)c).setSaveFO(new java.io.File(inputfilepath+".fo"));
            ((org.docx4j.convert.out.pdf.viaXSLFO.Conversion)c).setSaveFO(new java.io.File(inputfilepath + ".fo"));
            OutputStream os = new java.io.FileOutputStream(inputfilepath.substring(0,inputfilepath.length()-2) + ".pdf");
            c.output(os, new PdfSettings());
            System.out.println("Saved " + inputfilepath + ".pdf");
        }
        catch (Exception ex) {
            result = false;
        }
        if (!result) {
            FieldUpdater updater = new FieldUpdater(wordMLPackage);
            updater.update(true);

            Mapper fontMapper = new IdentityPlusMapper();
            wordMLPackage.setFontMapper(fontMapper);

            PhysicalFont font
                    = PhysicalFonts.get("Arial Unicode MS");
            if (font!=null) {
                fontMapper.put("Times New Roman", font);
                fontMapper.put("Arial", font);
            }
              fontMapper.put("Libian SC Regular", PhysicalFonts.get("SimSun"));

            wordMLPackage.setFontMapper(new IdentityPlusMapper());
            FOSettings foSettings = Docx4J.createFOSettings();
            if (saveFO) {
                foSettings.setFoDumpFile(new java.io.File("dump" + ".fo"));

            }
            foSettings.setWmlPackage(wordMLPackage);

            String outputfilepath;
            if (inputfilepath==null) {
                outputfilepath = System.getProperty("user.dir") + "/OUT_FontContent.pdf";
            } else {
                outputfilepath = "result" + ".pdf";
            }
            OutputStream os = new java.io.FileOutputStream(outputfilepath);
            Docx4J.toFO(foSettings, os, Docx4J.FLAG_EXPORT_PREFER_NONXSL);
            System.out.println("Saved: " + outputfilepath);
        }
    }
}
