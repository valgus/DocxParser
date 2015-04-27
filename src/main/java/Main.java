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
    static  WordprocessingMLPackage doc;
   /* public static void main(String[] args) {
      AlternativeFlow comparisonWithTemplate = new AlternativeFlow();
        comparisonWithTemplate.setTwoDocx("docx/template.docx","docx/gost19_tehnicheskoe_zadanie.docx");
        try {
            comparisonWithTemplate.setAppropriateText();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }*/
    public static ObjectFactory factory = Context.getWmlObjectFactory();
    public static void main(String[] args) throws Exception{
        doc = DocxMethods.getTemplate("docx/2.docx");
        setSecondPage();
        try {
            doc.save(new File("3.docx"));
        } catch (Docx4JException e) {
            e.printStackTrace();
        }
    }

    private static void setSecondPage() throws Exception{
     //   doc.getMainDocumentPart().getContent().add(setP("", "Times New Roman", null, null, null, 480, null, null));
        doc.getMainDocumentPart().getContent().add(setP("", "Times New Roman", null, null, null, 480, null, null));
        doc.getMainDocumentPart().getContent().add(setP("", "Times New Roman", null, null, null, 480, null, null));
        CTVerticalJc ctVerticalJc = new CTVerticalJc();
        ctVerticalJc.setVal(STVerticalJc.CENTER);
        Tbl table = new Tbl();
        table.setTblPr(new TblPr());
        TblWidth width = new TblWidth();
        width.setType("pct");
        width.setW(BigInteger.valueOf(4500));
        table.getTblPr().setTblW(width);
        CTTblPrBase.TblStyle style = new CTTblPrBase.TblStyle();
        style.setVal("a3");
        table.getTblPr().setTblStyle(style);
        TblGrid tblGrid = Context.getWmlObjectFactory().createTblGrid();
        table.setTblGrid(tblGrid);
        TblWidth width2 = factory.createTblWidth();
        width2.setType("dxa");
        width2.setW(BigInteger.valueOf(15));
        table.getTblPr().setTblCellSpacing(width2);
        CTTblCellMar cellMar = new CTTblCellMar();
        cellMar.setBottom(width2);
        cellMar.setLeft(width2);
        cellMar.setRight(width2);
        cellMar.setTop(width2);
        table.getTblPr().setTblCellMar(cellMar);
        TblGridCol tblGridCol1 = new TblGridCol();
        tblGridCol1.setW(BigInteger.valueOf(500));
        TblGridCol tblGridCol2 = new TblGridCol();
        tblGridCol2.setW(BigInteger.valueOf(1500));
        TblGridCol tblGridCol3 = new TblGridCol();
        tblGridCol3.setW(BigInteger.valueOf(1500));
        TblGridCol tblGridCol4 = new TblGridCol();
        tblGridCol4.setW(BigInteger.valueOf(1500));
        table.getTblGrid().getGridCol().add(tblGridCol1);
        table.getTblGrid().getGridCol().add(tblGridCol2);
        table.getTblGrid().getGridCol().add(tblGridCol3);
        table.getTblGrid().getGridCol().add(tblGridCol4);
        CTTblLook tblLook = new CTTblLook();
        tblLook.setNoVBand(STOnOff.ONE);
        tblLook.setVal("04A0");
        tblLook.setNoHBand(STOnOff.ZERO);
        tblLook.setLastColumn(STOnOff.ZERO);
        tblLook.setFirstColumn(STOnOff.ONE);
        tblLook.setLastRow(STOnOff.ZERO);
        tblLook.setFirstRow(STOnOff.ONE);
        table.getTblPr().setTblLook(tblLook);

        File file = new File("resource/table.gif");
        P[] pr1 = {setP("УТВЕРЖДЕНО", "Times New Roman", null, null, null, 480, null, null),
                setP("А.В.00001-01 33 01-1-ЛУ", "Courier New", null, null, null, 240, "LEFT", "20"),
                setP("", "Times New Roman", null, null, null, 240, null, null),setP("", "Times New Roman", null, null, null, 240, null, null),
                setP("", "Times New Roman", null, null, null, 240, null, null),setP("", "Times New Roman", null, null, null, 240, null, null)};

        P[] pr = {setP("", "Times New Roman", null, null, null, 240, null, null) };
        table.getContent().add(addRowWithMergedCells(file, pr1,pr,pr, 1500,1500, 1500, 1 ));

        P[] pr3 = {
                setP("ЕДИНАЯ СИСТЕМА ЭЛЕКТРОННЫХ ВЫЧИСЛИТЕЛЬНЫХ МАШИН ОПЕРАЦИОННАЯ СИСТЕМА", "Times New Roman",null, null, null, 240, null, null),
                setP("", "Arial","30", "0", "0", 120, null, null),
                setP("Загрузчик", "Arial","30", "0", "0", 280, null, null),
                setP("Руководство программиста", "Arial", "30", "0", "0", 360, null, null),
                setP("112.23.568.89-1", "Arial", "30", "0", "0", 280, null, null),
                setP("(вид носителя данных)", "Arial", "30", "0", "0", 600, null, null),
                setP("Листов 13", "Arial", "30", "0", "0", 360, null, null)};
        table.getContent().add(addRowWithMergedCells(null, null, pr3, null, 0, 4500, 0, 2));
        P[] pr_ = {new P(), new P(), new P(), new P(),new P(), new P(), new P(), new P(),new P(), new P(), new P(), new P(),
                new P(), new P(), new P(), new P(),new P(), new P(), new P(), new P()};
        table.getContent().add(addRowWithMergedCells(null, pr, pr_, pr, 1500, 1500, 1500, 3));
        P[] pr5 = {setP("1982", "Times New Roman", null, null, null, 240, null, null),
                setP("Литера", "Times New Roman", null, null, null, 240, "RIGHT", null)};
        table.getContent().add(addRowWithMergedCells(null, null, pr5, null, 0, 4500, 0, 4));
        doc.getMainDocumentPart().getContent().add(table);
    }

    private static Tr addRowWithMergedCells(File file, P[] ps, P[] ps2, P[] ps3, int width1, int width2 , int width3, int num) throws  Exception{

        Tr row = factory.createTr();
        row.setTrPr(factory.createTrPr());

        if (num == 1) {
            addMergedColumn(row, file, 0, 500);
            addTableCell(row, ps, width1, 1);
            addTableCell(row, ps2, width2, 2);
            addTableCell(row, ps3, width3, 3);
        }
        else if (num == 2)
        {
            addMergedColumn(row, file, -1, 0);
            addTableCell(row, ps2, width2, 6);
        }
        else if (num == 3)
        {
            addMergedColumn(row, file, 0, 0);
            addTableCell(row, ps, width1, 1);
            addTableCell(row, ps2, width2, 2);
            addTableCell(row, ps3, width3, 3);
        }
        else
        {
            addMergedColumn(row, file, -1, 0);
            addTableCell(row, ps2, width2, 6);
        }
        return row;
    }

    private static void addMergedColumn(Tr row, File file, int grid, int width) throws Exception{
        if (file == null) {
            addMergedCell(row, null, null, grid, width);
        } else {
            addMergedCell(row, file, "restart", grid, width);
        }
    }

    private static void addMergedCell(Tr row, File file, String vMergeVal, int grid, int width) throws  Exception{
        Tc tableCell = factory.createTc();
        TcPr tableCellProperties = new TcPr();
        if(file != null) {
            tableCell.getContent().add(newImage());
        }
        TcPrInner.VMerge merge = new TcPrInner.VMerge();
        merge.setVal(vMergeVal);
        CTVerticalJc ctVerticalJc = new CTVerticalJc();
        ctVerticalJc.setVal(STVerticalJc.BOTTOM);
        tableCellProperties.setVAlign(ctVerticalJc);
        tableCellProperties.setVMerge(merge);

        tableCell.setTcPr(tableCellProperties);
        if (grid!=-1) {
            TcPrInner.GridSpan gridSpan = new TcPrInner.GridSpan();
            gridSpan.setVal(BigInteger.valueOf(grid));
            tableCellProperties.setGridSpan(gridSpan);
        }
        TblWidth tableWidth = factory.createTblWidth();
        tableWidth.setType("dxa");
        tableWidth.setW(BigInteger.valueOf(width));
        tableCell.getTcPr().setTcW(tableWidth);
        row.getContent().add(tableCell);
    }

    private static void addTableCell(Tr tr, P[] content, int width, int grid) {
        Tc tc1 = factory.createTc();
        for (P p : content) {
            tc1.getContent().add(p);
        }
        TcPr tableCellProperties = new TcPr();
        TblWidth tableWidth = factory.createTblWidth();
        tableWidth.setType("dxa");
        tableWidth.setW(BigInteger.valueOf(width));
        CTVerticalJc ctVerticalJc = factory.createCTVerticalJc();
        ctVerticalJc.setVal(STVerticalJc.CENTER);
        tableCellProperties.setVAlign(ctVerticalJc);
        BooleanDefaultTrue value = new BooleanDefaultTrue();

        tableCellProperties.setHideMark(value);
        if (grid!=-1) {
            TcPrInner.GridSpan gridSpan = new TcPrInner.GridSpan();
            gridSpan.setVal(BigInteger.valueOf(grid));
            tableCellProperties.setGridSpan(gridSpan);
        }
        tableCellProperties.setTcW(tableWidth);
        tc1.setTcPr(tableCellProperties);
        tr.getContent().add(tc1);
    }

    private static P setP (String text, String font, String style, String ilvl, String numId, int spacing, String align, String size) {
        P p = new P();
        DocBase.setText(p, text, false);
        p.setPPr(new PPr());
        DocBase.setStyle(p, size, font, style, ilvl, numId, spacing);
        Jc jc = factory.createJc();
        if (align!= null && align.equals("RIGHT"))
            jc.setVal(JcEnumeration.RIGHT);
        else if (align!= null && align.equals("LEFT"))
            jc.setVal(JcEnumeration.LEFT);
        else
            jc.setVal(JcEnumeration.CENTER);
        p.getPPr().setJc(jc);
        R run = factory.createR();
        p.getContent().add(run);
        return p;
    }

    public static org.docx4j.wml.P newImage() throws Exception {

        File file = new File("resource/table.gif" );
        //File file = new File("C:\\Documents and Settings\\Jason Harrop\\My Documents\\LANL\\fig1.pdf" );

        // Our utility method wants that as a byte array

        java.io.InputStream is = new java.io.FileInputStream(file );
        long length = file.length();
        // You cannot create an array using a long type.
        // It needs to be an int type.
        if (length > Integer.MAX_VALUE) {
            System.out.println("File too large!!");
        }
        byte[] bytes = new byte[(int)length];
        int offset = 0;
        int numRead;
        while (offset < bytes.length
                && (numRead=is.read(bytes, offset, bytes.length-offset)) >= 0) {
            offset += numRead;
        }
        // Ensure all the bytes have been read in
        if (offset < bytes.length) {
            System.out.println("Could not completely read file "+file.getName());
        }
        is.close();

        String filenameHint = null;
        String altText = null;
        int id1 = 0;
        int id2 = 1;

        BinaryPartAbstractImage imagePart = BinaryPartAbstractImage.createImagePart(doc, bytes);

        Inline inline = imagePart.createImageInline( filenameHint, altText,
                id1, id2, true);

        org.docx4j.wml.P  p = factory.createP();
        p.setPPr(factory.createPPr());
        org.docx4j.wml.R  run = factory.createR();
        RPr rPr = factory.createRPr();
        BooleanDefaultTrue value = new BooleanDefaultTrue();
        rPr.setNoProof(value );
        p.getContent().add(run);
        org.docx4j.wml.Drawing drawing = factory.createDrawing();
        run.getContent().add(drawing);
        drawing.getAnchorOrInline().add(inline);
        Text t = factory.createText();
        run.getContent().add(t);

        return p;

    }
}