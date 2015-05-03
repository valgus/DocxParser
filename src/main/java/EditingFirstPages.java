import org.docx4j.dml.wordprocessingDrawing.Inline;
import org.docx4j.jaxb.Context;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.exceptions.InvalidFormatException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.BinaryPartAbstractImage;
import org.docx4j.wml.*;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.math.BigInteger;
import java.util.List;

public class EditingFirstPages {

    private    WordprocessingMLPackage doc;
    private  ObjectFactory factory;
    private int numGOST;
    private WordprocessingMLPackage newDoc;
    String type;
    //------------------------------
    String year   ;  // год выпуска
    String nPages;    // количество станиц
    String letter;     // литера
    String medium;            //вид носителя данных
    String docNumber;
    String name;      // название документа
    String subName;   // тип программы
    String company;
    boolean setType = false;
    String agreement;
    String approve;
    List<P> remained;
    int bound = -1;
    //------------------------------



    private void processDoc () throws Exception {
        if (doc == null)
            throw new Exception("Doc is empty");
        List<P> docPara = DocBase.deleteEmptyPara(DocxMethods.createParagraphJAXBNodes(doc));
        String s;
        int yearIndex = -1, letterIndex = -1;
        for (int i = 0; i < docPara.size(); i++ ) {
            s = DocBase.getText(docPara.get(i));
            if (s.matches("[0-9]{4}")) {
                year = s;
                yearIndex = i;
            }
            if (s.matches("\"[ПЭТОАБИ]{1}[12]?\"")) {
                letter = s;
                letterIndex = i;
                break;
            }
        }
        if (year == null && letter == null)
            throw new Exception("year or letter must be set");
        Attempt a = new Attempt(docPara.subList(0, (yearIndex!= -1)?yearIndex - 1 : letterIndex - 1),
                type, numGOST, name);
        a.maa();
        nPages = a.getPageNumber();
        medium = a.getMedium();
        docNumber = a.getDocNumber();
        subName = a.getSubName();
        company = a.getNameOfCompany();
        setType = a.isSetType();
        agreement = a.getAgreement();
        approve = a.getApprove();
        remained = a.getRemained();

        if (!checkForRIGHTAlign()) {
            findInTableParagraphs();
        }
        setFirstPage();
        setSecondPage();
    }

    private void setDoc(WordprocessingMLPackage doc, String typeOfDoc, int numGOST, String name){
       this.doc = doc;
       this.type = typeOfDoc;
       this.numGOST = numGOST;
       this.name = name;
    }

    public static void main(String[] args) throws Exception{
        EditingFirstPages e = new EditingFirstPages();
        e.factory = Context.getWmlObjectFactory();
        e.doc = DocxMethods.getTemplate("docx/2.docx");
        e.setDoc(DocxMethods.getTemplate("docx/2_1.docx"), "Руководство оператора", 19,
                "АНАЛИЗАТОР ПОКЕРНЫХ ИГР НА ОСНОВЕ МНОГОСЛОЙНОГО ПЕРСЕПТРОНА");
        try {
            e.newDoc = WordprocessingMLPackage.createPackage();
        } catch (InvalidFormatException ex) {
            ex.printStackTrace();
        }
        e.processDoc();
        try {
            e.newDoc.save(new File("3.docx"));
        } catch (Docx4JException ex) {
            ex.printStackTrace();
        }
    }

    private  void setFirstPage() {

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
        tblGridCol2.setW(BigInteger.valueOf(3000));
        TblGridCol tblGridCol3 = new TblGridCol();
        tblGridCol3.setW(BigInteger.valueOf(3000));
        TblGridCol tblGridCol4 = new TblGridCol();
        tblGridCol4.setW(BigInteger.valueOf(3000));
        table.getTblGrid().getGridCol().add(tblGridCol1);
        table.getTblGrid().getGridCol().add(tblGridCol2);
        table.getTblGrid().getGridCol().add(tblGridCol3);
        table.getTblGrid().getGridCol().add(tblGridCol4);

        P[] pr0 = new P[1];
        if (!(company == null) && !company.equals("")) {
            pr0[0] = setP(company, "Arial", "1", "0", "0", 360, null, null, false);
            table.getContent().add(addRowWithMergedCells(false, null, pr0, null, 0, 9500, 0, 0));
        }
        int i = 0;
        P[] pr1 = null;
        if (agreement == null) {
            pr1 = new P[1];
            pr1[0] = setP(agreement, "Times New Roman", null, null, null, 240, null, null, false);
        }
        P[] pr2 = null;
        if (approve == null) {
            pr2 = new P[1];
            pr2[0] = setP(approve, "Times New Roman", null, null, null, 240, null, null, false);
        }
        P[] pr = {setP("", "Times New Roman", null, null, null, 240, null, null, false) };
        table.getContent().add(addRowWithMergedCells(true, pr1,pr,pr2, 3000,3000, 3000, 1 ));

        P[] pr3 = new P[10];
        pr3[i] = setP("","Times New Roman", "30", "0", "0", 240, null, null, false);i++;
        pr3[i] = setP(name, "Times New Roman",null, null, null, 240, null, null, false); i++;
        pr3[i] = setP("", "Arial","30", "0", "0", 240, null, null, false);i++;
        if (!(subName == null) && !subName.equals("")) { setP(subName, "Arial","30", "0", "0", 360, null, null, false); i++;}
        if (setType) { setP(type, "Arial", null, null, null, 360, null, null, false); i++;}
        pr3[i] = setP("Лист утверждения", "Arial", "1", "0", "0", 360, null, null, false);i++;
        if (docNumber.contains("{wrong}")) docNumber = docNumber.replace("{wrong}", "");
        pr3[i] = setP(docNumber, "Arial", "30", "0", "0", 360, null, null, docNumber.contains("{wrong}"));i++;
        if (!(medium == null) && !medium.equals("")) {setP(medium, "Arial", "30", "0", "0", 360, null, null, false); i++;}
        pr3[i] = setP("", "Arial", "30", "0", "0", 360, null, null, false);i++;
        if (nPages.contains("{wrong}")) {nPages = nPages.replace("{wrong}", "");
            pr3[i] = setP(nPages, "Arial", "30", "0", "0", 360, null, null, true);}
        table.getContent().add(addRowWithMergedCells(false, null, pr3, null, 0, 9500, 0, 2));
        P[] pr_ = {new P(), new P(), new P(), new P(),new P(), new P(), new P(),
                new P(),new P(), new P(), new P(), new P()};

        P[] pr4;
        P[] pr5;
        List<String> remainStrings = DocBase.changeToString(remained);
        if (bound != -1) {
            pr4 = new P[bound+1];
            pr4[0] = setP("СОГЛАСОВАНО","Times New Roman", null, null, null, 240, null, null, false);
            for (int k = 1; k < remainStrings.size() - bound - 1; k ++ ) {
                pr4[k] = setP(remainStrings.get(k), "Times New Roman", null, null, null, 240, null, null, false);
            }
            pr5 = new P[remainStrings.size() - bound];
            for (int k = 0; k < remainStrings.size() - bound; k ++ ) {
                pr5[k] = setP(remainStrings.get(k), "Times New Roman", null, null, null, 240, null, null, false);
            }
        }
        else {
            pr4 = null;
            pr5 = new P[remainStrings.size()];
            for (int k = 0; k < remainStrings.size(); k ++ ) {
                pr5[k] = setP(remainStrings.get(k), "Times New Roman", null, null, null, 240, null, null, false);
            }
        }
        P[] pr6 = {setP(year, "Times New Roman", null, null, null, 240, null, null, false),
                setP(letter, "Times New Roman", null, null, null, 240, "RIGHT", null, false)};
        table.getContent().add(addRowWithMergedCells(false, pr4, pr_, pr5, 3000, 3000, 3000, 3));
        table.getContent().add(addRowWithMergedCells(false, null, pr6, null, 0, 4500, 0, 4));
        newDoc.getMainDocumentPart().getContent().add(table);
        Br objBr = new Br();
        objBr.setType(STBrType.PAGE);
        P p = factory.createP();
        R r = factory.createR();
        r.getContent().add(objBr);
        p.getContent().add(r);
        newDoc.getMainDocumentPart().getContent().add(p);
    }
    private  void setSecondPage() {
        //   doc.getMainDocumentPart().getContent().add(setP("", "Times New Roman", null, null, null, 480, null, null));
        newDoc.getMainDocumentPart().getContent().add(setP("", "Times New Roman", null, null, null, 480, null, null, false));
        newDoc.getMainDocumentPart().getContent().add(setP("", "Times New Roman", null, null, null, 480, null, null, false));
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

        P[] pr1 = {setP("УТВЕРЖДЕНО", "Times New Roman", null, null, null, 480, null, null, false),
                setP(docNumber, "Courier New", null, null, null, 240, "LEFT", "20", false),
                };

        P[] pr = {setP("", "Times New Roman", null, null, null, 240, null, null, false),
                setP("", "Times New Roman", null, null, null, 240, null, null, false),
                setP("", "Times New Roman", null, null, null, 240, null, null, false) };
        table.getContent().add(addRowWithMergedCells(true, pr1,pr,pr, 1500,1500, 1500, 1 ));

        P[] pr3 = new P[9];
        int i = 0;
        pr3[i] = setP("","Times New Roman", "30", "0", "0", 240, null, null, false);
        i++;
        pr3[i] = setP(name, "Times New Roman",null, null, null, 240, null, null, false); i++;
        pr3[i] = setP("", "Arial","30", "0", "0", 240, null, null, false);i++;
        if (!(subName == null) && !subName.equals("")) {pr3[i] =  setP(subName, "Arial","30", "0", "0", 360, null, null, false); i++;}
        if (setType) { pr3[i] = setP(type, "Arial", null, null, null, 360, null, null, false); i++;}
        pr3[i] = setP(docNumber, "Arial", "30", "0", "0", 360, null, null, false);i++;
        if (!(medium == null) && !medium.equals("")) {pr3[i] = setP(medium, "Arial", "30", "0", "0", 360, null, null, false); i++;}
        pr3[i] = setP("", "Arial", "30", "0", "0", 360, null, null, false);i++;
        pr3[i] = setP("Number is 2", "Arial", "30", "0", "0", 360, null, null, false);
        table.getContent().add(addRowWithMergedCells(false, null, pr3, null, 0, 4500, 0, 2));
        P[] pr_ = {new P(), new P(), new P(), new P(),new P(), new P(), new P(), new P(),new P(), new P(), new P(), new P()};
        table.getContent().add(addRowWithMergedCells(false, pr, pr_, pr, 1500, 1500, 1500, 3));
        P[] pr5 = {setP(year, "Times New Roman", null, null, null, 240, null, null, false),
                setP(letter, "Times New Roman", null, null, null, 240, "RIGHT", null, false)};
        table.getContent().add(addRowWithMergedCells(false, null, pr5, null, 0, 4500, 0, 4));
        newDoc.getMainDocumentPart().getContent().add(table);
    }

    private  Tr addRowWithMergedCells(boolean image, P[] ps, P[] ps2, P[] ps3, int width1, int width2 , int width3, int num) {

        Tr row = factory.createTr();
        row.setTrPr(factory.createTrPr());

        if (num == 0 || num == 2 || num == 4) {
            addMergedColumn(row, image, -1, 0);
            addTableCell(row, ps2, width2, 6);
        }
        if (num == 1 || num == 3) {
            int width = (num == 1) ? 500 : 0;
            addMergedColumn(row, image, 0, width);
            addTableCell(row, ps, width1, 1);
            addTableCell(row, ps2, width2, 2);
            addTableCell(row, ps3, width3, 3);
        }

        return row;
    }

    private  void addMergedColumn(Tr row, boolean image, int grid, int width) {
        if (!image) {
            addMergedCell(row, image, "continue", grid, width);
        } else {
            addMergedCell(row, image, "restart", grid, width);
        }
    }

    private  void addMergedCell(Tr row, boolean image, String vMergeVal, int grid, int width){
        Tc tableCell = factory.createTc();
        TcPr tableCellProperties = new TcPr();

        TcPrInner.VMerge merge = new TcPrInner.VMerge();
        merge.setVal(vMergeVal);
        tableCellProperties.setVMerge(merge);

      //  if (image) {
            CTVerticalJc ctVerticalJc = new CTVerticalJc();
            ctVerticalJc.setVal(STVerticalJc.BOTTOM);
            tableCellProperties.setVAlign(ctVerticalJc);
            TblWidth tableWidth = factory.createTblWidth();
            tableWidth.setType("dxa");
            tableWidth.setW(BigInteger.valueOf(width));
            tableCell.setTcPr(tableCellProperties);
            tableCell.getTcPr().setTcW(tableWidth);


//         TcPrInner.GridSpan gridSpan = new TcPrInner.GridSpan();
//            gridSpan.setVal(BigInteger.valueOf(grid));
//            tableCellProperties.setGridSpan(gridSpan);

        if(image) {
            tableCell.getContent().add(newImage());
        }
        else {tableCell.getContent().add(new P());}

        row.getContent().add(tableCell);
    }

    private  void addTableCell(Tr tr, P[] content, int width, int grid) {
        Tc tc1 = factory.createTc();
        if (content == null)
            tc1.getContent().add(new P());
        else
            for (P p : content) {
                if (p!=null)
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

    private  P setP (String text, String font, String style, String ilvl, String numId, int spacing, String align, String size, boolean highlight) {
        P p = new P();
        DocBase.setText(p, text, false);
        p.setPPr(new PPr());
        DocBase.setStyle(p, size, font, style, ilvl, numId, spacing);
        PPrBase.Spacing space = new PPrBase.Spacing();
        Jc jc = factory.createJc();
        if (align!= null && align.equals("RIGHT"))
            jc.setVal(JcEnumeration.RIGHT);
        else if (align!= null && align.equals("LEFT"))
            jc.setVal(JcEnumeration.LEFT);
        else
            jc.setVal(JcEnumeration.CENTER);
        p.getPPr().setJc(jc);
        p.getPPr().setSpacing(space);
        if (highlight)
            DocBase.setHighlight(p, "yellow");
        return p;
    }

    private  org.docx4j.wml.P newImage() {

        File file = new File("resource/table.gif" );
        java.io.InputStream is = null;
        try {
            is = new java.io.FileInputStream(file );
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        long length = file.length();
        if (length > Integer.MAX_VALUE) {
            System.out.println("File too large!!");
        }
        byte[] bytes = new byte[(int)length];
        int offset = 0;
        int numRead;
        try {
            while (offset < bytes.length
                    && (numRead=is.read(bytes, offset, bytes.length-offset)) >= 0) {
                offset += numRead;
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
        if (offset < bytes.length) {
            System.out.println("Could not completely read file "+file.getName());
        }
        try {
            is.close();
        } catch (IOException e) {
            e.printStackTrace();
        }

        String filenameHint = null;
        String altText = null;
        int id1 = 0;
        int id2 = 1;
        BinaryPartAbstractImage imagePart = null;
        try {
            imagePart = BinaryPartAbstractImage.createImagePart(newDoc, bytes);
        } catch (Exception e) {
            e.printStackTrace();
        }

        Inline inline = null;
        try {
            inline = imagePart.createImageInline( filenameHint, altText,
                    id1, id2, true);
        } catch (Exception e) {
            e.printStackTrace();
        }
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


    private boolean checkForRIGHTAlign () {
        JcEnumeration jc;
        for (int i = 0; i < remained.size(); i++) {
            jc = DocBase.getAlign(remained.get(i));
            if (jc != null && jc == JcEnumeration.RIGHT) {
                bound = i;
                return true;
            }
        }
        return false;
    }

    private void findInTableParagraphs() {
        WordprocessingMLPackage uml;
        try {
            doc.save(new File("temp.docx"));
            uml = DocxMethods.getTemplate("temp.docx");
            List<Object> tables = DocxMethods.createTableJAXBNodes(uml);
            for (Object o : tables) {
                if (o instanceof Tbl) {
                    Tbl table = (Tbl) o;
                    List<Object> rows = table.getContent();
                    boolean found = false;
                    int lastIndex = -1;
                    for (Object o2 : rows) {
                        if (o2 instanceof Tr) {
                            List<Object> cells = ((Tr) o2).getContent();
                            for (int j = cells.size() -1; j >= 0; j--) {
                                Object cell = cells.get(j);
                                if (cell instanceof Tc) {
                                    List<Object> para = ((Tc) cell).getContent();
                                    for (int i = 0; i < remained.size(); i++) {
                                        if (remained.get(i).equals(para.get(i))) {
                                            found = true;
                                            lastIndex = i;
                                        }
                                        else {
                                            if (found)
                                                bound = lastIndex + 1;
                                            return;
                                        }
                                    }
                                }
                            }
                        }
                    }
                    return;
                }
            }
        } catch (Docx4JException e) {
            e.printStackTrace();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }

    }

}
