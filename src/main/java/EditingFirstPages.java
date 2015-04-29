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
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class EditingFirstPages {

    private    WordprocessingMLPackage doc;
    private  ObjectFactory factory;
    private int numGOST;
    private WordprocessingMLPackage newDoc;
    String type;
    //------------------------------
    String year, regYear = "\\d{4}";       // год выпуска
    String nPages, nPages2, regNPages = "Листов \\d+";     // количество станиц
    String letter, regLetter = "\"[ПЭТОАБ]{1}[12]?\"";     // литера
    String medium;            //вид носителя данных
    String docNumber,                              // обозначение документа
            regDocNumber = (numGOST == 19)? "[А-ЯA-Z]+.\\d+.\\d+-\\d{2}.( ){1}\\d{2}-?\\d*-(ЛУ){1}" :
                                           "(\\d+.){2}\\d{3}.[А-Я]{1,2}\\d?.?\\d*.?\\d*-?\\d*.?M?-(ЛУ){1}";
    String name;      // название документа
    int indexName;
    String subName;   // тип программы
    String company;
    boolean setType = false;
    int index_soglasovano_1 = -1, index_soglasovano_2 = -1;
    boolean utverjdau;
    List<P> soglasovano1;
    List<P> utvergdau;
    String[] soglasovano2;
    List<P> signatures = new ArrayList<>();
    int bound = -1;
    //------------------------------



    private void processDoc () {
        if (doc == null)
            return;
        List<P> documentParagraphes = DocBase.deleteEmptyPara(DocxMethods.createParagraphJAXBNodes(doc));
        setType = !name.toLowerCase().contains(type.toLowerCase());
        List<P>  firstPagePara = new ArrayList<>();
        int  i = find(documentParagraphes, regYear, "year", "ГОД НЕ УКАЗАН");
        if (i != -1) {
            firstPagePara = documentParagraphes.subList(0, i+2);
            find(firstPagePara, regNPages, "nPages", "");
            find(firstPagePara, regDocNumber, "docNumber", "НОМЕР ДОКУМЕНТА НЕ УКАЗАН ИЛИ УКАЗАН НЕ ПО ГОСТУ");
            find(firstPagePara, regLetter, "letter", "ЛИТЕРА НЕ УКАЗАНА");
            find(firstPagePara, "name");
            find(firstPagePara, "docNumber");
        }
        else
        {
            i = find(documentParagraphes, regLetter, "letter", "ЛИТЕРА НЕ УКАЗАНА");
            if (i != -1) {
                firstPagePara = documentParagraphes.subList(0, i);
                find(firstPagePara, regNPages, "nPages", "");
                find(firstPagePara, regDocNumber, "docNumber", "НОМЕР ДОКУМЕНТА НЕ УКАЗАН ИЛИ УКАЗАН НЕ ПО ГОСТУ");
                find(firstPagePara, regYear, "year", "ГОД НЕ УКАЗАН");
                find(firstPagePara, "name");
                find(firstPagePara, "docNumber");
            }
            else {
                find(documentParagraphes, regNPages, "nPages", "");
                find(documentParagraphes, regDocNumber, "docNumber", "НОМЕР ДОКУМЕНТА НЕ УКАЗАН ИЛИ УКАЗАН НЕ ПО ГОСТУ");
                find(documentParagraphes, "name");
                find(documentParagraphes, "docNumber");
            }
        }
        findSOGLAS_UTVERGD((firstPagePara.size() == 0)?firstPagePara:documentParagraphes);
        findBlocksInUtv_SOGLAS((firstPagePara.size() == 0)?firstPagePara:documentParagraphes);
        System.out.println(medium);
        System.out.println(nPages);
        System.out.println(name);
        System.out.println(subName);
        System.out.println(year);
        System.out.println(docNumber);
        System.out.println(setType);

 //       setFirstPage();
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
        System.out.println(e.ngrammPossibility("утверждено", "утвeржденo"));
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

        P[] pr0 = new P[0];
        if (!(company == null) && !company.equals("")) {
            pr0[0] = setP(company, "Arial", "1", "0", "0", 360, null, null);
            table.getContent().add(addRowWithMergedCells(false, null, pr0, null, 0, 9500, 0, 0));
        }
        int i = 0;
        //TODO обработать когда массив пустой
        P[] pr1 = null;
        if (soglasovano1!= null) {
            pr1 = new P[soglasovano1.size() + 1];
            pr1[i] = setP("СОГЛАСОВАНО", "Times New Roman", null, null, null, 240, null, null); i++;
            for (P p : soglasovano1) {
                pr1[i] = setP(DocBase.getText(p), "Times New Roman", null, null, null, 240, null, null); i++;
            }
            i = 0;
        }
        P[] pr2 = null;
        if (utvergdau!= null) {
            pr2 = new P[utvergdau.size() + 1];
            pr2[i] = setP("УТВЕРЖДАЮ", "Times New Roman", null, null, null, 240, null, null); i++;
            for (P p : utvergdau) {
                pr2[i] = setP(DocBase.getText(p), "Times New Roman", null, null, null, 240, null, null); i++;
            }
        }
        P[] pr = {setP("", "Times New Roman", null, null, null, 240, null, null) };
        table.getContent().add(addRowWithMergedCells(true, pr1,pr,pr2, 3000,3000, 3000, 1 ));

        P[] pr3 = new P[10];
        i = 0;
        pr3[i] = setP("","Times New Roman", "30", "0", "0", 240, null, null);
        i++;
        pr3[i] = setP(name, "Times New Roman",null, null, null, 240, null, null); i++;
        pr3[i] = setP("", "Arial","30", "0", "0", 240, null, null);i++;
        if (!(subName == null) && !subName.equals("")) { setP(subName, "Arial","30", "0", "0", 360, null, null); i++;}
        if (setType) { setP(type, "Arial", null, null, null, 360, null, null); i++;}
        pr3[i] = setP("Лист утверждения", "Arial", "1", "0", "0", 360, null, null);i++;
        pr3[i] = setP(docNumber, "Arial", "30", "0", "0", 360, null, null);i++;
        if (!(medium == null) && !medium.equals("")) {setP(medium, "Arial", "30", "0", "0", 360, null, null); i++;}
        pr3[i] = setP("", "Arial", "30", "0", "0", 360, null, null);i++;
        pr3[i] = setP(nPages, "Arial", "30", "0", "0", 360, null, null);
        table.getContent().add(addRowWithMergedCells(false, null, pr3, null, 0, 9500, 0, 2));
        P[] pr_ = {new P(), new P(), new P(), new P(),new P(), new P(), new P(), new P(),new P(), new P(), new P(), new P(),
                new P(), new P(), new P(), new P(),new P(), new P(), new P(), new P()};

        P[] pr4 = {setP("СОГЛАСОВАНО","Times New Roman", null, null, null, 240, null, null),
                setP("Главный", "Times New Roman", null, null, null, 240, null, null),
                setP("должность", "Times New Roman", null, null, null, 240, null, null),
                setP("(подпись) ________________", "Times New Roman", null, null, null, 240, null, null),
                setP("дата", "Times New Roman", null, null, null, 240, null, null)};
        P[] pr5 = {setP("Нормоконтролер", "Times New Roman", null, null, null, 240, null, null),
                setP("(подпись) ________________", "Times New Roman", null, null, null, 240, null, null),
                setP("дата", "Times New Roman", null, null, null, 240, null, null)};
        P[] pr6 = {setP(year, "Times New Roman", null, null, null, 240, null, null),
                setP(letter, "Times New Roman", null, null, null, 240, "RIGHT", null)};
        table.getContent().add(addRowWithMergedCells(false, pr4, pr_, pr5, 3000, 3000, 3000, 3));
        table.getContent().add(addRowWithMergedCells(false, null, pr6, null, 0, 4500, 0, 4));
        newDoc.getMainDocumentPart().getContent().add(table);
    }
    private  void setSecondPage() {
        //   doc.getMainDocumentPart().getContent().add(setP("", "Times New Roman", null, null, null, 480, null, null));
        newDoc.getMainDocumentPart().getContent().add(setP("", "Times New Roman", null, null, null, 480, null, null));
        newDoc.getMainDocumentPart().getContent().add(setP("", "Times New Roman", null, null, null, 480, null, null));
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

//        table.getTblPr().setTblLook(tblLook);

        P[] pr1 = {setP("УТВЕРЖДЕНО", "Times New Roman", null, null, null, 480, null, null),
                setP(docNumber, "Courier New", null, null, null, 240, "LEFT", "20"),
                setP("", "Times New Roman", null, null, null, 240, null, null),setP("", "Times New Roman", null, null, null, 240, null, null),
                setP("", "Times New Roman", null, null, null, 240, null, null),setP("", "Times New Roman", null, null, null, 240, null, null)};

        P[] pr = {setP("", "Times New Roman", null, null, null, 240, null, null) };
        table.getContent().add(addRowWithMergedCells(true, pr1,pr,pr, 1500,1500, 1500, 1 ));

//        P[] pr3 = new P[9];
//        int i = 0;
//        pr3[i] = setP("","Times New Roman", "30", "0", "0", 240, null, null);
//        i++;
//        pr3[i] = setP(name, "Times New Roman",null, null, null, 240, null, null); i++;
//        pr3[i] = setP("", "Arial","30", "0", "0", 240, null, null);i++;
//        if (!(subName == null) && !subName.equals("")) { setP(subName, "Arial","30", "0", "0", 360, null, null); i++;}
//        if (setType) { setP(type, "Arial", null, null, null, 360, null, null); i++;}
//        pr3[i] = setP(docNumber, "Arial", "30", "0", "0", 360, null, null);i++;
//        if (!(medium == null) && !medium.equals("")) {setP(medium, "Arial", "30", "0", "0", 360, null, null); i++;}
//        pr3[i] = setP("", "Arial", "30", "0", "0", 360, null, null);i++;
//        pr3[i] = setP(nPages2, "Arial", "30", "0", "0", 360, null, null);
//        table.getContent().add(addRowWithMergedCells(false, null, pr3, null, 0, 4500, 0, 2));
//        P[] pr_ = {new P(), new P(), new P(), new P(),new P(), new P(), new P(), new P(),new P(), new P(), new P(), new P(),
//                new P(), new P(), new P(), new P(),new P(), new P(), new P(), new P()};
//        table.getContent().add(addRowWithMergedCells(false, pr, pr_, pr, 1500, 1500, 1500, 3));
//        P[] pr5 = {setP(year, "Times New Roman", null, null, null, 240, null, null),
//                setP(letter, "Times New Roman", null, null, null, 240, "RIGHT", null)};
       // table.getContent().add(addRowWithMergedCells(false, null, pr5, null, 0, 4500, 0, 4));
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
            addMergedColumn(row, image, 0, 500);
            addTableCell(row, ps, width1, 1);
            addTableCell(row, ps2, width2, 2);
            addTableCell(row, ps3, width3, 3);
        }

        return row;
    }

    private  void addMergedColumn(Tr row, boolean image, int grid, int width) {
        if (!image) {
            addMergedCell(row, image, null, grid, width);
        } else {
            addMergedCell(row, image, "restart", grid, width);
        }
    }

    private  void addMergedCell(Tr row, boolean image, String vMergeVal, int grid, int width){
        Tc tableCell = factory.createTc();
        TcPr tableCellProperties = new TcPr();

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
        if(!image) {
            tableCell.getContent().add(newImage());
        }

        row.getContent().add(tableCell);
    }

    private  void addTableCell(Tr tr, P[] content, int width, int grid) {
        Tc tc1 = factory.createTc();
        if (content == null)
            tc1.getContent().add(new P());
        else
            for (P p : content) {
                if (p!= null)
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

    private  P setP (String text, String font, String style, String ilvl, String numId, int spacing, String align, String size) {
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
        R run = factory.createR();
        p.getContent().add(run);
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
            System.out.println("File is too large!!");
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
            imagePart = BinaryPartAbstractImage.createImagePart(doc, bytes);
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

    private int find (List<P> para, String reg, String who, String notFound) {
        Pattern pattern = Pattern.compile(reg);
        Matcher m = pattern.matcher(DocBase.getText(para.get(0)));
        int i = 1;
        String q = "";
        while(i != para.size() && !m.matches()) {
            q = DocBase.getText(para.get(i));
            m = pattern.matcher(q);
            if (ngrammPossibility(q, name) > 0.7)
                indexName = i;
            ++i;
        }
        if (!m.matches())
            q = notFound;
        switch (who) {
            case ("year"): year = q;break;
            case ("nPages"): nPages = q;break;
            case ("letter"): letter = q;break;
            case ("docNumber"): docNumber = q; break;
         }
        return (m.matches())? i - 1 : -1;
    }

    private void find(List<P> para, String afterWhat) {
        int i = 1;
        String q ;
        while(i != para.size()) {
            q = DocBase.getText(para.get(i));
            if (afterWhat.equals("name"))
                if (q.toLowerCase().equals(name.toLowerCase())) {
                    indexName = i;
                    break;
                }
            if (afterWhat.equals("docNumber"))
            if (q.toLowerCase().equals(docNumber.toLowerCase()))
                break;
            ++i;
        }
        i++;
        if (i < para.size() ) {
            q = DocBase.getText(para.get(i));
            if (q.toLowerCase().equals("лист утверждения") &&
                    q.toLowerCase().equals(type.toLowerCase()) &&
                    q.toLowerCase().equals(nPages.toLowerCase()) &&
                    q.toLowerCase().equals(docNumber.toLowerCase()) &&
                    q.trim().equals(""))
                if (afterWhat.equals("name"))
                     subName = q;
                else
                    medium = q;
        }
    }

    private void findSOGLAS_UTVERGD (List<P> para) {
        String s;
        for (int i = 0; i< para.size(); i++) {
            s = DocBase.getText(para.get(i)).toLowerCase();
            if (s.toLowerCase().equals("cогласовано") || ngrammPossibility("согласовано", s) >= 0.5) {
                if (i < indexName)
                    index_soglasovano_1 = i;
                else
                    index_soglasovano_2 = i;
            }

            if (s.toLowerCase().equals("утверждаю")|| ngrammPossibility("утверждаю", s) >= 0.5) {
                utverjdau = true;
            }
        }
        Pattern p = Pattern.compile("[а-яА-ЯA-Za-z ]+]");
        Matcher m;
        for (int i = 0; i < index_soglasovano_1; i++) {
            s = DocBase.getText(para.get(i));
            m = p.matcher(s);
            if (m.matches() && !s.toLowerCase().equals(name.toLowerCase()) &&
            !s.toLowerCase().equals(subName.toLowerCase()) &&
            !s.toLowerCase().equals(medium.toLowerCase())) {
                company = s;
            }

        }
    }

    private void findBlocksInUtv_SOGLAS (List<P> para) {
        Pattern p1 = Pattern.compile("[А-Яа-я ]+");
        Pattern p2 = Pattern.compile("([А-Я]{1}[а-я ]+ ?)+ ?([А-Я]{1}.?)+");
        Pattern p3 = Pattern.compile("([0-9]{0,4}[./\\-]?){3}");
            if (index_soglasovano_1!= -1 && para.size() - 2 > index_soglasovano_1 ) {
                String s1, s2 = "";
                int firstLine = index_soglasovano_1 +1;
                int lastLine = 0;
                for (int i = index_soglasovano_1 + 1; i < para.size() - 1; i++) {
                    s1 = DocBase.getText(para.get(i));
                    s2 = DocBase.getText(para.get(i+1));
                    if (p1.matcher(s2).matches() || s1.toLowerCase().equals("утверждаю")) {
                        if (p2.matcher(s1).matches() || p3.matcher(s1).matches()
                                || s1.toLowerCase().contains("дата") || s1.toLowerCase().contains("м.п.")) {
                            lastLine = i;
                            soglasovano1 = para.subList(firstLine, lastLine);
                            break;
                        }
                    }
                }

                if (s2.toLowerCase().equals("утверждаю")) {
                    firstLine = lastLine +2;
                    for (int i = firstLine; i < para.size() - 1; i++) {
                        s1 = DocBase.getText(para.get(i));
                        if (p2.matcher(s1).matches() || p3.matcher(s1).matches() || s1.toLowerCase().contains("дата")
                                || s1.toLowerCase().contains("м.п.")) {
                            lastLine = i;
                            utvergdau = para.subList(firstLine, lastLine);
                            break;
                        }
                    }
                }
        }
    }

    private void findOtherBlocks (List<P> para) {
        String s;
        List<String> signatures = new ArrayList<>();
        if (index_soglasovano_2!=-1)
            for (int i = index_soglasovano_2; i < para.size(); i++) {
                s = DocBase.getText(para.get(i));
                if (!s.equals(name) & !s.equals(subName) & !s.equals(docNumber) & !s.equals(medium)
                        &!s.equals(nPages) &!s.equals(nPages2) &!s.equals(year) & !s.equals(letter)) {
                    signatures.add(s);
                }
            }
        soglasovano2 = (String[])signatures.toArray();
    }

    private void checkForExistingInTable (List<P> para) {
     WordprocessingMLPackage uml = new WordprocessingMLPackage();
     uml.getMainDocumentPart().getContent().addAll(para);
     try {
         uml.save(new File("temp.docx"));
         uml = DocxMethods.getTemplate("temp.docx");
         List<Object> tables = DocxMethods.createTableJAXBNodes(uml);
         boolean found;
         for (Object o : tables) {
             Tbl table = (Tbl) o;
             found = findInTableParagraphes(table);
             if (found) {
                 break;
             }
         }
     } catch (Docx4JException e) {
         e.printStackTrace();
     } catch (FileNotFoundException e) {
         e.printStackTrace();
     }
 }

    private boolean findInTableParagraphes (Tbl table) {
        List<Object> rows = table.getContent();
        boolean found = false;
        int lastIndex = -1;
        for (Object o : rows) {
            if (o instanceof Tr) {
                List<Object> cells = ((Tr) o).getContent();
                for (int j = cells.size() -1; j >= 0; j--) {
                    Object cell = cells.get(j);
                    if (cell instanceof Tc) {
                        List<Object> para = ((Tc) cell).getContent();
                        for (int i = 0; i < signatures.size(); i++) {
                            if (signatures.get(i).equals(para.get(i))) {
                                found = true;
                            }
                            else {
                                if (found & lastIndex > 1)
                                    bound = i;
                                return true;
                            }
                        }
                    }
                }
            }
        }
        return false;
    }

    private boolean checkForRIGHTAlign () {
        JcEnumeration jc;
        for (int i = 0; i < signatures.size(); i++) {
            P signature = signatures.get(i);
            jc = DocBase.getAlign(signature);
            if (jc != null && jc == JcEnumeration.RIGHT) {
                bound = i;
                return true;
            }
        }
        return false;
    }

    private double ngrammPossibility (String actual, String checked) {
        String[] actualGramm  = new String[actual.length() - 2];
        String[] checkedGramm  = new String[checked.length() - 2];
        int index = 0;
        for (int i = 0; i < checked.length(); i++) {
            if (i ==checked.length() - 3) {
                checkedGramm[index] = checked.substring(i, i+3);
                break;
            }
            else checkedGramm[index] = checked.substring(i, i+3);
            index++;
        }
        index = 0;
        for (int i = 0; i < actual.length(); i++) {
            if (i ==actual.length() - 3) {
                actualGramm[index] = actual.substring(i, i+3);
                break;
            }
            else actualGramm[index] = actual.substring(i, i+3);
            index++;
        }

        double coincidence = 0.0;
        index = 0;
        int max = (checkedGramm.length>actualGramm.length)? actualGramm.length :checkedGramm.length;
        while (index != max) {
            if (checkedGramm[index].equals(actualGramm[index]))
                coincidence++;
            ++index;
        }
        return coincidence/max;
    }
}
