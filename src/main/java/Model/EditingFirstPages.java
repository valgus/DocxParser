package Model;

import org.apache.commons.lang.StringUtils;
import org.docx4j.convert.out.common.preprocess.CoverPageSectPrMover;
import org.docx4j.jaxb.Context;
import org.docx4j.model.structure.PageDimensions;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.*;

import java.io.File;
import java.io.FileNotFoundException;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.List;

public class EditingFirstPages {

    private    WordprocessingMLPackage doc;
    private Tbl tbl;
    private  ObjectFactory factory;
    private int numGOST;
    String type;
    //------------------------------
    private  String year   ;  // год выпуска
    private  String nPages;    // количество станиц
    private  String letter;     // литера
    private  String changeString;
    private  String medium;            //вид носителя данных
    private  String docNumber;
    private  String name;      // название документа
    private  String subName;   // тип программы
    private  String company;
    private  String albom;
    private  String agreement;
    private  String approve;
    private  List<P> remained;
    private String nPages2;
    private  int bound = -1;
    //------------------------------
    private int pageWidth = new PageDimensions().getWritableWidthTwips();
    private ProcessDocument process;

    public boolean sendIndfo(String s) {
        return process.sendInfo(s);
    }

    public EditingFirstPages(WordprocessingMLPackage doc, String typeOfDoc, int numGOST, String name,
                             ProcessDocument process){
        factory = Context.getWmlObjectFactory();
        this.doc = doc;
        this.type = typeOfDoc;
        this.numGOST = numGOST;
        this.name = name;
        this.process = process;
        nPages2 = "Листов ___";
        nPages = "";
        letter = "Литера";
        medium = "(вид носителя данных)";
        docNumber = "номер документа";
        subName = "Наименование документа";
        company = "Наименование министерства или ведомства, " +
                "в систему которого входит организация, разработавшая данный документ";
        agreement = "СОГЛАСОВАНО \n должность \n (подпись) \t ФИО \n дата";
        approve = "УТВЕРЖДАЮ \n должность \n (подпись) \t ФИО \n дата";
    }

    public WordprocessingMLPackage process()throws Exception{
        if (doc == null) {
            sendIndfo("Docx is empty");
            return null;
        }
        P first = null;
        CoverPageSectPrMover.process(doc);
        List<P> docPara = DocBase.deleteEmptyPara(DocxMethods.createParagraphJAXBNodes(doc));
        String s;
        int yearIndex = -1, letterIndex = -1;
        for (int i = 0; i < docPara.size(); i++ ) {
            s = DocBase.getText(docPara.get(i));
            if (s.replace("г.", "").matches("[0-9]{4}")) {
                year = s;
                yearIndex = i;
                first = docPara.get(i);
                if (letterIndex != -1)
                    break;
            }
            String k = s.substring((s.length()>=4)?s.length()-4:0);
            if (k.matches("[«“\"]?[ПЭТОАБИ]{1}[12]?[»”\"]?")) {
                letter = s;
                letterIndex = i;
                first = docPara.get(i);
                if (yearIndex != -1)
                    break;
            }
            if (yearIndex!=-1 && i - yearIndex > 10 ||
                    letterIndex!=-1 && i - letterIndex > 10)
                break;
        }
        if (yearIndex == -1 && letterIndex == -1) {
            boolean is = sendIndfo("year or letter must be set");
            if (is) {
                year = "ДАТА";
                setFirstPage();
                setSecondPage();
                return doc;
            }
            else {
                return null;
            }
        }
        String res = "";
        if (yearIndex != -1 && letterIndex != -1)
            if (letterIndex - yearIndex == 2)
                changeString = DocBase.getText(docPara.get(letterIndex - yearIndex));
        if (yearIndex!= 0 &&  yearIndex!= -1  || letterIndex != -1) {
            ProcessorOfTitles a = new ProcessorOfTitles(docPara.subList(0, (yearIndex!= -1)?yearIndex : letterIndex),
                    type, numGOST, name, process);
             res = a.findMainElements();
            if (res == null) {
                if (!a.getPageNumber().equals("")) nPages = a.getPageNumber();
                if (!a.getMedium().equals("")) medium = a.getMedium();
                if (!a.getDocNumber().equals("")) docNumber = a.getDocNumber();
                if (docNumber.equals("")) {
                    docNumber = "номер документа{wrong}";
                }
                else {
                    if (!docNumber.toLowerCase().contains("-лу"))
                        docNumber+="-ЛУ";
                }
                if (!a.getSubName().equals("")) subName = a.getSubName();
                if (!a.getNameOfCompany().equals("")) company = a.getNameOfCompany();
                if (!a.getAgreement().equals("")) agreement = a.getAgreement();
                if (!a.getApprove().equals("")) approve = a.getApprove();
                remained = a.getRemained();
                if (!a.getAlbom().equals("")) albom = a.getAlbom();

                if (!checkForRIGHTAlign()) {
                    findInTableParagraphs();
                }
            }

            int year2Index = -1;
            int page2index = -1;
            int nLetter = -1;
            int temp = ((letterIndex != -1) ? letterIndex : yearIndex  ) + 1;
            P last = null;
            for (int i = temp; i < ((4*temp <docPara.size())?4*temp:docPara.size()); i++) {
                s = DocBase.getText(docPara.get(i));
                if (s.equals(year)) {
                    year2Index = i;
                    last = docPara.get(i);
                }
                if ((s.toLowerCase().contains("листов")||s.toLowerCase().contains("лист")) && s.length() < 15) {
                    page2index = i;
                    nPages2 = s;
                    last = docPara.get(i);
                }

                if (s.toLowerCase().contains(letter.trim())||
                        s.contains("\"") && s.replaceAll("\"", "").trim().matches("[ПЭТОАБИ]{1}[12]?")) {
                    nLetter = i;
                    last = docPara.get(i);
                    break;
                }
            }

            int indexToFind = (nLetter == -1)?(year2Index == -1)? (page2index == -1)?temp+1 :page2index : year2Index
                    :nLetter;

            boolean[] inTable = getParagraphesFromTbl(first, last);
            if (inTable[0] == true)
                indexToFind-=DocxMethods.getIndexOfParagraph(doc.getMainDocumentPart(), first);
            if (inTable[1] == false)  {

                int k = 0;
                int l = 0;
                if (DocBase.getText((P)doc.getMainDocumentPart().getContent().get(indexToFind+1)).contains("\"")
                        ||DocBase.getText((P)doc.getMainDocumentPart().getContent().get(indexToFind+1)).contains("«") ||
                        DocBase.getText((P)doc.getMainDocumentPart().getContent().get(indexToFind+1)).contains("“")) {
                    indexToFind+=2;
                }
                else
                    indexToFind+=1;
                while (k!=indexToFind) {
                    if (!StringUtils.deleteWhitespace(DocBase.getText((P)doc.getMainDocumentPart().getContent().get(l))).
                            equals(""))
                        k++;
                    DocBase.setHighlight((P)doc.getMainDocumentPart().getContent().get(l), "blue");
                    l++;
                }
            }
            while (true) {
                if (doc.getMainDocumentPart().getContent().get(indexToFind+1) instanceof P) {
                    P p = (P)doc.getMainDocumentPart().getContent().get(0);
                    s = DocBase.getText(p);
                    if (s.isEmpty() || s.matches("[ ]*")) {
                        doc.getMainDocumentPart().getContent().remove(doc.getMainDocumentPart().getContent().get(0));
                    }
                    else
                        break;
                }
                else
                    break;
            }
        }
        if (res.equals("Создать пустой лист утверждения")) setFirstPage();
        setSecondPage();

   //     System.out.println(doc.getMainDocumentPart().getContent().size());

        return doc;
    }

    private  void setFirstPage() {


        CTVerticalJc ctVerticalJc = new CTVerticalJc();
        ctVerticalJc.setVal(STVerticalJc.CENTER);
        Tbl table = factory.createTbl();
        table.setTblPr(new TblPr());
        TblWidth width = new TblWidth();
        width.setType("pct");
        width.setW(BigInteger.valueOf((int)(0.5*pageWidth)));
        table.getTblPr().setTblW(width);
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

        P[] pr0 = new P[1];
        if (!(company == null) && !company.equals("")) {
            pr0[0] = setP(company.toUpperCase(), "Arial", null, -1, -1, 360, "CENTER", null, false, false);
            table.getContent().add(addRowWithMergedCells(false, null, pr0, null, 0, (int)(pageWidth*0.9), 0, 0));
        }
        int i = 0;
        P[] pr1 = null;
        if (!(agreement == null) && !agreement.equals("")) {
            pr1 = new P[1];
            pr1[0] = setP(agreement, "Times New Roman", null, -1, -1, 240, "CENTER", null, false, false);
        }
        P[] pr2 = null;
        if (!(approve == null) && !approve.equals("")) {
            pr2 = new P[1];
            pr2[0] = setP(approve, "Times New Roman", null, -1, -1, 240, "CENTER", null, false, false);
        }
        P[] pr = {setP("", "Times New Roman", null, -1, -1, 240, null, null, false, false) };
        table.getContent().add(addRowWithMergedCells(true, pr1,pr,pr2, (int)(0.3*pageWidth) ,(int)(0.3*pageWidth),
                (int)(0.3*pageWidth), 1 ));

        P[] pr3 = new P[10];
        pr3[i] = setP("","Times New Roman", null, -1, -1, 240, null, null, false, false);i++;
        pr3[i] = setP(name.toUpperCase(), "Times New Roman",null, -1, -1, 240, "CENTER", "28", false, false); i++;
        pr3[i] = setP("", "Arial",null, -1, -1, 240, null, "24", false, false);i++;
        if (!(subName == null) && !subName.equals("") && !subName.isEmpty() && !subName.equals("Наименование документа")) {
            pr3[i] = setP(subName, "Arial",null, -1, -1,240, "CENTER", "24", false, true); i++;}
        pr3[i] = setP(type, "Arial", null, -1, -1, 360, "CENTER", "24", false, false); i++;
        if (albom!= null && !albom.isEmpty() && !albom.equals("")) {
            pr3[i] =  setP(albom, "Arial", null, -1, -1, 360, "CENTER", "24", false, true); i++;}
        pr3[i] = setP("ЛИСТ УТВЕРЖДЕНИЯ", "Arial", null, -1, -1, 360, "CENTER", "32", false, true);i++;
        pr3[i] = setP(docNumber.replace("{wrong}", ""),
                "Arial", null, -1, -1, 360, "CENTER", null, docNumber.contains("{wrong}"), true);i++;
        if (!(medium == null) && !medium.equals("") && !medium.equals("(вид носителя данных)")) {setP(medium, "Arial",null,
                -1, -1, 360, "CENTER", null, false, true); i++;}
        pr3[i] = setP("", "Arial", null, -1, -1, 360, "CENTER", "20", false, false);i++;
        if (!nPages.isEmpty()) {
            pr3[i] = setP(nPages, "Arial",null, -1, -1, 360, "CENTER", "28", false, true);}
        table.getContent().add(addRowWithMergedCells(false, null, pr3, null, 0, (int)(pageWidth*0.9), 0, 2));
        int num = 20;
        if (company.length() > 100 && company.length() <= 150)
            num--;
        if (company.length() >150 && company.length() <= 200 )
            num-=2;
        if (company.length() >200 && company.length() <= 300 )
            num-=3;
        if (albom!= null && !albom.isEmpty() && !albom.equals(""))
            num-=3;
        if (name.length() > 50 & name.length() <=75)
            num --;
        if (name.length() > 75 & name.length() <=120)
            num -=2;
        if (name.length() > 120 & name.length() <=170)
            num -=3;
        if (!nPages.isEmpty())
            num--;
        if (!(subName == null) && !subName.equals("") && !subName.isEmpty() && !subName.equals("Наименование документа"))
            num--;
        if (!(medium == null) && !medium.equals("") && !medium.equals("(вид носителя данных)"))
            num--;

        P[] pr_ = new P[num];
        for (int index = 0; index < num; index++)
            pr_[index] = factory.createP();

        P[] pr4 = null;
        P[] pr5 = null;
        List<String> remainStrings = new ArrayList<>();
        if (remained!= null && remained.size()!=0) remainStrings = DocBase.changeToString(remained);
        if (bound != -1) {
            pr4 = new P[bound+1];
            pr4[0] = setP("СОГЛАСОВАНО","Times New Roman", null, -1, -1, 240, null, null, false, false);
            for (int k = 1; k < remainStrings.size() - bound - 1; k ++ ) {
                pr4[k] = setP(remainStrings.get(k), "Times New Roman", null, -1, -1, 240, "CENTER", null, false, false);
            }
            pr5 = new P[remainStrings.size() - bound];
            for (int k = 0; k < remainStrings.size() - bound; k ++ ) {
                pr5[k] = setP(remainStrings.get(k), "Times New Roman", null, -1, -1, 240, "CENTER", null, false, false);
            }
        }
        else {
            if (remainStrings.size() != 0) {
                pr5 = new P[remainStrings.size()];
                for (int k = 0; k < remainStrings.size(); k ++ ) {
                    pr5[k] = setP(remainStrings.get(k), "Times New Roman", null, -1, -1, 240, "CENTER", null, false, false);
                }
            }
        }
        P[] pr6 = {setP(year, "Times New Roman", null, -1, -1, 240, "CENTER", null, false, true),
                setP(changeString, "Times New Roman", null, -1, -1, 240, "CENTER", null, false, true),
                setP(letter, "Times New Roman", null, -1, -1, 240, "RIGHT", null, false, true)};
        table.getContent().add(addRowWithMergedCells(false, pr4, pr_, pr5, (int)(0.3*pageWidth) ,(int)(0.3*pageWidth),
                (int)(0.3*pageWidth), 3));
        table.getContent().add(addRowWithMergedCells(false, null, pr6, null, 0, (int)(pageWidth*0.9), 0, 4));
        doc.getMainDocumentPart().getContent().add(0,table);
        //  template.getMainDocumentPart().getContent().add(1, DocBase.makePageBr());
        SectPr sectPr= null;
        try {
            sectPr = HeaderFooter.process(doc);
        } catch (Exception e) {
            sendIndfo(e.getMessage());
        }
        CTPageNumber pageNumber = sectPr.getPgNumType();
        if (pageNumber==null) {
            pageNumber = Context.getWmlObjectFactory().createCTPageNumber();
            sectPr.setPgNumType(pageNumber);
        }
        pageNumber.setStart(BigInteger.ONE);
        P p = factory.createP();
        PPr ppr = factory.createPPr();
        p.setPPr(ppr);
        ppr.setSectPr(sectPr);
        doc.getMainDocumentPart().getContent().add(1, p);


    }
    private  void setSecondPage() {
        //   template.getMainDocumentPart().getContent().add(setP("", "Times New Roman", null, null, null, 480, null, null));
//        newDoc.getMainDocumentPart().getContent().add(setP("",
//                "Times New Roman", null, null, null, 480, null, null, false, false));
//        newDoc.getMainDocumentPart().getContent().add(setP("",
//                "Times New Roman", null, null, null, 480, null, null, false, false));
        CTVerticalJc ctVerticalJc = new CTVerticalJc();
        ctVerticalJc.setVal(STVerticalJc.CENTER);
        Tbl table = new Tbl();
        table.setTblPr(new TblPr());
        TblWidth width = new TblWidth();
        width.setType("pct");
        width.setW(BigInteger.valueOf((int) (0.5 * pageWidth)));
        table.getTblPr().setTblW(width);
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

        P[] pr1 = {setP("УТВЕРЖДЕНО", "Times New Roman", null, -1, -1, 480, null, null, false, false),
                setP(docNumber.replace("{wrong}",""), "Courier New", null, -1, -1, 240, "LEFT", "20", false, false),
        };

        P[] pr = {setP("", "Times New Roman", null, -1, -1, 240, null, null, false, false),
                setP("", "Times New Roman", null, -1, -1, 240, null, null, false, false),
                setP("", "Times New Roman", null, -1, -1, 240, null, null, false, false) };
        table.getContent().add(addRowWithMergedCells(true, pr1,pr,pr,(int)(0.3*pageWidth) ,(int)(0.3*pageWidth),
                (int)(0.3*pageWidth), 1 ));

        P[] pr3 = new P[9];
        int i = 0;
        pr3[i] = setP("","Times New Roman", null, -1, -1, 240, null, null, false, false);i++;
        pr3[i] = setP(name.toUpperCase(), "Times New Roman",null, -1, -1, 240, "CENTER", null, false, false); i++;
        pr3[i] = setP("", "Arial",null, -1, -1, 240, null, null, false, false);i++;
        if (!(subName == null) && !subName.equals("") && !subName.equals("Наименование документа")) {
            pr3[i] = setP(subName, "Arial",null, -1, -1,240, "CENTER", null, false, true); i++;}

        pr3[i] = setP(type, "Arial", null, -1, -1, 360, "CENTER", "24", false, false); i++;
        if (albom!= null && !albom.isEmpty() && !albom.equals("")) {
            pr3[i] =  setP(albom, "Arial", null, -1, -1, 360, "CENTER", "24", false, true); i++;}
        pr3[i] = setP(docNumber.replace("{wrong}", "").replace("-ЛУ",""),
                "Arial", null, -1, -1, 360, "CENTER", null, docNumber.contains("{wrong}"), true);i++;
        if (!(medium == null) && !medium.equals("")&& !medium.equals("(вид носителя данных)")) {setP(medium, "Arial",null,
                -1, -1, 360, "CENTER", null, false, true); i++;}
        pr3[i] = setP("", "Arial", null, -1, -1, 360, "CENTER", "20", false, false);i++;
            pr3[i] = setP(nPages2, "Arial",null, -1, -1, 360, "CENTER", "28", nPages2.equals("Листов ___"), true);
        table.getContent().add(addRowWithMergedCells(false, null, pr3, null, 0, (int)(0.9*pageWidth), 0, 2));
        int num = 26;
        if (albom!= null && !albom.isEmpty() && !albom.equals(""))
            num-=3;
        if (name.length() > 50 & name.length() <=75)
            num --;
        if (name.length() > 75 & name.length() <=120)
            num -=2;
        if (name.length() > 120 & name.length() <=170)
            num -=3;
        if (!(subName == null) && !subName.equals("") && !subName.isEmpty() && !subName.equals("Наименование документа"))
            num--;
        if (!(medium == null) && !medium.equals("") && !medium.equals("(вид носителя данных)"))
            num--;
        if (docNumber.length() > 15)
            num--;

        P[] pr_ = new P[num];
        for (int index = 0; index < num; index++)
            pr_[index] = factory.createP();
        table.getContent().add(addRowWithMergedCells(false, pr, pr_, pr,(int)(0.3*pageWidth) ,(int)(0.3*pageWidth),
                (int)(0.3*pageWidth), 3));
        P[] pr5 = {setP(year, "Times New Roman", null, -1, -1, 240, "CENTER", null, false, true),
                setP(changeString, "Times New Roman", null, -1, -1, 240, "CENTER", null, false, true),
                setP(letter, "Times New Roman", null, -1, -1, 240, "RIGHT", null, false, true)};
        table.getContent().add(addRowWithMergedCells(false, null, pr5, null, 0, (int)(0.9*pageWidth), 0, 4));
        doc.getMainDocumentPart().getContent().add(2, table);
        Br objBr = new Br();
        objBr.setType(STBrType.TEXT_WRAPPING);
        P p = factory.createP();
        R r = factory.createR();
        r.getContent().add(objBr);
        p.getContent().add(r);
        doc.getMainDocumentPart().getContent().add(3, p);
    }

    private  Tr addRowWithMergedCells(boolean image, P[] ps, P[] ps2, P[] ps3, int width1, int width2 , int width3, int num) {

        Tr row = factory.createTr();
        row.setTrPr(factory.createTrPr());

        if (num == 0 || num == 2 || num == 4) {
            addMergedColumn(row, image, -1, 0);
            addTableCell(row, ps2, width2, 2);
        }
        if (num == 1 || num == 3) {
            int width = (num == 1) ? 500 : 0;
            addMergedColumn(row, image, -1, width);
            addTableCell(row, ps, width1, -1);
            addTableCell(row, ps2, width2, -1);
            addTableCell(row, ps3, width3, -1);
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

        CTVerticalJc ctVerticalJc = new CTVerticalJc();
        ctVerticalJc.setVal(STVerticalJc.BOTTOM);
        tableCellProperties.setVAlign(ctVerticalJc);
        TblWidth tableWidth = factory.createTblWidth();
        tableWidth.setType("dxa");
        tableWidth.setW(BigInteger.valueOf(width));
        tableCell.setTcPr(tableCellProperties);
        tableCell.getTcPr().setTcW(tableWidth);

//
//         TcPrInner.GridSpan gridSpan = new TcPrInner.GridSpan();
//            gridSpan.setVal(BigInteger.valueOf(grid));
//            tableCellProperties.setGridSpan(gridSpan);

        if(image) {
            tableCell.getContent().add(DocxMethods.newImage(doc, new File("resource/table.gif")));
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
        if (grid == 6) {
            CTVerticalJc valign = factory.createCTVerticalJc();
            valign.setVal(STVerticalJc.TOP);
            tableCellProperties.setVAlign(valign);
        }

        //     tableCellProperties.setHideMark(value);
        if (grid!=-1) {
            TcPrInner.GridSpan gridSpan = new TcPrInner.GridSpan();
            gridSpan.setVal(BigInteger.valueOf(3));
            tableCellProperties.setGridSpan(gridSpan);
        }
        tableCellProperties.setTcW(tableWidth);
        tc1.setTcPr(tableCellProperties);
        tr.getContent().add(tc1);
    }

    private  P setP (String text, String font, String style, int ilvl, int numId, int spacing, String align,
                     String size, boolean highlight, boolean setBold) {
        P p = new P();
        DocBase.setRightP(p, text);
        p.setPPr(new PPr());

        if (highlight)
            DocBase.setHighlight(p, "yellow");
        DocBase.setStyle(p, size, font, style, ilvl, numId, spacing, align, setBold);
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
        if (remained.size() == 0)
            return;
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
            sendIndfo(e.getMessage());
        } catch (FileNotFoundException e) {
            sendIndfo(e.getMessage());
        }

    }

    private boolean[] getParagraphesFromTbl (P content, P content2) {
        boolean[] inTable = {false, false};
        List<Object> paragraphes = new ArrayList<>();
        MainDocumentPart main =  doc.getMainDocumentPart();
        for (int i = 0; i <main.getContent().size(); i++) {
            if (main.getContent().get(i) instanceof Tbl) {
                Tbl tbl = (Tbl)main.getContent().get(i);
                List<Object> contents = tbl.getContent();
                for (Object o : contents) {
                    if (o instanceof  P) {
                        if (o.equals(content)){
                            inTable[0] = true;
                            paragraphes = DocxMethods.getAllElementFromObject(tbl, P.class);
                            main.getContent().remove(tbl);
                            break;
                        }

                    }
                }
            }
        }
        if (content2!= null) {
            for (int i = 0; i <main.getContent().size(); i++) {
                if (main.getContent().get(i) instanceof Tbl) {
                    Tbl tbl = (Tbl)main.getContent().get(i);
                    List<Object> contents = tbl.getContent();
                    for (Object o : contents) {
                        if (o instanceof  P) {
                            if (o.equals(content2)){
                                inTable[1] = true;
                                paragraphes.addAll(DocxMethods.getAllElementFromObject(tbl, P.class));
                                main.getContent().remove(tbl);
                                break;
                            }

                        }
                    }
                }
            }
        }
//        for (int i = paragraphes.size() - 1; i>=0; i--) {
//            main.getContent().add(0, paragraphes.get(i));
//        }
        return inTable;
    }

}
