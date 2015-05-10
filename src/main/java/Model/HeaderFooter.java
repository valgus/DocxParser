package Model;

import org.docx4j.convert.out.common.wrappers.ConversionSectionWrapperFactory;
import org.docx4j.jaxb.Context;
import org.docx4j.model.structure.HeaderFooterPolicy;
import org.docx4j.model.structure.PageDimensions;
import org.docx4j.model.structure.SectionWrapper;
import org.docx4j.openpackaging.exceptions.InvalidFormatException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.Part;
import org.docx4j.openpackaging.parts.PartName;
import org.docx4j.openpackaging.parts.WordprocessingML.FooterPart;
import org.docx4j.openpackaging.parts.WordprocessingML.HeaderPart;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.openpackaging.parts.relationships.RelationshipsPart;
import org.docx4j.relationships.Relationship;
import org.docx4j.wml.*;

import javax.xml.bind.JAXBElement;
import java.io.File;
import java.math.BigInteger;
import java.util.List;

public class HeaderFooter extends  SectionWrapper{
    private static int pageWidth = new PageDimensions().getWritableWidthTwips();

    private static org.docx4j.wml.ObjectFactory factory = new org.docx4j.wml.ObjectFactory();

    protected HeaderFooter(SectPr sectPr, HeaderFooterPolicy previousHF, RelationshipsPart rels, BooleanDefaultTrue evenAndOddHeaders) {
        super(sectPr, previousHF, rels, evenAndOddHeaders);
    }


    public static Tbl makeTable(Part part) {

        Tbl tbl = factory.createTbl();
        tbl.setTblPr(factory.createTblPr());
        TblWidth width = new TblWidth();
        width.setType("pct");
        width.setW(BigInteger.valueOf((int)(0.5*pageWidth)));
        tbl.getTblPr().setTblW(width);
        CTBorder border = new CTBorder();
        border.setSz(new BigInteger("0"));
        border.setSpace(new BigInteger("0"));
        border.setVal(STBorder.SINGLE);
        TblBorders borders = new TblBorders();
        borders.setBottom(border);
        borders.setLeft(border);
        borders.setRight(border);
        borders.setTop(border);
        borders.setInsideH(border);
        borders.setInsideV(border);
        tbl.getTblPr().setTblBorders(borders);
        tbl.getContent().add(addTr(null,null,null,null,null));
        tbl.getContent().add(addTr("Изм.","Лист","№ докум.","Подп.","Дата"));
        tbl.getContent().add(addTr(null,null,null,null,null));
        tbl.getContent().add(addTr("Инв. № подл.","Подп. и дата","Взам. инв. ","№Инв. № дубл.","Подп. и дата"));
        return tbl;
    }

    public static Tr addTr ( String s1, String s2, String s3, String s4, String s5) {
        Tr tr = factory.createTr();
        tr.setTrPr(factory.createTrPr());
        tr.getContent().add(addTc(s1));
        tr.getContent().add(addTc(s2));
        tr.getContent().add(addTc(s3));
        tr.getContent().add(addTc(s4));
        tr.getContent().add(addTc(s5));
        CTHeight ctHeight = new CTHeight();
        ctHeight.setHRule(STHeightRule.AT_LEAST);
        JAXBElement<CTHeight> jaxbElement = factory.createCTTrPrBaseTrHeight(ctHeight);
        tr.getTrPr().getCnfStyleOrDivIdOrGridBefore().add(jaxbElement);
        return tr;
    }

    public static Tc addTc (String text) {
        Tc tc = factory.createTc();
        P p = factory.createP();
        p.setPPr(factory.createPPr());
        R r = factory.createR();
        p.getContent().add(r);
        Text t = factory.createText();
        r.getContent().add(t);
        t.setValue((text == null) ? "" : text);
        DocBase.setAlign(p, "CENTER");
        DocBase.setSpacing(p, 0, 0);
        DocBase.setSize(p, "16");
        DocBase.setFont(p, "Times New Roman");
        tc.getContent().add(p);
        return tc;
    }

    public static Hdr getHdr() throws Exception
    {
        // AddPage Numbers
        CTSimpleField pgnum = factory.createCTSimpleField();
        pgnum.setInstr(" PAGE \\* MERGEFORMAT ");
        RPr RPr = factory.createRPr();
        RPr.setNoProof(new BooleanDefaultTrue());
        PPr ppr = factory.createPPr();
        Jc jc = factory.createJc();
        jc.setVal(JcEnumeration.CENTER);
        ppr.setJc(jc);
        PPrBase.Spacing pprbase = factory.createPPrBaseSpacing();
        pprbase.setBefore(BigInteger.valueOf(240));
        pprbase.setAfter(BigInteger.valueOf(0));
        ppr.setSpacing(pprbase);

        R run = factory.createR();
        run.getContent().add(RPr);
        pgnum.getContent().add(run);

        JAXBElement<CTSimpleField> fldSimple = factory.createPFldSimple(pgnum);
        P para = factory.createP();
        para.getContent().add(fldSimple);
        para.setPPr(ppr);
        // Now add our paragraph to the footer
        Hdr hdr = factory.createHdr();
        hdr.getContent().add(para);
        return hdr;
    }

    public static SectPr process(WordprocessingMLPackage word) throws Exception {
        MainDocumentPart mdp = word.getMainDocumentPart();


        HeaderPart cover_hdr_part = new HeaderPart(new PartName(
                "/word/cover-header.xml")),
                content_hdr_part = new HeaderPart(
                new PartName("/word/content-header.xml"));
        word.getParts().put(cover_hdr_part);
        word.getParts().put(content_hdr_part);
        cover_hdr_part.setPackage(word);
        content_hdr_part.setPackage(word);

        Hdr cover_hdr = factory.createHdr(), content_hdr = getHdr();
        P p = factory.createP();
        DocBase.setText(p, "", false);
        cover_hdr.getContent().add(p);
        content_hdr.getContent().add(p);

        // Bind the header JAXB elements as representing their header parts
        cover_hdr_part.setJaxbElement(cover_hdr);
        content_hdr_part.setJaxbElement(content_hdr);

        // Add the reference to both header parts to the Main Document Part
        Relationship cover_hdr_rel = mdp.addTargetPart(cover_hdr_part);
        Relationship content_hdr_rel = mdp.addTargetPart(content_hdr_part);





        //DO FOOTER PART NOW ***********************************************************************

        FooterPart cover_ftr_part = new FooterPart(new PartName(
                "/word/cover-footer.xml")), content_ftr_part = new FooterPart(
                new PartName("/word/content-footer.xml"));

        word.getParts().put(cover_ftr_part);
        word.getParts().put(content_ftr_part);
        cover_ftr_part.setPackage(word);
        content_ftr_part.setPackage(word);
        //Ftr cover_ftr = factory.createFtr(), content_ftr = factory.createFtr();
        //page number test

        Ftr cover_ftr = factory.createFtr(), content_ftr = factory.createFtr();
        content_ftr.getContent().add(p);
        cover_ftr.getContent().add(p);

        // Bind the header JAXB elements as representing their header parts
        cover_ftr_part.setJaxbElement(cover_ftr);
        content_ftr_part.setJaxbElement(content_ftr);

        // Add the reference to both header parts to the Main Document Part
        Relationship cover_ftr_rel = mdp.addTargetPart(cover_ftr_part);
        Relationship content_ftr_rel = mdp
                .addTargetPart(content_ftr_part);


     //   cover_ftr.getContent().add(makeTable(cover_ftr_part));
        content_ftr.getContent().add(makeTable(content_hdr_part));


        //PUT THE DOCUMENT TOGETHER


        List<SectionWrapper> sections = word.getDocumentModel().getSections();

        SectPr sectPr1 = sections.get(sections.size() - 1).getSectPr();
        if (sectPr1 == null)  {
            sectPr1 = factory.createSectPr();
            sections.get(sections.size() - 1).setSectPr(sectPr1);
        }
        BooleanDefaultTrue booleanDefaultTrue = new BooleanDefaultTrue();
        booleanDefaultTrue.setVal(true);

        // link cover and content headers
        HeaderReference hdr_ref;

        hdr_ref = factory.createHeaderReference();
        hdr_ref.setId(cover_hdr_rel.getId());
        hdr_ref.setType(HdrFtrRef.FIRST);
        sectPr1.getEGHdrFtrReferences().add(hdr_ref);

        hdr_ref = factory.createHeaderReference();
        hdr_ref.setId(content_hdr_rel.getId());
        hdr_ref.setType(HdrFtrRef.DEFAULT);
        sectPr1.getEGHdrFtrReferences().add(hdr_ref);

        BooleanDefaultTrue boolanDefaultTrue = new BooleanDefaultTrue();
        sectPr1.setTitlePg(boolanDefaultTrue);


        // link cover and content footers
        FooterReference ftr_ref;

        ftr_ref = factory.createFooterReference();
        ftr_ref.setId(cover_ftr_rel.getId());
        ftr_ref.setType(HdrFtrRef.FIRST);
        sectPr1.getEGHdrFtrReferences().add(ftr_ref);
        ftr_ref = factory.createFooterReference();
        ftr_ref.setId(content_ftr_rel.getId());
        ftr_ref.setType(HdrFtrRef.DEFAULT);
        sectPr1.getEGHdrFtrReferences().add(ftr_ref);
        return sectPr1;
    }

    public static void main(String[] args) throws Exception {
        WordprocessingMLPackage word = WordprocessingMLPackage.createPackage();
        word.getMainDocumentPart().addParagraphOfText("sfvesdv");
        P p = factory.createP();
        PPr ppr = factory.createPPr();
        p.setPPr(ppr);
        SectPr sectPr= process(word);
        CTPageNumber pageNumber = sectPr.getPgNumType();
        if (pageNumber==null) {
            pageNumber = Context.getWmlObjectFactory().createCTPageNumber();
            sectPr.setPgNumType(pageNumber);
        }
        pageNumber.setStart(BigInteger.ONE);
        ppr.setSectPr(sectPr);
        word.getMainDocumentPart().addObject(p);
        word.getMainDocumentPart().addParagraphOfText("after sectPr");
        word.getMainDocumentPart().addObject( DocBase.makePageBr());
        word.getMainDocumentPart().addParagraphOfText("after sectPr");
        word.save(new File("header_test.docx"));
        DocxToPDFConverter.convert(DocxMethods.getTemplate("header_test.docx"));
    }
}
