package Model;

import org.docx4j.XmlUtils;
import org.docx4j.jaxb.Context;
import org.docx4j.openpackaging.exceptions.InvalidFormatException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.NumberingDefinitionsPart;
import org.docx4j.wml.Numbering;
import org.docx4j.wml.ObjectFactory;

import javax.xml.bind.JAXBException;

public class StyleSetter {

    WordprocessingMLPackage word;
    ObjectFactory factory = Context.getWmlObjectFactory();
    NumberingDefinitionsPart ndp;
    public StyleSetter(WordprocessingMLPackage word) {
        this.word = word;
    }

    public  Boolean setStyle () {
        try {
            ndp = new NumberingDefinitionsPart();
            word.getMainDocumentPart().addTargetPart(ndp);
            ndp.setJaxbElement( (Numbering) XmlUtils.unmarshalString(numbering) );
        } catch (InvalidFormatException e) {
            return false;

        } catch (JAXBException e) {
            return false;
        }
        return true;
    }
    private String numbering = "<w:numbering xmlns:wpc=\"http://schemas.microsoft.com/office/word/2010" +
            "/wordprocessingCanvas\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\"" +
            " xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:r=\"http://schemas.openxmlformats.org/" +
            "officeDocument/2006/relationships\" xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"" +
            " xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:wp14=\"http://schemas.microsoft.com/office/word/2010/" +
            "wordprocessingDrawing\" xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\"" +
            " xmlns:w10=\"urn:schemas-microsoft-com:office:word\" xmlns:w=\"http://schemas.openxmlformats.org/" +
            "wordprocessingml/2006/main\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:w15=\"" +
            "http://schemas.microsoft.com/office/word/2012/wordml\" xmlns:wpg=\"http://schemas.microsoft.com/" +
            "office/word/2010/wordprocessingGroup\" xmlns:wpi=\"http://schemas.microsoft.com/office/word/2010" +
            "/wordprocessingInk\" xmlns:wne=\"http://schemas.microsoft.com/office/word/2006/wordml\" xmlns:wps=\"" +
            "http://schemas.microsoft.com/office/word/2010/wordprocessingShape\" mc:Ignorable=\"w14 w15 wp14\">" +
            "<w:abstractNum w:abstractNumId=\"0\"><w:nsid w:val=\"0A7D028C\"/><w:multiLevelType w:val=\"multilevel\"/>" +
            "<w:tmpl w:val=\"6504C7A4\"/><w:numStyleLink w:val=\"MyStyle\"/></w:abstractNum>" +
            "<w:abstractNum w:abstractNumId=\"1\"><w:nsid w:val=\"14CA30EA\"/>" +
            "<w:multiLevelType w:val=\"multilevel\"/><w:tmpl w:val=\"6504C7A4\"/><w:numStyleLink w:val=\"MyStyle\"/>" +
            "</w:abstractNum><w:abstractNum w:abstractNumId=\"2\"><w:nsid w:val=\"1CBA320B\"/>" +
            "<w:multiLevelType w:val=\"multilevel\"/><w:tmpl w:val=\"6504C7A4\"/><w:numStyleLink w:val=\"MyStyle\"/>" +
            "</w:abstractNum><w:abstractNum w:abstractNumId=\"3\"><w:nsid w:val=\"67CB30D5\"/>" +
            "<w:multiLevelType w:val=\"multilevel\"/><w:tmpl w:val=\"6504C7A4\"/><w:styleLink w:val=\"MyStyle\"/>" +
            "<w:lvl w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"decimal\"/><w:lvlText w:val=\"%1    \"/>" +
            "<w:lvlJc w:val=\"left\"/><w:pPr><w:ind w:left=\"360\" w:hanging=\"360\"/>" +
            "</w:pPr><w:rPr><w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:hint=\"default\"/><w:b/>" +
            "<w:sz w:val=\"32\"/></w:rPr></w:lvl><w:lvl w:ilvl=\"1\">" +
            "<w:start w:val=\"1\"/><w:numFmt w:val=\"decimal\"/><w:lvlText w:val=\"%1.%2    \"/>" +
            "<w:lvlJc w:val=\"left\"/><w:pPr><w:ind w:left=\"360\" w:hanging=\"360\"/></w:pPr><w:rPr>" +
            "<w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:hint=\"default\"/>" +
            "<w:b/><w:sz w:val=\"28\"/></w:rPr></w:lvl><w:lvl w:ilvl=\"2\"><w:start w:val=\"1\"/>" +
            "<w:numFmt w:val=\"decimal\"/><w:lvlText w:val=\"%1.%2.%3\"/><w:lvlJc w:val=\"left\"/><w:pPr>" +
            "<w:ind w:left=\"360\" w:hanging=\"360\"/></w:pPr><w:rPr><w:rFonts w:hint=\"default\"/><w:b/>" +
            "<w:sz w:val=\"24\"/></w:rPr></w:lvl><w:lvl w:ilvl=\"3\"><w:start w:val=\"1\"/>" +
            "<w:numFmt w:val=\"russianLower\"/><w:lvlText w:val=\"%4)    \"/><w:lvlJc w:val=\"left\"/><w:pPr>" +
            "<w:ind w:left=\"1068\" w:hanging=\"360\"/></w:pPr><w:rPr>" +
            "<w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:hint=\"default\"/><w:sz w:val=\"24\"/>" +
            "</w:rPr></w:lvl><w:lvl w:ilvl=\"4\"><w:start w:val=\"1\"/><w:numFmt w:val=\"decimal\"/>" +
            "<w:lvlText w:val=\"%5)    \"/><w:lvlJc w:val=\"left\"/><w:pPr><w:ind w:left=\"1776\" w:hanging=\"360\"/>" +
            "</w:pPr><w:rPr><w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:hint=\"default\"/>" +
            "<w:sz w:val=\"24\"/></w:rPr></w:lvl><w:lvl w:ilvl=\"5\"><w:start w:val=\"1\"/><w:numFmt w:val=\"lowerRoman\"/>" +
            "<w:lvlText w:val=\"(%6)\"/><w:lvlJc w:val=\"left\"/><w:pPr><w:ind w:left=\"9636\" w:hanging=\"360\"/>" +
            "</w:pPr><w:rPr><w:rFonts w:hint=\"default\"/></w:rPr></w:lvl><w:lvl w:ilvl=\"6\"><w:start w:val=\"1\"/>" +
            "<w:numFmt w:val=\"decimal\"/><w:lvlText w:val=\"%7.\"/><w:lvlJc w:val=\"left\"/><w:pPr>" +
            "<w:ind w:left=\"9996\" w:hanging=\"360\"/></w:pPr><w:rPr><w:rFonts w:hint=\"default\"/></w:rPr>" +
            "</w:lvl><w:lvl w:ilvl=\"7\"><w:start w:val=\"1\"/><w:numFmt w:val=\"lowerLetter\"/><w:lvlText w:val=\"%8.\"/>" +
            "<w:lvlJc w:val=\"left\"/><w:pPr><w:ind w:left=\"10356\" w:hanging=\"360\"/></w:pPr><w:rPr>" +
            "<w:rFonts w:hint=\"default\"/></w:rPr></w:lvl><w:lvl w:ilvl=\"8\"><w:start w:val=\"1\"/>" +
            "<w:numFmt w:val=\"lowerRoman\"/><w:lvlText w:val=\"%9.\"/><w:lvlJc w:val=\"left\"/>" +
            "<w:pPr><w:ind w:left=\"10716\" w:hanging=\"360\"/></w:pPr><w:rPr><w:rFonts w:hint=\"default\"/>" +
            "</w:rPr></w:lvl></w:abstractNum><w:num w:numId=\"1\"><w:abstractNumId w:val=\"3\"/>" +
            "</w:num><w:num w:numId=\"2\"><w:abstractNumId w:val=\"1\"/></w:num><w:num w:numId=\"3\">" +
            "<w:abstractNumId w:val=\"0\"/></w:num><w:num w:numId=\"4\"><w:abstractNumId w:val=\"2\"/></w:num>" +
            "<w:num w:numId=\"5\"><w:abstractNumId w:val=\"3\"/><w:lvlOverride w:ilvl=\"0\">" +
            "<w:startOverride w:val=\"1\"/></w:lvlOverride></w:num></w:numbering>";
}
