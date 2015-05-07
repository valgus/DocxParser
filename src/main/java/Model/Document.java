package Model;

public enum Document {

    Tecnicheskoe_zadanie(19, false, false, true, true, ""),
    Rukovodstvo_operatora(19, true, true, true, true, "" ),
    Spezifikation(19, false, false, false, false, "" ),
    Poiasnitelnaya_zapiska(19, false, false, true, true, "" ),
    Rukovodstvo_po_technicheskomu_obsluzhivaniu(19, false, false, true, false,""),
    Opisanie_yazika(19, true, true, true, true, "" ),
    Rukovodstvo_programmista(19, true, true, true, true,""),
    Rukovodstvo_sistemnogo_programmista(19, true, true, true, true,""),
    Opisanie_primeneniya(19, true, true, true, true, ""),
    Vedomost_ekspluatacionnich_dokumentov(19, false, false, false, false, ""),
    Formulyar(19, false, false, true, true, ""),
    Programma_i_metodika_ispitanii(19, false, false, true, false, ""),
    Opisanie_Programmi(19, true, true, true, true, ""),
    Tekst_programmi(19, false, false, true, false, ""),
    Rukovodstvo_administratora(19, true, true, true, true, "");

    private final boolean annotation,contents,newPart, merge;
    private final int gost;
    String doc;
    Document(int gost, boolean annotation, boolean contents, boolean newPart, boolean merge, String doc){
        this.gost = gost;
        this.annotation = annotation;
        this.contents = contents;
        this.merge = merge;
        this.newPart = newPart;
        this.doc = doc;
    }

    public boolean isAnnotation() {
        return annotation;
    }

    public boolean isContents() {
        return contents;
    }

    public boolean isNewPart() {
        return newPart;
    }

    public boolean isMerge() {
        return merge;
    }

    public int getGost() {
        return gost;
    }

    public String getDoc() {
        return doc;
    }
}
