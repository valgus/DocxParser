package Model;

public enum Document {

    Tecnicheskoe_zadanie(19, false, false, true, true, "docx/gost19_tehnicheskoe_zadanie.docx"),
    Tecnicheskoe_zadanie_34(34, true, true, true, true, "docx/tz_as.docx"),
    Rukovodstvo_operatora(19, true, true, true, true, "docx/gost19_rukovodstvo_operatora.docx" ),
    Spezifikation(19, false, false, false, false, "docx/gost19_specifikacija.docx" ),
    Poiasnitelnaya_zapiska(19, false, false, true, true, "docx/gost19_pojasnitelnaja_zapiska.docx" ),
    Rukovodstvo_po_technicheskomu_obsluzhivaniu(19, false, false, true, false,"docx/"),
 //   Opisanie_yazika(19, true, true, true, true, "docx/" ),
    Rukovodstvo_programmista(19, true, true, true, true,"docx/gost19_rukovodstvo_programmista.docx"),
    Rukovodstvo_sistemnogo_programmista(19, true, true, true, true,
            "docx/gost19_rukovodstvo_sistemnogo_programmista.docx"),
    Opisanie_primeneniya(19, true, true, true, true, "docx/gost19_opisanie_primenenija.docx"),
    Vedomost_ekspluatacionnich_dokumentov(19, false, false, false, false,
            "docx/gost19_vedomost_jekspluatacionnyh_dokumentov.docx"),
 //   Vedomost_ekspluatacionnich_dokumentov_34(34, false, false, false, false, "docx/"),
    Formulyar(19, false, false, true, true, "docx/gost19_formuljar.docx"),
    Formulyar_34(34, false, false, true, true, "docx/rd fo.docx"),
    Programma_i_metodika_ispitanii(19, false, false, true, false, "docx/gost19_programma_i_metodika_ispytanij.docx"),
    Opisanie_Programmi(19, true, true, true, true, "docx/gost19_opisanie_programmy.docx"),
    Tekst_programmi(19, false, false, true, false, "docx/gost19_tekst_programmy.docx"),
 //   Rukovodstvo_administratora(19, true, true, true, true, "docx/"),
 //   Vedomost_derzhatelei_podlenikov(19, false, false, false, false, "docx/"),
    Programma_i_metodika_ispitanii_34(34, false, false, true, false, "docx/rd pm.docx"),
    TEchnologicheskaya_instrukcia(34, false, false, false, false, "docx/gost34_tehnologicheskaja_instrukcija.docx"),
    Schema_funkcion_strukturi(34, false, false, false, false, "docx/gost34_shema_funkcionalnoj_struktury.docx"),
    Schema_struct_kompleksa_tech_sredstv(34,false, false, false, false,
            "docx/gost34_shema_strukturnaja_kompleksa_tehnicheskih_sredstv.docx" ),
    Schema_organiz_structuri(34,false, false, false, false, "docx/gost34_shema_organizacionnoj_struktury.docx" ),
    Schema_avtomatizacii(34,false, false, false, false, "docx/gost34_shema_avtomatizacii.docx" ),
    Rukovodstvo_polzovatelya(34,false, false, false, false, "docx/gost34_rukovodstvo_polzovatelja.docx" ),
    Proektnaya_ocenka_nadeznosti_systemy(34,false, false, false, false,
            "docx/gost34_proektnaja_ocenka_nadezhnosti_sistemy.docx"),
    Perechen_vichodnih_signalov(34,false, false, false, false, "docx/gost34_perechen_vyhodnyh_signalov_dokumentov_.docx"),
    Perechen_vchodnih_signalov(34,false, false, false, false, "docx/gost34_perechen_vhodnyh_signalov_i_dannyh.docx"),
    Pasport(34,false, false, false, false, "docx/gost34_pasport.docx"),
    Opisanie_KTS(34,false, false, false, false, "docx/gost34_opisanie_kts.docx"),
    Opisanie_system_klassifik_i_kodir(34,false, false, false, false,
            "docx/gost34_opisanie_sistem_klassifikacii_i_kodirovanija.docx"),
    Opisanie_program_obespecheniya(34,false, false, false, false, "docx/gost34_opisanie_programmnogo_obespechenija.docx"),
    Opisanie_postanovki_zadach(34,false, false, false, false, "docx/gost34_opisanie_postanovki_zadachi.docx"),
    Opisanie_organ_structuri(34,false, false, false, false, "docx/gost34_opisanie_organizacionnoj_struktury.docx"),
    Opisanie_organ_inf_bazi(34,false, false, false, false, "docx/gost34_opisanie_organizacii_informacionnoj_bazy.docx"),
    Opisanie_inf_onespech_systemi(34,false, false, false, false,
            "docx/gost34_opisanie_informacionnogo_obespechenija.docx"),
    Opisanie_avtomat_funkcii(34,false, false, false, false, "docx/gost34_opisanie_avtomatiziruemyh_funkcij.docx"),
    Opisanie_tech_processa(34,false, false, false, false,
            "docx/gost34_opisanie_tehnologicheskogo_processa_obrabotki_dannyh_vkljuchaja_teleobrabotku_.docx"),
    Opisanie_proektnoi_proceduri(34,false, false, false, false, "docx/gost34_opisanie_proektnoj_procedury.docx"),
    Opisanie_algoritma(34,false, false, false, false, "docx/gost34_opisanie_algoritma.docx"),
    Opisanie_obcee_opisanie_systemi(34,false, false, false, false, "docx/gost34_obshhee_opisanie_sistemy.docx"),
    Massiv_vhodnich_dannich(34,false, false, false, false, "docx/gost34_massiv_vhodnyh_dannyh.docx"),
    Katalog_BD(34,false, false, false, false, "docx/gost34_katalog_bd.docx"),
    Instrukciya_po_ekspluat_AS(34,false, false, false, false, "docx/gost34_instrukcija_po_jekspluatacii_kts.docx"),
 //   Sostav_vihodnich_dannich(34,false, false, false, false, "docx/"),
    Opisanie_massiva_informacii(34,false, false, false, false, "docx/opisanie_massiva_informazii.docx");

    private final boolean annotation,contents,newPart, merge;
    private final int gost;
    private String template;
    Document(int gost, boolean annotation, boolean contents, boolean newPart, boolean merge, String template){
        this.gost = gost;
        this.annotation = annotation;
        this.contents = contents;
        this.merge = merge;
        this.newPart = newPart;
        this.template = template;
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

    public String getTemplate() {
        return template;
    }


    public void getName() {

    }
}
