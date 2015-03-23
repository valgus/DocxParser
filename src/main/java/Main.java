public class Main {
    public static void main(String[] args) {
      AlternativeFlow comparisonWithTemplate = new AlternativeFlow();
        comparisonWithTemplate.setTwoDocx("docx/template.docx","docx/document.docx");
        try {
            comparisonWithTemplate.setAppropriateText();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

}