import org.docx4j.wml.P;

import java.util.ArrayList;
import java.util.List;

public class Paragraph {

    private P p;
    private List<Paragraph> internalParagraphes;
    private boolean special;

    public Paragraph (P p, boolean special) {
        this.p = p;
        internalParagraphes = new ArrayList<>();
        this.special = special;
    }

    public P getP() {
        return p;
    }

    public List<Paragraph> getInternalParagraphes() {
        return internalParagraphes;
    }

    public void addNewParagraph (Paragraph p) {
        internalParagraphes.add(p);
    }

    public boolean isSpecial() {
        return special;
    }

}
