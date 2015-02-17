import java.util.LinkedList;
import java.util.List;

public class Title {

    private final int lvl;
    private final String name;
    private final String attributes;
    public Title (int number, String name, String attributes) {
        lvl = number;
        this.name = name;
        this.attributes = attributes;
    }

    public int getLvl() {
        return lvl;
    }

    public String getName() {
        return name;
    }

    public String getAttributes() {
        return attributes;
    }

}
