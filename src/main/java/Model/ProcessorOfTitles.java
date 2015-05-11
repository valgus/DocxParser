package Model;

import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.wml.*;

import java.io.File;
import java.io.FileNotFoundException;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Pattern;

public class ProcessorOfTitles {


    private List<P> firstPage;
    private int GOSTnumber;
    boolean isFirstpage;
    private int index;
    private String name;
    private List<Integer> indexName;
    private String type;
    private int typeIndex;
    private String agreement;
    private int[] agreementIndexes;
    private int approveIndex;
    private String approve;
    private String nameOfCompany;
    private int docNumberIndex;
    private String docNumber;
    private String albom;
    private  String pageNumber;
    private int pageNumberIndex;
    private String medium;
    private String subName;
    private List<P> remained;
    private ProcessDocument process;

    public boolean sendIndfo(String s) {
        return process.sendInfo(s);
    }

    public ProcessorOfTitles(List<P> firstPage, String type, int GOSTnumber, String name, ProcessDocument process) {
        this.firstPage = new ArrayList<>(firstPage);
        this.remained = new ArrayList<>();
        docNumber = "";
        medium = "";
        subName = "";
        albom = "";
        pageNumber = "";
        approve = "";
        agreement = "";
        isFirstpage = false;
        index = -1;
        typeIndex = -1;
        agreementIndexes = new int[2];
        agreementIndexes[0] = -1;
        agreementIndexes[1] = -1;
        approveIndex  = -1;
        docNumberIndex = -1;
        pageNumberIndex = -1;
        nameOfCompany = "";
        this.type = type;
        this.GOSTnumber = GOSTnumber;
        this.name = name;
        this.process = process;

    }
    String findMainElements() throws Exception{
        String s;
        String res = null;
        for (int i = 0; i < firstPage.size(); i++)  {
            s = DocBase.getText(firstPage.get(i));
            if (s.toLowerCase().equals("лист утверждения") || ngrammPossibility(s, "лист утверждения") >=0.5 ) {
                isFirstpage = true;
                index = i;
            }
        }
        if (!isFirstpage) {
            boolean is = sendIndfo("Is not the first page!");
            if (is)
                res =  "Создать пустой лист утверждения";
        }
        indexName = findName(0);
        if (indexName == null || indexName.size() == 0) {
            sendIndfo("The inserted name is not correct");
            throw new Exception();

        }
        findType();
        findAgreementsAndApprove();
        if (agreementIndexes[0] != -1) {
            findNameOfCompany();
            setSubscribes();
        }
        setDocNumber();
        setAlbom();
        setPageNumber();
        setSubNameAndMedium();
        setRemained();
        return  res;
    }




    private double ngrammPossibility (String actual, String checked) {
        if (actual.length() < 2)
            return 0;
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

    private List<Integer> findName (int start) {
        String s;
        List<Integer> indexes = new ArrayList<>();
        for (int i = start; i< firstPage.size(); i++) {
            s = DocBase.getText(firstPage.get(i));
            if (s.toLowerCase().equals(name.toLowerCase())) {
                indexes.add(i);
                return indexes;
            }
            else {
                if (name.toLowerCase().contains(s.toLowerCase()) ||
                        ngrammPossibility(name, s) >=0.3) {
                    indexes.add(i);
                    indexes.addAll(findName(i));
                    return indexes;
                }else
                if (ngrammPossibility(name, s) >= 0.5) {
                    indexes.add(i);
                    return indexes;
                }
            }


        }
        return null;
    }

    private void findType () {
        String s;
        for (int i = 0; i< firstPage.size(); i++) {
            s = DocBase.getText(firstPage.get(i));
            if (ngrammPossibility(s, type) >= 0.5) {
                typeIndex = i;
                return;
            }
        }
    }

    private void findAgreementsAndApprove () {
        String s;
        for (int i = 0; i< firstPage.size(); i++) {
            s = DocBase.getText(firstPage.get(i)).toLowerCase();
            if (s.toLowerCase().equals("cогласовано") || ngrammPossibility("согласовано", s) >= 0.5) {
         //       DocBase.setHighlight(firstPage.get(i), "yellow");
                if (agreementIndexes[0]== -1)
                    agreementIndexes[0] = i;
                else {
                    agreementIndexes[1] = i;
                    return;
                }
            }

            if (s.toLowerCase().equals("утверждаю")|| ngrammPossibility("утверждаю", s) >= 0.5) {
            //    DocBase.setHighlight(firstPage.get(i), "yellow");
                approveIndex = i;
            }
        }

    }

    private void findNameOfCompany () {
        String s;
        StringBuffer name = new StringBuffer();
        for (int i = 0; i < agreementIndexes[0]; i++) {
            s = DocBase.getText(firstPage.get(i));
            name.append(s);
            name.append(" ");
        }
        nameOfCompany = name.toString();
    }

    private void setSubscribes () {
        String s;
        int lastIndex = (approveIndex == -1) ? indexName.get(0) : approveIndex;
        StringBuffer res = new StringBuffer();
        for (int i = agreementIndexes[0]; i < lastIndex; i++ )  {
            s = DocBase.getText(firstPage.get(i));
            if (s.toLowerCase().contains("согласовано") || ngrammPossibility("согласовано", s) >=0.5)
                s = "СОГЛАСОВАНО";
            res.append(s);
           res.append(" \n ");
        }
        agreement = res.toString();
        if (approveIndex != -1) {
            lastIndex = indexName.get(0);
            res = new StringBuffer();
            for (int i = approveIndex; i < lastIndex; i++ )  {
                s = DocBase.getText(firstPage.get(i));
                if (s.toLowerCase().contains("утверждаю")|| ngrammPossibility("утверждаю", s) >=0.5)
                    s = "УТВЕРЖДАЮ";
                res.append(s);
                res.append("\n");
            }
            approve = res.toString();
        }
    }
    private void setAlbom() {
        String s;
        for (int i = 0; i < firstPage.size(); i++ ) {
            s = DocBase.getText(firstPage.get(i));
            if (s.toLowerCase().contains("альбом")){
                albom = s;
                if (DocBase.getText(firstPage.get(i + 1)).toLowerCase().contains("всего альбомов")){
                    albom += " \n ";
                    albom += DocBase.getText(firstPage.get(i + 1));
                }
                return;
            }
        }
    }

    private void setDocNumber () {
        String regDocNumber = (GOSTnumber == 19)? "[А-ЯA-Z]+.\\d+.\\d+-\\d{2}.( ){1}\\d{2}-?\\d*-(ЛУ){1}" :
                "(\\d+.){2}\\d{3}.[А-Я]{1,2}\\d?.?\\d*.?\\d*-?\\d*.?M?-(ЛУ){1}";
        Pattern p = Pattern.compile(regDocNumber);
        String s;
        String[] temp;
        for (int i = 0; i < firstPage.size(); i++ ) {
            s = DocBase.getText(firstPage.get(i));
            if (p.matcher(s).matches()) {
                docNumber = s;
                docNumberIndex = i;
                return;
            }
            else {
                if (s.toLowerCase().contains("-лу")){
                    docNumber = s + "{wrong}";
                    docNumberIndex = i;
                    return;
                }
                temp = s.split("\\.");
                if (isDocNumberCorrect(temp)) {
                    docNumber = s + "{wrong}";
                    docNumberIndex = i;
                    return;
                }
            }
        }

    }

    private boolean isDocNumberCorrect (String[] strings) {
        Pattern p1 = Pattern.compile("[0-9-]{3,}");
        double coincedence = 0.0;
        for (String string : strings) {
            if (p1.matcher(string).matches() || string.toLowerCase().contains("лу"))
                coincedence++;
        }
        return (coincedence/strings.length >= 0.5)? true :false;
    }

    private void setPageNumber () {
        String s;
        for (int i = 0; i < firstPage.size(); i++) {
            s = DocBase.getText(firstPage.get(i));
            if (s.toLowerCase().contains("листов 1") || s.toLowerCase().contains("лист 1")) {
                pageNumberIndex = i;
                return;
            }
            if (s.toLowerCase().contains("листов") || s.toLowerCase().contains("лист")) {
                if (s.toLowerCase().replace("листов", "").matches("[0-9 ]+") ||
                        s.toLowerCase().replace("лист", "").matches("[0-9 ]+")) {
                    pageNumber = s;
                    pageNumberIndex = i;
                    return;
                }
            }

        }
    }

    private void setSubNameAndMedium () {
        String s;
        if (typeIndex != -1 && indexName.size() != 0 &
                typeIndex - indexName.get(indexName.size() - 1) > 1) {
            for (int i = indexName.get(indexName.size()); i < typeIndex; i++ ) {
                s = DocBase.getText(firstPage.get(i));
                if (s.matches("[А-Яа-яA-Za-z]+[ ]?[А-Яа-яA-Za-z ]+"))
                    subName = s;
            }
        }
        if (pageNumberIndex!= -1 && docNumberIndex!= -1 &
                pageNumberIndex - docNumberIndex > 1) {
            for (int i = docNumberIndex + 1; i < pageNumberIndex; i++ ) {
                s = DocBase.getText(firstPage.get(i));
                if (s.matches("[А-Яа-яA-Za-z]+[ ]?[А-Яа-яA-Za-z ]+"))
                    medium = s;
            }
        }
    }

    private void setRemained () {
        if (firstPage.size() - 1 == pageNumberIndex|| firstPage.size() - 1 == docNumberIndex ||
                firstPage.size() - 1 == index)
            return;
        remained = new ArrayList<>();
        int start = (agreementIndexes[1] == -1 ) ?
                (pageNumberIndex == -1) ?
                        (docNumberIndex == -1) ?
                                index +1 :
                                docNumberIndex +1 :
                        pageNumberIndex + 1 :
                agreementIndexes[1] +1;
        remained.addAll(firstPage.subList(start, firstPage.size() - 1));

    }




    public String getAgreement() {
        return agreement;
    }

    public String getApprove() {
        return approve;
    }

    public String getNameOfCompany() {
        return nameOfCompany;
    }

    public String getDocNumber() {
        return docNumber;
    }

    public String getPageNumber() {
        return pageNumber;
    }

    public String getMedium() {
        return medium;
    }

    public String getSubName() {
        return subName;
    }

    public List<P> getRemained() {
        return remained;
    }

    public boolean isSetType () {
        return !(typeIndex == -1);
    }

    public String getAlbom() {
        return albom;
    }

}
