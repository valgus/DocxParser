package Model;

import org.docx4j.convert.out.common.preprocess.CoverPageSectPrMover;
import org.docx4j.convert.out.common.preprocess.ParagraphStylesInTableFix;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;

import java.io.File;
import java.io.FileNotFoundException;
import java.util.concurrent.BlockingQueue;

public class ProcessDocument implements Runnable{
    private final BlockingQueue<Object> messageQueue ;
    private final BlockingQueue<Boolean> messageQueue2 ;
    private Document doc;
    private File file;
    private String name;
    private String type;
    private Object sync;
    public ProcessDocument(Document doc, File file, String name, String type,
                           BlockingQueue<Object> messageQueue,BlockingQueue<Boolean> messageQueue2, Object sync) {
        this.name = name;
        this.file = file;
        this.doc = doc;
        this.type = type;
        this.messageQueue = messageQueue ;
        this.messageQueue2 = messageQueue2 ;
        this.sync = sync;
    }
    public void run() {
        try{
            MainPart comparisonWithTemplate = new MainPart(doc);
            messageQueue.put(15);
            comparisonWithTemplate.setTwoDocx(doc.getTemplate(),file, this);

            messageQueue.put(20);
            WordprocessingMLPackage word;
            try {
                word = DocxMethods.getTemplate(file.getAbsolutePath());
            } catch (Docx4JException e) {
                messageQueue.put("rerun process");
                return;
            } catch (FileNotFoundException e) {
                messageQueue.put("rerun process");
                return;
            }
            if (word.getMainDocumentPart().getContent().size() == 0) {
                messageQueue.put("Docx is empty");
                messageQueue.put("exit");
                return;
            }
            messageQueue.put(30);
            StyleSetter styleSetter = new StyleSetter(word);
            styleSetter.setStyle();
            EditingFirstPages editingFirstPages = new EditingFirstPages(word, type ,doc.getGost(), name, this);
            try {
                word = editingFirstPages.process();
                messageQueue.put(65);
            } catch (Exception e) {
                messageQueue.put("Проблема при обработке документа.");
                messageQueue.put("exit");
                return;
            }
            if (word == null) {
                messageQueue.put("exit");
                return;
            }
            word = comparisonWithTemplate.setAppropriateText(word);
            if (word == null) {
                messageQueue.put("exit");
                return;
            }
            messageQueue.put(80);
            CoverPageSectPrMover.process(word);
            messageQueue.put(90);
            ParagraphStylesInTableFix.process(word);
            messageQueue.put(100);
            messageQueue.put(word);
        } catch (InterruptedException e) {
            e.printStackTrace();
        }
    }

    public boolean sendInfo(Object info) {
        try {
            if (info instanceof Integer){
                messageQueue.put(info);
                return true;
            }
            synchronized (sync) {
                messageQueue.put(info);
                sync.wait();
            }
            Boolean res = messageQueue2.poll();
            if (res!= null)
                return res;
            return false;
        } catch (InterruptedException e) {
            return false;
        }
    }

}
