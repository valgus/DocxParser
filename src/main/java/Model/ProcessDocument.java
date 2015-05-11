package Model;

import Model.Document;
import Model.EditingFirstPages;
import Model.MainPart;
import View2.Vista1Controller;
import org.docx4j.convert.out.common.preprocess.CoverPageSectPrMover;
import org.docx4j.convert.out.common.preprocess.ParagraphStylesInTableFix;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;

import java.io.File;
import java.util.concurrent.BlockingQueue;

public class ProcessDocument implements Runnable {
    private final BlockingQueue<Object> messageQueue ;

    private Vista1Controller controller;
    private Document doc;
    private File file;
    private String name;
    private String type;
    public ProcessDocument(Document doc, File file, String name, Vista1Controller controller, String type,
                           BlockingQueue<Object> messageQueue) {
        this.name = name;
        this.file = file;
        this.doc = doc;
        this.controller = controller;
        this.type = type;
        this.messageQueue = messageQueue ;
    }
    @Override
    public void run() {
        MainPart comparisonWithTemplate = new MainPart();
        comparisonWithTemplate.setTwoDocx(doc.getTemplate(),file, this);
        try {
            messageQueue.put(10);
            WordprocessingMLPackage word = comparisonWithTemplate.setAppropriateText();
            messageQueue.put(30);
            EditingFirstPages editingFirstPages = new EditingFirstPages(word, type ,doc.getGost(), name, this);
            try {
                word = editingFirstPages.process();
                messageQueue.put(60);
            } catch (Exception e) {
                sendInfo(e.getMessage());
            }
            if (word == null) {
                messageQueue.put(null);
                return;
            }
            CoverPageSectPrMover.process(word);
            messageQueue.put(80);
            ParagraphStylesInTableFix.process(word);
            messageQueue.put(100);
            messageQueue.put(word);
        } catch (InterruptedException e) {
            e.printStackTrace();
        }
    }

    public boolean sendInfo(String info) {
        return controller.setInfo(info);
    }

}
