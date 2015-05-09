package View;

import com.sun.pdfview.PDFFile;
import com.sun.pdfview.PDFPage;
import javafx.beans.binding.Bindings;
import javafx.beans.binding.IntegerBinding;
import javafx.beans.property.DoubleProperty;
import javafx.beans.property.ObjectProperty;
import javafx.beans.property.SimpleDoubleProperty;
import javafx.beans.property.SimpleObjectProperty;
import javafx.beans.value.ChangeListener;
import javafx.beans.value.ObservableValue;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.concurrent.Task;
import javafx.concurrent.WorkerStateEvent;
import javafx.embed.swing.SwingFXUtils;
import javafx.event.*;
import javafx.fxml.FXML;
import javafx.fxml.FXMLLoader;
import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.Node;
import javafx.scene.Scene;
import javafx.scene.control.*;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.ScrollPane;
import javafx.scene.image.Image;
import javafx.scene.image.ImageView;
import javafx.scene.layout.BorderPane;
import javafx.scene.layout.HBox;
import javafx.scene.layout.VBox;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import javafx.stage.StageStyle;
import javafx.stage.Window;
import javafx.util.Callback;

import java.awt.geom.Rectangle2D;
import java.awt.image.BufferedImage;
import java.io.*;
import java.nio.ByteBuffer;
import java.nio.channels.FileChannel;
import java.nio.file.Paths;
import java.util.Collection;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.ThreadFactory;

public class Controller {
    @FXML private BorderPane borderPane1,borderPane2;
    @FXML private Pagination pagination_1;
    @FXML private Pagination pagination_2;
    @FXML private ComboBox gostChooser;
    @FXML private Button OK;
    @FXML private ComboBox docChooser;
    @FXML private TextField name;
    @FXML private Label stillDisabled;
    private FileChooser fileChooser ;
    private ObjectProperty<PDFFile> currentFile ;
    private ObjectProperty<ImageView> currentImage ;
    @FXML  private ScrollPane scroller ;
    private PageDimensions currentPageDimensions ;
    private ExecutorService imageLoadService ;
    private int gostNumber;
    private String doc;


    // ************ Initialization *************

    public void initialize() {
        gostNumber = 0;
        createAndConfigureImageLoadService();
        createAndConfigureFileChooser();
        currentFile = new SimpleObjectProperty<>();
        updateWindowTitleWhenFileChanges();
        currentImage = new SimpleObjectProperty<>();
        scroller.contentProperty().bind(currentImage);
        bindPaginationToCurrentFile();
        createPaginationPageFactory();
        ObservableList<String> gosts = FXCollections.observableArrayList();
        gostChooser.setValue("Choose..");
        gosts.add("19");
        gosts.add("34");
        gostChooser.setItems(gosts);
        OK.setOnAction(new EventHandler<ActionEvent>() {
            @Override
            public void handle(ActionEvent actionEvent) {
                if (doc!= null & !name.getText().equals("Here..") && !name.getText().matches("[ ]*")) {
                    //TODO
                    stillDisabled.setText("Все хорошо!");
                    try {
                        borderPane1.getChildren().setAll((Node)FXMLLoader.load(getClass().getResource("Application_2.fxml")));
                    } catch (IOException e) {
                        e.printStackTrace();
                    }
                }
                else {
                    stillDisabled.setText("Не введены все данные!");
                }
            }
        });
        name.focusedProperty().addListener(new ChangeListener<Boolean>() {
            @Override
            public void changed(ObservableValue<? extends Boolean> observableValue, Boolean oldBool, Boolean newBool) {
                if (newBool) {
                    stillDisabled.setText("");
                    name.setText("");
                }
            }
        });
    }

    private void createAndConfigureImageLoadService() {
        imageLoadService = Executors.newSingleThreadExecutor(new ThreadFactory() {
            @Override
            public Thread newThread(Runnable r) {
                Thread thread = new Thread(r);
                thread.setDaemon(true);
                return thread;
            }
        });
    }

    private void createAndConfigureFileChooser() {
        fileChooser = new FileChooser();
        fileChooser.setInitialDirectory(Paths.get(System.getProperty("user.home")).toFile());
        fileChooser.getExtensionFilters().add(new FileChooser.ExtensionFilter("PDF Files", "*.pdf", "*.PDF"));
    }

    private void updateWindowTitleWhenFileChanges() {
        currentFile.addListener(new ChangeListener<PDFFile>() {
            @Override
            public void changed(ObservableValue<? extends PDFFile> observable, PDFFile oldFile, PDFFile newFile) {
                try {
                    String title = newFile == null ? "PDF Viewer" : newFile.getStringMetadata("Title") ;
                    Window window = pagination_1.getScene().getWindow();
                    if (window instanceof Stage) {
                        ((Stage)window).setTitle(title);
                    }
                } catch (IOException e) {
                    showErrorMessage("Could not read title from pdf file", e);
                }
            }

        });
    }

    private void bindPaginationToCurrentFile() {
        currentFile.addListener(new ChangeListener<PDFFile>() {
            @Override
            public void changed(ObservableValue<? extends PDFFile> observable, PDFFile oldFile, PDFFile newFile) {
                if (newFile != null) {
                    pagination_1.setCurrentPageIndex(0);
                }
            }
        });
        pagination_1.pageCountProperty().bind(new IntegerBinding() {
            {
                super.bind(currentFile);
            }
            @Override
            protected int computeValue() {
                return currentFile.get()==null ? 0 : currentFile.get().getNumPages() ;
            }
        });
        pagination_1.disableProperty().bind(Bindings.isNull(currentFile));
    }

    private void createPaginationPageFactory() {
        pagination_1.setPageFactory(new Callback<Integer, Node>() {
            @Override
            public Node call(Integer pageNumber) {
                if (currentFile.get() == null) {
                    return null;
                } else {
                    if (pageNumber >= currentFile.get().getNumPages() || pageNumber < 0) {
                        return null;
                    } else {
                        updateImage(pageNumber);
                        return scroller;
                    }
                }
            }
        });
    }

    // ************** Event Handlers ****************

    @FXML private void loadFile() {
        final File file = fileChooser.showOpenDialog(pagination_1.getScene().getWindow());
        if (file != null) {
            final Task<PDFFile> loadFileTask = new Task<PDFFile>() {
                @Override
                protected PDFFile call() throws Exception {
                    try (
                            RandomAccessFile raf = new RandomAccessFile(file, "r");
                            FileChannel channel = raf.getChannel()
                    ) {
                        ByteBuffer buffer = channel.map(FileChannel.MapMode.READ_ONLY, 0, channel.size());
                        return new PDFFile(buffer);
                    }
                }
            };
            loadFileTask.setOnSucceeded(new EventHandler<WorkerStateEvent>() {
                @Override
                public void handle(WorkerStateEvent event) {
                    pagination_1.getScene().getRoot().setDisable(false);
                    final PDFFile pdfFile = loadFileTask.getValue();
                    currentFile.set(pdfFile);
                }
            });
            loadFileTask.setOnFailed(new EventHandler<WorkerStateEvent>() {
                @Override
                public void handle(WorkerStateEvent event) {
                    pagination_1.getScene().getRoot().setDisable(false);
                    showErrorMessage("Could not load file "+file.getName(), loadFileTask.getException());
                }
            });
            pagination_1.getScene().getRoot().setDisable(true);
            imageLoadService.submit(loadFileTask);
        }
    }
    @FXML private void showDifference () {

    }

    @FXML private void setName () {
        if (name.isPressed()) {
            stillDisabled.setText("");
            name.setText("");
        }

    }

    @FXML private void chooseGost() {
        stillDisabled.setText("");
        gostNumber = Integer.valueOf((String) gostChooser.getSelectionModel().getSelectedItem());
        docChooser.setDisable(false);
        ObservableList<String> docs = FXCollections.observableArrayList();
        switch (gostNumber) {
            case (19):
                docs.add("Формуляр");
                docs.add("Спецификация");
                docs.add("Ведомость держателей подлинников");
                docs.add("Текст программы");
                docs.add("Описание программы");
                docs.add("Ведомость эксплуатационных документов");
                docs.add("Описание применения");
                docs.add("Руководство системного программиста");
                docs.add("Руководство программиста");
                docs.add("Руководство оператора");
                docs.add("Описание языка");
                docs.add("Руководство по техническому обслуживанию");
                docs.add("Программа и методика испытаний");
                docs.add("Пояснительная записка");
                break;
            case (34):
                docs.add("Технологическая инструкция");
                docs.add("Схема функциональной структуры");
                docs.add("Схема структурная комплекса технических средств");
                docs.add("Схема организационной структуры");
                docs.add("Схема автоматизации");
                docs.add("Спецификация оборудования");
                docs.add("Руководство пользователя");
                docs.add("Проектная оценка надежности системы");
                docs.add("Перечень выходных сигналов (документов)");
                docs.add("Перечень входных сигналов и данных");
                docs.add("Паспорт");
                docs.add("Описание систем классификации и кодирования");
                docs.add("Описание программного обеспечения");
                docs.add("Описание постановки задач");
                docs.add("Описание организационной структуры");
                docs.add("Описание организации информационной базы");
                docs.add("Описание КТС");
                docs.add("Описание информационного обеспечения системы");
                docs.add("Описание автоматизируемых функций");
                docs.add("Описание технологического процесса обработки данных");
                docs.add("Описание проектной процедуры");
                docs.add("Описание алгоритма");
                docs.add("Общее описание системы");
                docs.add("Массив входных данных");
                docs.add("Каталог базы данных");
                docs.add("Инструкция по эксплуатации КТС");
                docs.add("Ведомость машинных носителей информации");
                docs.add("Ведомость эксплуатационных документов");
                docs.add("Формуляр");
                docs.add("Программа и методика испытаний");
                docs.add("Состав выходных данных");
                docs.add("Описание массива информации");
                docs.add("Ведомость оборудования и материалов");
                docs.add("Техническое задание");
                break;
        }
        docChooser.setItems(docs);
    }

    @FXML private void chooseDoc() {
        stillDisabled.setText("");
        doc = (String)docChooser.getSelectionModel().getSelectedItem();
    }

    @FXML private void setParameters() {



    }

    // *************** Background image loading ****************

    private void updateImage(final int pageNumber) {
        final Task<ImageView> updateImageTask = new Task<ImageView>() {
            @Override
            protected ImageView call() throws Exception {
                PDFPage page = currentFile.get().getPage(pageNumber+1);
                Rectangle2D bbox = page.getBBox();
                final double actualPageWidth = bbox.getWidth();
                final double actualPageHeight = bbox.getHeight();
                // record page dimensions for zoomToFit and zoomToWidth:
                currentPageDimensions = new PageDimensions(actualPageWidth, actualPageHeight);

                // width, height, clip, imageObserver, paintBackground, waitUntilLoaded:
                java.awt.Image awtImage = page.getImage((int)actualPageWidth, (int)actualPageHeight, bbox, null, true, true);
                // draw image to buffered image:
                BufferedImage buffImage = new BufferedImage((int)actualPageWidth, (int)actualPageHeight, BufferedImage.TYPE_INT_RGB);
                buffImage.createGraphics().drawImage(awtImage, 0, 0, null);
                // convert to JavaFX image:
                Image image = SwingFXUtils.toFXImage(buffImage, null);
                // wrap in image view and return:
                ImageView imageView = new ImageView(image);
                imageView.setPreserveRatio(true);
                return imageView ;
            }
        };

        updateImageTask.setOnSucceeded(new EventHandler<WorkerStateEvent>() {
            @Override
            public void handle(WorkerStateEvent event) {
                pagination_1.getScene().getRoot().setDisable(false);
                currentImage.set(updateImageTask.getValue());
            }
        });

        updateImageTask.setOnFailed(new EventHandler<WorkerStateEvent>() {
            @Override
            public void handle(WorkerStateEvent event) {
                pagination_1.getScene().getRoot().setDisable(false);
                updateImageTask.getException().printStackTrace();
            }

        });

        pagination_1.getScene().getRoot().setDisable(true);
        imageLoadService.submit(updateImageTask);
    }

    private void showErrorMessage(String message, Throwable exception) {

        // TODO: move to fxml (or better, use ControlsFX)

        final Stage dialog = new Stage();
        dialog.initOwner(pagination_1.getScene().getWindow());
        dialog.initStyle(StageStyle.UNDECORATED);
        final VBox root = new VBox(10);
        root.setPadding(new Insets(10));
        StringWriter errorMessage = new StringWriter();
        exception.printStackTrace(new PrintWriter(errorMessage));
        final Label detailsLabel = new Label(errorMessage.toString());
        TitledPane details = new TitledPane();
        details.setText("Details:");
        Label briefMessageLabel = new Label(message);
        final HBox detailsLabelHolder =new HBox();

        Button closeButton = new Button("OK");
        closeButton.setOnAction(new EventHandler<ActionEvent>() {
            @Override
            public void handle(ActionEvent event) {
                dialog.hide();
            }
        });
        HBox closeButtonHolder = new HBox();
        closeButtonHolder.getChildren().add(closeButton);
        closeButtonHolder.setAlignment(Pos.CENTER);
        closeButtonHolder.setPadding(new Insets(5));
        root.getChildren().addAll(briefMessageLabel, details, detailsLabelHolder, closeButtonHolder);
        details.setExpanded(false);
        details.setAnimated(false);

        details.expandedProperty().addListener(new ChangeListener<Boolean>() {

            @Override
            public void changed(ObservableValue<? extends Boolean> observable,
                                Boolean oldValue, Boolean newValue) {
                if (newValue) {
                    detailsLabelHolder.getChildren().add(detailsLabel);
                } else {
                    detailsLabelHolder.getChildren().remove(detailsLabel);
                }
                dialog.sizeToScene();
            }

        });
        final Scene scene = new Scene(root);

        dialog.setScene(scene);
        dialog.show();
    }


	/*
	 * Struct-like class intended to represent the physical dimensions of a page in pixels
	 * (as opposed to the dimensions of the (possibly zoomed) view.
	 * Used to compute zoom factors for zoomToFit and zoomToWidth.
	 *
	 */

    private class PageDimensions {
        private double width ;
        private double height ;
        PageDimensions(double width, double height) {
            this.width = width ;
            this.height = height ;
        }
        @Override
        public String toString() {
            return String.format("[%.1f, %.1f]", width, height);
        }
    }

}
