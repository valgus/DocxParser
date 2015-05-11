package View2;

import Model.Document;
import javafx.animation.AnimationTimer;
import javafx.beans.property.LongProperty;
import javafx.beans.property.SimpleLongProperty;
import javafx.beans.value.ChangeListener;
import javafx.beans.value.ObservableValue;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.event.ActionEvent;
import javafx.event.EventHandler;
import javafx.fxml.FXML;
import javafx.scene.control.*;
import javafx.scene.input.MouseEvent;
import javafx.scene.layout.BorderPane;
import javafx.stage.FileChooser;
import javafx.stage.Stage;

import java.io.File;
import java.io.IOException;
import java.nio.file.Paths;
import java.util.Optional;
import java.util.concurrent.ArrayBlockingQueue;
import java.util.concurrent.BlockingQueue;

import Model.ProcessDocument;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;


public class Vista1Controller {

    @FXML private BorderPane root;
    @FXML private BorderPane border;
    @FXML private Button loadFile, close, minimize;
    @FXML private ComboBox gostChooser;
    @FXML private Button process;
    @FXML private ComboBox docChooser;
    @FXML private TextField name;
    @FXML private Label stillDisabled, procent;
    @FXML private ProgressBar progress;
    @FXML private Label docName;
    private File file;
    private FileChooser fileChooser ;
    private int gostNumber;
    private String doc;
    private Document document;
    private Vista1Controller controller = this;
    final BlockingQueue<Object> messageQueue = new ArrayBlockingQueue<>(1);
    private static double xOffset = 0;
    private static double yOffset = 0;
    // ************ Initialization *************

    public void initialize() {

        gostNumber = 0;
        createAndConfigureFileChooser();
        ObservableList<String> gosts = FXCollections.observableArrayList();
        gostChooser.setValue("Выбрать...");
        gosts.add("19");
        gosts.add("34");
        gostChooser.setItems(gosts);
        process.setOnAction(new EventHandler<ActionEvent>() {
            @Override
            public void handle(ActionEvent actionEvent) {
                if (doc!= null & !name.getText().equals("Here..") && !name.getText().matches("[ ]*") && file!= null) {
                    //TODO
                    stillDisabled.setText("Все хорошо!");
                    progress.setVisible(true);
                    procent.setVisible(true);
                    ProcessDocument processDocument = new ProcessDocument(document, file, name.getText(), controller,
                            doc, messageQueue);
                    Thread app = new Thread(processDocument);
                    app.start();

                    final LongProperty lastUpdate = new SimpleLongProperty();

                    final long minUpdateInterval = 0 ; // nanoseconds. Set to higher number to slow output.

                    AnimationTimer timer = new AnimationTimer() {

                        @Override
                        public void handle(long now) {
                            if (now - lastUpdate.get() > minUpdateInterval) {
                                final Object message = messageQueue.poll();
                                if (message instanceof Integer) {
                                    if (message != null) {
                                            procent.setText((int)message + "%");
                                            progress.setProgress((double)((int)message)/100);
                                    }
                                }
                                else if (message instanceof WordprocessingMLPackage){
                                        setWord((WordprocessingMLPackage)message);
                                    }
                                lastUpdate.set(now);

                            }
                        }

                    };

                    timer.start();


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
        root.setOnMousePressed(new EventHandler<MouseEvent>() {
            @Override
            public void handle(MouseEvent event) {
                xOffset = root.getScene().getWindow().getX() - event.getScreenX();
                yOffset = root.getScene().getWindow().getY() - event.getScreenY();
            }
        });

        root.setOnMouseDragged(new EventHandler<MouseEvent>() {
            @Override
            public void handle(MouseEvent event) {
                root.getScene().getWindow().setX(event.getScreenX() + xOffset);
                root.getScene().getWindow().setY(event.getScreenY() + yOffset);
            }
        });

        minimize.setOnMouseClicked(new EventHandler<MouseEvent>() {
            public void handle(MouseEvent me) {
                ((Stage)root.getScene().getWindow()).setIconified(true);
            }
        });
    }
    private void createAndConfigureFileChooser() {
        fileChooser = new FileChooser();
        fileChooser.setInitialDirectory(Paths.get(System.getProperty("user.home")).toFile());
        fileChooser.getExtensionFilters().add(new FileChooser.ExtensionFilter("DOCX Files", "*.docx", "*.DOCX"));
    }
    // ************** Event Handlers ****************

    @FXML private void loadFile() throws IOException {
        file = fileChooser.showOpenDialog(border.getScene().getWindow());
        if (file !=  null) {
            docName.setText(file.getName());
            docName.setVisible(true);
        }
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
                docs.add("Техническое задание");
                docs.add("Формуляр");
                docs.add("Спецификация");
                docs.add("Текст программы");
                docs.add("Описание программы");
                docs.add("Ведомость эксплуатационных документов");
                docs.add("Описание применения");
                docs.add("Руководство системного программиста");
                docs.add("Руководство программиста");
                docs.add("Руководство оператора");
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
                docs.add("Формуляр");
                docs.add("Программа и методика испытаний");
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
        switch (doc) {
            case("Формуляр") :
                if (gostNumber == 19) {
                    document = Document.Formulyar;
                }
                else {
                    document = Document.Formulyar_34;
                }
                break;
            case("Спецификация") : document = Document.Spezifikation; break;
            case("Текст программы") : document = Document.Tekst_programmi; break;
            case("Описание программы") : document = Document.Opisanie_Programmi; break;
            case("Ведомость эксплуатационных документов") :document = Document.Vedomost_ekspluatacionnich_dokumentov;
               break;
            case("Описание применения") :  document = Document.Opisanie_primeneniya;break;
            case("Руководство системного программиста") :  document = Document.Rukovodstvo_sistemnogo_programmista;break;
            case("Руководство программиста") : document = Document.Rukovodstvo_programmista; break;
            case("Руководство оператора") :  document = Document.Rukovodstvo_operatora;break;
            case("Руководство по техническому обслуживанию") :
                document = Document.Rukovodstvo_po_technicheskomu_obsluzhivaniu; break;
            case("Программа и методика испытаний") :
                if (gostNumber == 19) {
                    document = Document.Programma_i_metodika_ispitanii;
                }
                else {
                    document = Document.Programma_i_metodika_ispitanii_34;
                }break;
            case("Пояснительная записка") :  document = Document.Poiasnitelnaya_zapiska;break;
            case("Технологическая инструкция") :  document = Document.TEchnologicheskaya_instrukcia;break;
            case("Схема функциональной структуры") : document = Document.Schema_funkcion_strukturi; break;
            case("Схема структурная комплекса технических средств") :
                document = Document.Schema_struct_kompleksa_tech_sredstv; break;
            case("Схема организационной структуры") : document = Document.Schema_organiz_structuri; break;
            case("Схема автоматизации") : document = Document.Schema_avtomatizacii;break;
            case("Руководство пользователя") : document = Document.Rukovodstvo_polzovatelya; break;
            case("Проектная оценка надежности системы") : document = Document.Proektnaya_ocenka_nadeznosti_systemy;
                break;
            case("Перечень выходных сигналов (документов)") : document = Document.Perechen_vichodnih_signalov; break;
            case("Перечень входных сигналов и данных") : document = Document.Perechen_vchodnih_signalov; break;
            case("Паспорт") : document = Document.Pasport; break;
            case("Описание систем классификации и кодирования") :
                document = Document.Opisanie_system_klassifik_i_kodir;break;
            case("Описание программного обеспечения") : document = Document.Opisanie_program_obespecheniya; break;
            case("Описание постановки задач") : document = Document.Opisanie_postanovki_zadach; break;
            case("Описание организационной структуры") : document = Document.Opisanie_organ_structuri; break;
            case("Описание организации информационной базы") : document = Document.Opisanie_organ_inf_bazi; break;
            case("Описание КТС") :  document = Document.Opisanie_KTS; break;
            case("Описание информационного обеспечения системы") : document = Document.Opisanie_inf_onespech_systemi; break;
            case("Описание автоматизируемых функций") :  document = Document.Opisanie_avtomat_funkcii;break;
            case("Описание технологического процесса обработки данных") : document = Document.Opisanie_tech_processa; break;
            case("Описание проектной процедуры") : document = Document.Opisanie_proektnoi_proceduri; break;
            case("Описание алгоритма") :  document = Document.Opisanie_algoritma; break;
            case("Общее описание системы") :  document = Document.Opisanie_obcee_opisanie_systemi;break;
            case("Массив входных данных") : document = Document.Massiv_vhodnich_dannich; break;
            case("Каталог базы данных") : document = Document.Katalog_BD; break;
            case("Инструкция по эксплуатации КТС") : document = Document.Instrukciya_po_ekspluat_AS; break;
            case("Описание массива информации") :  document = Document.Opisanie_massiva_informacii;break;
            case("Техническое задание") :
                if (gostNumber == 19) {
                    document = Document.Tecnicheskoe_zadanie;
                }
                else {
                    document = Document.Tecnicheskoe_zadanie_34;
                }break;




        }
    }

    @FXML private void setParameters() {



    }

    public void close(ActionEvent actionEvent) {
        Stage stage = (Stage) close.getScene().getWindow();
        stage.close();
    }

//    /**
//     * @param args the command line arguments
//     */
//    public static void main(String[] args) {
//        Application.launch(args);
//    }
//
//    @Override
//    public void start(Stage primaryStage) {
//        primaryStage.setTitle("Hello World");
//        Group root = new Group();
//        Scene scene = new Scene(root, 300, 250);
//        Button btn = new Button();
//        btn.setLayoutX(100);
//        btn.setLayoutY(80);
//        btn.setText("Hello World");
//        btn.setOnAction(new EventHandler<ActionEvent>() {
//
//            public void handle(ActionEvent event) {
//                String[] cmdArray = {"cmd", "/c", "start", "c:\\q.docx"};
//
//                try {
//                    java.lang.Runtime.getRuntime().exec(cmdArray);
//                } catch (Exception s) {
//                }
//
//            }
//        });
//        root.getChildren().add(btn);
//        primaryStage.setScene(scene);
//        primaryStage.show();
//    }


    public boolean setInfo(String info) {
        String message, answer_1, answer_2 = null;
        switch (info) {
            case ("Файл не соответствует шаблону!"):
                message = "Загруженный файл не соответствует структуре файла \"" + docName + "\".";
                answer_1 = "Остановить процесс";
                break;
            case ("Docx is empty"):
                message = "Загруженный файл пуст. Процесс завершается.";
                answer_1 = "ОК";
                break;
            case ("year or letter must be set"):
                message = "Необходимо наличие даты или литеры на листе утверждения и(или) титульном листе." +
                        "Создать пустые лист утверждения и титульный лист?";
                answer_1 = "Да";
                answer_2 = "Остановить процесс";
                break;
            case ("Is not the first page!"):
                message = "В файле отсутствует лист утверждения. Создать пустой лист утверждения?";
                answer_1 = "Да";
                answer_2 = "Остановить процесс";
                break;
            case ("The inserted name is not correct"):
                message = "Введенное название документа некорректно. Процесс завершается.";
                answer_1 = "ОК";
                break;
            default: message = "Произошла внутрисистемная ошибка. Попробовать снова?";
                answer_1 = "Да";
                answer_2 = "Нет";
                break;
        }
        return answer_1.equals(showErrorMessage(message, answer_1, answer_2));
    }

    private String showErrorMessage(String message, String answer_1, String answer_2) {

        Alert alert = new Alert(Alert.AlertType.CONFIRMATION);
        alert.setTitle("");
        // alert.setHeaderText("Look, a Confirmation Dialog with Custom Actions");
        alert.setContentText(message);

        ButtonType buttonTypeOne = new ButtonType(answer_1);
        ButtonType buttonTypeTwo = null;
        if (answer_2!= null) {
            buttonTypeTwo = new ButtonType(answer_2);
            alert.getButtonTypes().setAll(buttonTypeOne, buttonTypeTwo);
        }
        else {

            alert.getButtonTypes().setAll(buttonTypeOne);
        }


        Optional<ButtonType> result = alert.showAndWait();
        if (result.get() == buttonTypeOne){
            return answer_1;
        } else if ( buttonTypeTwo!= null && result.get() == buttonTypeTwo) {
            return answer_2;
        } else {
            return answer_1;
        }
    }

    public  void setWord(WordprocessingMLPackage word) {
        if (word == null)
            showErrorMessage("Файл пуст.", "ОК", null);

        File file = fileChooser.showSaveDialog(root.getScene().getWindow());
        try {
            word.save(file);
            String[] cmdArray = {"cmd", "/c", "start", file.getAbsolutePath()};
            java.lang.Runtime.getRuntime().exec(cmdArray);
        } catch (Docx4JException e) {
            showErrorMessage("Не получилось сохранить в файл" + file.getName() + ".", "ОК", null);
        } catch (IOException e) {
            showErrorMessage("Не получается открыть приложение для демонстрации файла.", "ОК", null);
        }
    }
}