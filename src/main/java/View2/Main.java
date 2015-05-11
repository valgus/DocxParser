package View2;//import javafx.application.Application;
//import javafx.fxml.FXMLLoader;
//import javafx.scene.Scene;
//import javafx.scene.layout.Pane;
//import javafx.stage.Stage;
//import org.apache.fop.fo.flow.PageNumber;
//import org.docx4j.wml.ParaRPr;
//
//import java.awt.*;
//import java.io.IOException;
//
///**
// * Main application class.
// */
//public class Main extends Application {
//
//    Dimension screenSize = Toolkit.getDefaultToolkit().getScreenSize();
//    double width = screenSize.getWidth();
//    double height = screenSize.getHeight();
//    Stage currentStage;
//
//    @Override
//    public void start(Stage stage) throws Exception{
//
//        stage.setTitle("Vista Viewer");
//
//        stage.setScene(createScene(loadMainPane()));
//        currentStage = stage;
//        stage.setWidth(width / 2);
//        stage.setHeight(height / 1.1);
//        stage.show();
//
//        stage.show();
//    }
//
//    /**
//     * Loads the main fxml layout.
//     * Sets up the vista switching PaneNavigator.
//     * Loads the first vista into the fxml layout.
//     *
//     * @return the loaded pane.
//     * @throws IOException if the pane could not be loaded.
//     */
//    private Pane loadMainPane() throws IOException {
//        FXMLLoader loader = new FXMLLoader();
//
//        Pane mainPane =  loader.load(
//                getClass().getResourceAsStream(
//                        PaneNavigator.MAIN
//                )
//        );
//        MainController mainController = loader.getController();
//
//        PaneNavigator.setMainController(mainController);
//        PaneNavigator.loadPane(PaneNavigator.VISTA_1);
//
//        return mainPane;
//    }
//
//    /**
//     * Creates the main application scene.
//     *
//     * @param mainPane the main application layout.
//     *
//     * @return the created scene.
//     */
//    private Scene createScene(Pane mainPane) {
//        Scene scene = new Scene(
//                mainPane
//        );
//
//        scene.getStylesheets().setAll(

//
//        return scene;
//    }
//
//    public static void main(String[] args) {
//        launch(args);
//    }
//}

import javafx.application.Application;
import javafx.fxml.FXMLLoader;
import javafx.scene.Scene;
import javafx.scene.layout.Pane;
import javafx.stage.Stage;
import javafx.stage.StageStyle;

import java.awt.*;
import java.io.IOException;


public class Main extends Application {

    public void start(Stage primaryStage) throws IOException {
        final FXMLLoader  loader = new FXMLLoader(getClass().getResource("/app.fxml"));
        Dimension screenSize = Toolkit.getDefaultToolkit().getScreenSize();
        double width = screenSize.getWidth();
        double height = screenSize.getHeight();
        Scene scene = new Scene((Pane)loader.load(),width/2.5, height/2);
        scene.getStylesheets().setAll(getClass().getResource("/styles.css").toExternalForm());
        primaryStage.setScene(scene);
        primaryStage.initStyle(StageStyle.UNDECORATED);
        primaryStage.initStyle(StageStyle.TRANSPARENT);
        primaryStage.show();
    }
    public static void main(String[] args) {
        launch(args);
    }

}

