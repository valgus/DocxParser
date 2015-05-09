package View;

import javafx.fxml.FXMLLoader;
import javafx.scene.Parent;
import javafx.scene.Scene;
import javafx.stage.Stage;

import java.awt.*;
import java.io.IOException;

public class Application extends javafx.application.Application {
    @Override
    public void start(Stage primaryStage) throws IOException {
        final Parent parent = FXMLLoader.load(getClass().getResource("/Application.fxml"));
        Dimension screenSize = Toolkit.getDefaultToolkit().getScreenSize();
        double width = screenSize.getWidth();
        double height = screenSize.getHeight();
        primaryStage.setScene(new Scene(parent,width/1.2, height/1.2));
        primaryStage.show();
    }
    public static void main(String[] args) {
        launch(args);
    }
}
