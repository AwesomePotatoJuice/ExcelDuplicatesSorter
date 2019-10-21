package surin;

import javafx.application.Application;
import javafx.scene.Scene;
import javafx.scene.layout.StackPane;
import javafx.scene.text.Text;
import javafx.stage.Stage;

public class FXController extends Application {
    @Override
    public void start(Stage primaryStage) throws Exception {
        Text helloWorld = new Text("Hello world");

        StackPane root = new StackPane(helloWorld);
        primaryStage.setScene(new Scene(root, 300, 120));
        primaryStage.centerOnScreen();
        primaryStage.show();
    }
    public static void main(String[] args) {
        launch(FXController.class, args);
    }
}
