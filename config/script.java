// MyJavaScraper.java
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;

public class MyJavaScraper {
    public static void main(String[] args) {
        try {
            // Lee el contenido de la página desde el archivo temporal
            String content = new String(Files.readAllBytes(Paths.get("temp_page.html")));

            // Procesa el contenido de la página aquí (puedes usar cualquier lógica de scraping en Java)
            // ...

            System.out.println("Contenido procesado:");
            System.out.println(content);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}