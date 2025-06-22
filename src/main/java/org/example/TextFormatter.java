package org.example;
import org.apache.poi.xwpf.usermodel.*;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class TextFormatter {

    public static void main(String[] args) {
        String inputFilePath = "input.docx";
        String outputFilePath = "output.docx";

        formatDocument(inputFilePath, outputFilePath);
    }

    public static void formatDocument(String inputFilePath, String outputFilePath) {
        try (XWPFDocument document = new XWPFDocument(new FileInputStream(inputFilePath))) {
            for (XWPFParagraph paragraph : document.getParagraphs()) {
                formatParagraph(paragraph);
            }

            try (FileOutputStream out = new FileOutputStream(outputFilePath)) {
                document.write(out);
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void formatParagraph(XWPFParagraph paragraph) {
        for (XWPFRun run : paragraph.getRuns()) {
            String text = run.getText(0);

            // Проверка на шрифт Courier New
            if ("Courier New".equals(run.getFontFamily())) {
                continue;
            }

            if (text != null) {
                String originalText = text; // Сохраняет оригинальный текст

                // Заменяет одиночные и двойные кавычки на бэктики
                text = text.replace("'", "`").replace("\"", "`");

                // Если текст был курсивом, обрамляет его в бэктики и убираем курсив
                if (run.isItalic()) {
                    text = "`" + text + "`"; // Обрамляет текст в бэктики
                    run.setItalic(false); // Убирает курсив
                }

                // Заменяет дефис и среднее тире на длинное тире
                text = text.replaceAll(" - ", " — ").replaceAll(" – ", " — ");

                // Убирает пробелы вокруг длинного тире только для английских текстов
                if (isEnglish(text)) {
                    text = text.replaceAll(" ?— ?", "—"); // Убирает пробелы вокруг длинного тире
                }

                // Обновляет текст
                run.setText(text, 0);

                // Устанавливает цвет текста в желтый, если текст изменен
                if (!text.equals(originalText)) {
                    run.setColor("FFFF00"); // Выделяет желтым
                }

                // Если текст жирный, выделяет желтым
                if (run.isBold()) {
                    run.setColor("FFFF00");
                }
            }
        }
    }

    private static boolean isEnglish(String text) {
        // Метод для проверки, содержит ли текст английские символы
        return text.matches(".*[A-Za-z].*");
    }
}