package jp.classicorange;


import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;

import java.io.*;
import java.nio.file.*;

public class SearchShape {


    public void search(String keyword, String dirPath) throws IOException {

        Path dir = Paths.get(dirPath);
        Files.walk(dir)
             .filter(p -> p.toString().endsWith(".xlsx"))
             .forEach(p -> {
                 try (InputStream is = Files.newInputStream(p);
                      XSSFWorkbook workbook = new XSSFWorkbook(is)) {

                     for (Sheet sheet : workbook) {
                         if (sheet instanceof XSSFSheet xssfSheet) {
                             XSSFDrawing drawing = xssfSheet.getDrawingPatriarch();
                             if (drawing != null) {
                                 for (XSSFShape shape : drawing.getShapes()) {
                                     if (shape instanceof XSSFSimpleShape) {
                                         String text = ((XSSFSimpleShape) shape).getText();
                                         if (text != null && text.contains(keyword)) {
                                             System.out.println("★ 見つかりました: " + p + " → シート: " + sheet.getSheetName());
                                         }
                                     }
                                 }
                             }
                         }
                     }

                 } catch (IOException e) {
                     System.err.println("エラー: " + p + " - " + e.getMessage());
                 }
             });
    }

}
