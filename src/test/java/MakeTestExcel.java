import jp.classicorange.utils.ExcelUtils;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;

import java.io.File;
import java.io.FileOutputStream;
import java.util.Random;

/**
 * テスト用のExcelファイルを生成
 * ディレクトリ ３個、ファイル5個ずつ生成
 *
 */
public class MakeTestExcel {
    public static void main(String[] args) {
        try {
            Random random = new Random();
            String projectRoot = System.getProperty("user.dir");
            for (int dirNum = 1; dirNum <= 3; dirNum++) {
                String dirPath = String.format("%s/testData/dir%02d/", projectRoot, dirNum);
                new File(dirPath).mkdirs();

                for (int fileNum = 1; fileNum <= 5; fileNum++) {
                    XSSFWorkbook workbook = new XSSFWorkbook();

                    //シートごと
                    for (int sheetNum = 1; sheetNum <= 3; sheetNum++) {
                        XSSFSheet sheet = workbook.createSheet("Sheet"+sheetNum);

                        // セルにデータ + ランダムで「Jakarta」
                        for (int rowNum = 0; rowNum < 1000; rowNum++) {
                            Row row = sheet.createRow(rowNum);
                            for (int colNum = 0; colNum < 10; colNum++) {
                                //セル
                                String cellValue = String.format("DataDir%02d File%02d Sheet%02d R%dC%d", dirNum, fileNum, sheetNum, rowNum + 1, colNum + 1);
                                if (random.nextInt(100) < 2) {  // 約2%の確率でJakartaを追記 検索対象ワード
                                    cellValue += " Jakarta";
                                }
                                row.createCell(colNum).setCellValue(cellValue);
                            }
                        }

                        // シェイプを15個作成 半透明
                        XSSFDrawing drawing = sheet.createDrawingPatriarch();
                        for (int s = 0; s < 15; s++) {
                            int col = random.nextInt(15);  // 0～14列目あたりに
                            int row = random.nextInt(50); // 上50行くらいに
                            XSSFClientAnchor anchor = new XSSFClientAnchor();
                            // シェイプの位置
                            anchor.setCol1(col);
                            anchor.setRow1(row);
                            anchor.setCol2(col+1);
                            anchor.setRow2(row+3);
                            anchor.setDx1(0);
                            anchor.setDy1(0);
                            anchor.setDx2(300 * 9525);
                            anchor.setDy2(200 * 9525);

                            // テキストあるシェイプを生成
                            XSSFTextBox textBox = drawing.createTextbox(anchor);
                            //textBox.setFillColor(128,255,128);

                            String shapeText = String.format("シェイプD%02dF%02d S%02d", dirNum, fileNum, s+1);
                            if (random.nextBoolean()) {  // 50%でJakarta含める
                                shapeText += " Jakarta";
                            }
                            textBox.setText(shapeText);

                            int alpha = 128; // 0 (透明) から 255 (不透明) までの値
                            ExcelUtils.setFillColorWithAlpha(textBox, 128,255,128, alpha);
                        }
                    }

                    String filename = String.format(dirPath + "テスト用D%02dF%02d.xlsx", dirNum, fileNum);
                    try (FileOutputStream out = new FileOutputStream(filename)) {
                        workbook.write(out);
                    }
                    workbook.close();
                    System.out.println(filename + " を作成しました！");
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }



}
