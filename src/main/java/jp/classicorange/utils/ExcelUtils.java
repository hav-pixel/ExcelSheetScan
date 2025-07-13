package jp.classicorange.utils;

import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.util.List;

import javax.imageio.ImageIO;

import jp.classicorange.utils.entity.Size;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.ClientAnchor.AnchorType;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.util.Units;
import org.apache.poi.xssf.usermodel.XSSFSimpleShape;
import org.openxmlformats.schemas.drawingml.x2006.main.CTPositiveFixedPercentage;
import org.openxmlformats.schemas.drawingml.x2006.main.CTSRgbColor;
import org.openxmlformats.schemas.drawingml.x2006.main.CTShapeProperties;
import org.openxmlformats.schemas.drawingml.x2006.main.CTSolidColorFillProperties;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * Excel ユーティリティ
 * <p>
 * use POI 4.1.1
 *
 * @author Hav-pixel
 */
public interface ExcelUtils {

    /**
     * ロガー
     */
    Logger log = LoggerFactory.getLogger(ExcelUtils.class);


    /**
     * Workbookの97形式か返す
     *
     * @return boolean
     */
    static boolean is97(Workbook wb) {
        //xlsx 形式
        return wb.getSpreadsheetVersion() == SpreadsheetVersion.EXCEL97;
    }

    /**
     * Workbookの2007形式か返す
     *
     * @return boolean
     */
    static boolean is2007(Workbook wb) {
        //xlsx 形式
        return wb.getSpreadsheetVersion() == SpreadsheetVersion.EXCEL2007;
    }


    /**
     * Excelのアドレスを返す
     * A1形式
     * 形式 $sheet1.$A$1 $Sheet1.$A$1:$B$2など
     * <p>
     * POI 'Sheet1'!$A$1:$A$3
     */
    static String getExcelAddress(String sheetName, int col1, int row1, int col2, int row2) {
        //("Excel 用 26進数 A=0 Z=25");
        String res1 = getStrXlsAddress(col1, row1);
        String res2 = ":" + getStrXlsAddress(col2, row2);
        String res = "'" + sheetName + "'!" + res1 + res2;    //r1
        log.trace("ExcelAddress data sheetName={}, col1={},row1={},col2={},row2={} [{}]",
            sheetName, col1, row1, col2, row2, res);

        return res;
    }

    /**
     * Excel $A$1形式で返す
     *
     * @param col 0 is "A" , 25 is "Z"
     * @param row same.
     */
    static String getStrXlsAddress(int col, int row) {
        //digit 桁
        final int d = 26;    //桁数
        int d1 = col % d;
        int d10 = ((col - d1) - d) / d;
        String s1 = d10 < 0 ? "" : "" + (char) (65 + d10);
        String s2 = "" + (char) (65 + d1);
        return "$" + s1 + s2 + "$" + (row + 1);
    }


    /**
     * Cell object を String で返す。
     * null は 0文字返す。
     */
    static String getCellVal(Cell cell) {

        if (cell == null) return "";

        String formula;
        CellType type = cell.getCellType();

        formula = switch (type) {
            case STRING -> cell.getStringCellValue();
            case NUMERIC -> cell.getNumericCellValue() + "";
            case BOOLEAN -> cell.getBooleanCellValue() + "";
            case FORMULA -> cell.getCellFormula();
            case ERROR -> cell.getErrorCellValue() + "";
            //case BLANK  : formula = ""; break;
            default -> "";
        };

        return formula;

    }




    /**
     * 画像のサイズ(px)を返す。
     *
     * @param path 測定したい画像のパス
     * @return Size
     */
    static Size getPicureSize(String path) throws Exception {
        BufferedImage img = ImageIO.read(new File(path));
        return new Size(img.getWidth(), img.getHeight());
    }



    /**
     * シェイプの指定タイプからAnchorTypeを返す
     *
     * @param type 仮
     * @return AnchorType
     */
    static AnchorType getAnchorType(int type) {
        return switch (type) {
            case 0 -> AnchorType.DONT_MOVE_AND_RESIZE;
            case 1 -> AnchorType.MOVE_AND_RESIZE;
            case 2 -> AnchorType.MOVE_DONT_RESIZE;
            case 3 -> AnchorType.DONT_MOVE_DO_RESIZE;
            default -> null;
        };
    }




    /**
     * ワークブックにイメージをロードし、イメージへのインデックスを返す。
     *
     * @param        path    イメージのパス
     * @param        wb        HSSFWorkbook
     * @return イメージへのインデックス
     **/
    static int loadPicture(String path, Workbook wb) {
        int pictureIndex = 0;
        try (
            FileInputStream fis = new FileInputStream(path);
            ByteArrayOutputStream bos = new ByteArrayOutputStream()
        ) {
            int c;
            while ((c = fis.read()) != -1) {
                bos.write(c);
            }

            pictureIndex = wb.addPicture(bos.toByteArray(), Workbook.PICTURE_TYPE_JPEG);

        } catch (Exception e) {
            log.error(e.getMessage());
        }
        log.info("picture added {} : {}", pictureIndex, path);
        return pictureIndex;
    }



    /**
     * シート上の絶対位置xの、位置を示すカラム数とセル内X座標を返す
     *
     * @param sheet    測定対象シート
     * @param xInPixel シート上x軸の位置
     * @return int[] [0]:カラムの個数、[1]:セル内の X座標 emu
     */
    static int[] getAnchorX(Sheet sheet, int xInPixel) {
        int[] a;
        if (is97(sheet.getWorkbook())) {
            a = getHSSFAnchorX(sheet, xInPixel);
        } else {
            a = getXSSFAnchorX(sheet, xInPixel);
        }
        return a;
    }

    /**
     * シート上の絶対位置yの、位置を示すカラム数とセル内Y座標を返す
     *
     * @param sheet    測定対象シート
     * @param yInPixel シート上y軸の位置
     * @return int[] [0]:ローの個数 、[1]:セル内の Y座標 emu
     */
    static int[] getAnchorY(Sheet sheet, int yInPixel) {
        int[] a;
        if (is97(sheet.getWorkbook())) {
            a = getHSSFAnchorY(sheet, yInPixel);
        } else {
            a = getXSSFAnchorY(sheet, yInPixel);
        }
        return a;
    }



    /**
     * シート上の絶対位置yの、位置を示すカラム数とセル内Y座標を返す
     * for XSSF
     *
     * @param sheet    測定対象シート
     * @param xInPixel セル内x軸の位置
     * @return int[] [0]:カラムの個数、[1]:セル内の X座標 emu
     */
    private static int[] getXSSFAnchorX(Sheet sheet, int xInPixel) {
        int xInEMU = xInPixel * Units.EMU_PER_PIXEL;
        int currentInXEMU = 0;

        Row row = sheet.getRow(0);
        if (row == null) {
            row = sheet.createRow(0);
        }
        for (int columnNum = 0; ; columnNum++) {
            if (row.getCell(columnNum) == null) {
                row.createCell(columnNum);
            }

            if (sheet.getColumnWidth(columnNum) == sheet.getDefaultColumnWidth() * 256) {
                // Excelのデフォルト値は違うので明示的に設定する
                sheet.setColumnWidth(columnNum, 9 * 256);
            }

            // Excelでは256columnWidthが8ピクセル
            int columnWidthInEMU = sheet.getColumnWidth(columnNum) / 256 * 8 * Units.EMU_PER_PIXEL;
            if (xInEMU < currentInXEMU + columnWidthInEMU) {
                return new int[]{columnNum, xInEMU - currentInXEMU};
            }
            currentInXEMU += columnWidthInEMU;
        }
    }


    /**
     * シート上の絶対位置yの、位置を示すカラム数とセル内Y座標を返す
     * for HSSF
     *
     * @param sheet    測定対象シート
     * @param xInPixel セル内y軸の位置
     * @return int[] [0]:ローの個数 、[1]:セル内の Y座標 emu
     */
    private static int[] getHSSFAnchorX(Sheet sheet, int xInPixel) {

        log.error("TODO : ★ 修正");

        //TODO 計算方法がおかしいようだ。

        int xInEMU = xInPixel * Units.EMU_PER_PIXEL;
        int currentInXEMU = 0;

        Row row = sheet.getRow(0);
        if (row == null) {
            row = sheet.createRow(0);
        }
        for (int columnNum = 0; ; columnNum++) {
            if (row.getCell(columnNum) == null) {
                row.createCell(columnNum);
            }

            if (sheet.getColumnWidth(columnNum) == sheet.getDefaultColumnWidth() * 256) {
                // Excelのデフォルト値は違うので明示的に設定する
                sheet.setColumnWidth(columnNum, 9 * 256);
            }

            // Excelでは256columnWidthが8ピクセル
            int columnWidthInEMU = sheet.getColumnWidth(columnNum) / 256 * 8 * Units.EMU_PER_PIXEL;
            double colPix = sheet.getColumnWidthInPixels(columnNum);
            if (xInEMU < currentInXEMU + columnWidthInEMU) {
                return new int[]{columnNum, xInEMU - currentInXEMU};
            }
            currentInXEMU += columnWidthInEMU;
        }
    }



    /**
     * シート上の絶対位置yの、位置を示すカラム数とセル内Y座標を返す
     * for XSSF
     *
     * @param sheet    測定対象シート
     * @param yInPixel セル内y軸の位置
     * @return int[] [0]:ローの個数 、[1]:セル内の Y座標 emu
     */
    private static int[] getXSSFAnchorY(Sheet sheet, int yInPixel) {
        int yInEMU = yInPixel * Units.EMU_PER_PIXEL;
        int currentYInEMU = 0;

        for (int rowNum = 0; ; rowNum++) {
            Row currentRow = sheet.getRow(rowNum);
            if (currentRow == null) {
                currentRow = sheet.createRow(rowNum);
            }

            if (currentRow.getHeight() == sheet.getDefaultRowHeight()) {
                // Excelのデフォルト値は違うので明示的に設定する
                currentRow.setHeightInPoints(13.5f);
            }

            // 行の高さ（EMU）
            // ※row.getHeight()は1/20ポイント単位
            int rowHeightInEMU = (int) (currentRow.getHeight() / 20.0D * Units.EMU_PER_POINT);
            if (yInEMU < currentYInEMU + rowHeightInEMU) {
                return new int[]{rowNum, (yInEMU - currentYInEMU)};
            }
            currentYInEMU += rowHeightInEMU;
        }
    }


    /**
     * シート上の絶対位置yの、位置を示すカラム数とセル内Y座標を返す
     * for HSSF
     *
     * @param sheet    測定対象シート
     * @param yInPixel セル内y軸の位置
     * @return int[] [0]:ローの個数 、[1]:セル内の Y座標 emu
     */
    @Deprecated
    private static int[] getHSSFAnchorY(Sheet sheet, int yInPixel) {
        log.warn("TODO : ★ 修正 FIXME");
        int yInEMU = yInPixel * Units.EMU_PER_PIXEL;
        int currentYInEMU = 0;

        for (int rowNum = 0; ; rowNum++) {
            Row currentRow = sheet.getRow(rowNum);
            if (currentRow == null) {
                currentRow = sheet.createRow(rowNum);
            }

            if (currentRow.getHeight() == sheet.getDefaultRowHeight()) {
                // Excelのデフォルト値は違うので明示的に設定する
                currentRow.setHeightInPoints(13.5f);
            }

            // 行の高さ（EMU）
            // ※row.getHeight()は1/20ポイント単位
            int rowHeightInEMU = (int) (currentRow.getHeight() / 20.0D * Units.EMU_PER_POINT);
            if (yInEMU < currentYInEMU + rowHeightInEMU) {
                return new int[]{rowNum, (yInEMU - currentYInEMU)};
            }
            currentYInEMU += rowHeightInEMU;
        }
    }


    /**
     * セットしたCellが含まれるRangeを返す。
     *
     * @param cell Cell
     * @return CellRangeAddress
     */
    static CellRangeAddress getExpectCell(Cell cell) {

        Sheet targetSheet = cell.getSheet();
        int startRow = cell.getRowIndex();
        int startCol = cell.getColumnIndex();

        List<CellRangeAddress> collect = targetSheet.getMergedRegions().stream().filter(e ->
            startRow >= e.getFirstRow()
                && startRow <= e.getLastRow()
                && startCol >= e.getFirstColumn()
                && startCol <= e.getLastColumn()
        ).toList();

        if (collect.isEmpty()) {
            return null; //not exist
        }

        CellRangeAddress address = collect.getFirst();

        CellRangeAddress margeAddr = new CellRangeAddress(
            address.getFirstRow(), address.getLastRow(),
            address.getFirstColumn(), address.getLastColumn());

        log.debug("merged cell in sheetName = {}", targetSheet.getSheetName());

        return margeAddr;
    }




    /**
     * Copy sheet.
     * 下記の設定を含む。
     * ・印刷範囲
     * ・ページ改行 Row,Column
     * ・印刷余白
     *
     * @param wb         Workbook object.
     * @param sheetIndex index.
     * @return success clone Sheet object.  error: null
     */
    static Sheet cloneSheet(Workbook wb, int sheetIndex) {
        Sheet orgSheet;
        Sheet cloneSheet = null;
        int clonedSheetIndex;

        try {
            orgSheet = wb.getSheetAt(sheetIndex);
            cloneSheet = wb.cloneSheet(sheetIndex);

            log.trace("元シート名:{} コピーシート名: {} ", orgSheet.getSheetName(), cloneSheet.getSheetName());

            //プロパティーを取得
            PrintSetup printSetup = orgSheet.getPrintSetup();
            PrintSetup newSetup = cloneSheet.getPrintSetup();

            // コピー枚数
            newSetup.setCopies(printSetup.getCopies());
            // 下書きモード
            newSetup.setDraft(printSetup.getDraft());
            // シートに収まる高さのページ数
            newSetup.setFitHeight(printSetup.getFitHeight());
            // シートが収まる幅のページ数
            newSetup.setFitWidth(printSetup.getFitWidth());
            // フッター余白
            newSetup.setFooterMargin(printSetup.getFooterMargin());
            // ヘッダー余白
            newSetup.setHeaderMargin(printSetup.getHeaderMargin());
            // 水平解像度
            newSetup.setHResolution(printSetup.getHResolution());
            // 横向きモード
            newSetup.setLandscape(printSetup.getLandscape());
            // 左から右への印刷順序
            newSetup.setLeftToRight(printSetup.getLeftToRight());
            // 白黒
            newSetup.setNoColor(printSetup.getNoColor());
            // 向き
            newSetup.setNoOrientation(printSetup.getNoOrientation());
            // 印刷メモ
            newSetup.setNotes(printSetup.getNotes());
            // ページの開始
            newSetup.setPageStart(printSetup.getPageStart());
            // 用紙サイズ
            newSetup.setPaperSize(printSetup.getPaperSize());
            // スケール
            newSetup.setScale(printSetup.getScale());
            // 使用ページ番号
            newSetup.setUsePage(printSetup.getUsePage());
            // 有効な設定
            newSetup.setValidSettings(printSetup.getValidSettings());
            // 垂直解像度
            newSetup.setVResolution(printSetup.getVResolution());

            //上マージン
            cloneSheet.setMargin(PageMargin.TOP, orgSheet.getMargin(PageMargin.TOP));
            //下マージン
            cloneSheet.setMargin(PageMargin.BOTTOM, orgSheet.getMargin(PageMargin.BOTTOM));
            //左マージン
            cloneSheet.setMargin(PageMargin.LEFT, orgSheet.getMargin(PageMargin.LEFT));
            //右マージン
            cloneSheet.setMargin(PageMargin.RIGHT, orgSheet.getMargin(PageMargin.RIGHT));
            //ヘッダーマージン
            cloneSheet.setMargin(PageMargin.HEADER, orgSheet.getMargin(PageMargin.HEADER));
            //フッターマージン
            cloneSheet.setMargin(PageMargin.FOOTER, orgSheet.getMargin(PageMargin.FOOTER));

            //印刷プレビュー
            clonedSheetIndex = wb.getSheetIndex(cloneSheet);
            String orgSheetName = orgSheet.getSheetName();
            String clonedSheetName = cloneSheet.getSheetName();
            String printArea = wb.getPrintArea(sheetIndex);
            //シート名+!を除去。setPrintAreaでエラーになる。
            printArea = printArea.replace(orgSheetName + "!", "");
            log.trace("printArea [{}]", printArea);

            wb.setPrintArea(clonedSheetIndex, printArea);
            wb.setSheetName(clonedSheetIndex, clonedSheetName);

            //削除
            //wb.removePrintArea(clonedSheetIndex);

        } catch (Exception e) {
            log.error(e.getMessage());
        }
        return cloneSheet;
    }




    /**
     * 指定したシートの印刷情報を出力
     */
    static void sheetPrintInfo(Sheet sheet) {
        try (Workbook wb = sheet.getWorkbook()) {
            int index = wb.getSheetIndex(sheet.getSheetName());
            log.trace("[WorkbookUtil] 印刷情報 シート名:{} sheetIndex: {} ",
                sheet.getSheetName(), index);

            /*プロパティーを取得*/
            PrintSetup printSetup = sheet.getPrintSetup();
            double inch = 2.54;

            log.trace("[WorkbookUtil] コピー枚数				printSetup.getCopies()         ={}", printSetup.getCopies());
            log.trace("[WorkbookUtil] 下書きモード				printSetup.getDraft()          ={}", printSetup.getDraft());
            log.trace("[WorkbookUtil] シートに収まる高さのページ数	printSetup.getFitHeight()      ={}", printSetup.getFitHeight());
            log.trace("[WorkbookUtil] シートが収まる幅のページ数	printSetup.getFitWidth()       ={}", printSetup.getFitWidth());
            log.trace("[WorkbookUtil] 水平解像度				printSetup.getHResolution()    ={}", printSetup.getHResolution());
            log.trace("[WorkbookUtil] 横向きモード				printSetup.getLandscape()      ={}", printSetup.getLandscape());
            log.trace("[WorkbookUtil] 左から右への印刷順序		printSetup.getLeftToRight()    ={}", printSetup.getLeftToRight());
            log.trace("[WorkbookUtil] 白黒						printSetup.getNoColor()        ={}", printSetup.getNoColor());
            log.trace("[WorkbookUtil] 向き						printSetup.getNoOrientation()  ={}", printSetup.getNoOrientation());
            log.trace("[WorkbookUtil] 印刷メモ					printSetup.getNotes()          ={}", printSetup.getNotes());
            log.trace("[WorkbookUtil] ページの開始				printSetup.getPageStart()      ={}", printSetup.getPageStart());
            log.trace("[WorkbookUtil] 用紙サイズ				printSetup.getPaperSize()      ={}", printSetup.getPaperSize());
            log.trace("[WorkbookUtil] スケール					printSetup.getScale()          ={}", printSetup.getScale());
            log.trace("[WorkbookUtil] 使用ページ番号			printSetup.getUsePage()        ={}", printSetup.getUsePage());
            log.trace("[WorkbookUtil] 有効な設定				printSetup.getValidSettings()  ={}", printSetup.getValidSettings());
            log.trace("[WorkbookUtil] 垂直解像度				printSetup.getVResolution()    ={}", printSetup.getVResolution());

            log.trace("[WorkbookUtil] フッター余白				printSetup.getFooterMargin()   ={}", printSetup.getFooterMargin() * inch);
            log.trace("[WorkbookUtil] ヘッダー余白				printSetup.getHeaderMargin()   ={}", printSetup.getHeaderMargin() * inch);

            log.trace("[WorkbookUtil] 上マージン	   orgSheet.getMargin(Sheet.TopMargin)    ={}", inch * sheet.getMargin(PageMargin.TOP));
            log.trace("[WorkbookUtil] 下マージン	   orgSheet.getMargin(Sheet.BottomMargin) ={}", inch * sheet.getMargin(PageMargin.BOTTOM));
            log.trace("[WorkbookUtil] 左マージン 	   orgSheet.getMargin(Sheet.LeftMargin)   ={}", inch * sheet.getMargin(PageMargin.LEFT));
            log.trace("[WorkbookUtil] 右マージン	   orgSheet.getMargin(Sheet.RightMargin)  ={}", inch * sheet.getMargin(PageMargin.RIGHT));
            log.trace("[WorkbookUtil] ヘッダーマージン orgSheet.getMargin(Sheet.HeaderMargin) ={}", inch * sheet.getMargin(PageMargin.HEADER));
            log.trace("[WorkbookUtil] フッターマージン orgSheet.getMargin(Sheet.FooterMargin) ={}", inch * sheet.getMargin(PageMargin.FOOTER));

        } catch (Exception e) {
            log.error(e.getMessage());
        }
    }



    /**
     * セルの位置情報を返す。
     * <pre>
     * 引数の行番号とカラム番号から、セルの位置情報を特定し返却する。
     * 例えば左上のセルは"A1"となる。
     * </pre>
     *
     * @param aRowNum (０から始まる)行番号
     * @param aColNum (０から始まる)カラム番号
     * @return セルを位置を表す文字列
     */
    static String convertCellPos(int aRowNum, int aColNum) throws Exception {
        // カラムを表すアルファベットの配列を生成
        final char[] charArray = "ABCDEFGHIJKLMNOPQRSTUVWXYZ".toCharArray();
        final int charSize = charArray.length;
        // オフセットを取得
        int offset = aColNum / charSize;

        String cellPos;
        if (offset == 0) {
            cellPos = String.valueOf(charArray[aColNum]);
        } else if (offset < charSize) {
            cellPos = String.valueOf(charArray[offset - 1])
                + charArray[aColNum - charSize * offset];
        } else {
            throw new Exception("範囲外のセルが指定されています。");
        }
        return String.format("%s%d", cellPos, aRowNum + 1);
    }



    /**
     * XSSFSimpleShapeの塗りつぶし色をアルファ値付きで設定する
     *
     * @param shape 対象のシェイプ
     * @param r     赤
     * @param g     緑
     * @param b     青
     * @param alpha アルファ値 (0-255)
     * @since POI 5.4
     */
    static void setFillColorWithAlpha(XSSFSimpleShape shape, int r, int g, int b, int alpha) {
        CTShapeProperties spPr = shape.getCTShape().getSpPr();

        // 塗りつぶしプロパティを作成
        CTSolidColorFillProperties fillProps = spPr.isSetSolidFill() ? spPr.getSolidFill() : spPr.addNewSolidFill();

        // RGB色を作成
        CTSRgbColor srgbColor = fillProps.isSetSrgbClr() ? fillProps.getSrgbClr() : fillProps.addNewSrgbClr();
        srgbColor.setVal(new byte[]{(byte) r, (byte) g, (byte) b});

        // アルファ値（透明度）を設定
        int alphaPct = (int) ((alpha / 255.0) * 100000);
        CTPositiveFixedPercentage alphaElem = srgbColor.addNewAlpha();
        alphaElem.setVal(alphaPct);
    }


    /**
     * cell の値の形を判別しStringにして返す
     *
     * @param cell Cell
     * @return String 文字列に変換した値
     */
    static String getStringValue(Cell cell) {
        CellType cellType = cell.getCellType();
        return switch (cellType) {
            case STRING -> cell.getStringCellValue();
            case NUMERIC -> String.valueOf(cell.getNumericCellValue());
            case BOOLEAN -> String.valueOf(cell.getBooleanCellValue());
            case FORMULA -> String.valueOf(cell.getCellFormula());
            default -> "";
        };
    }
}
