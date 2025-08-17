package jp.classicorange;

import jp.classicorange.types.SearchMode;
import jp.classicorange.utils.CheckParameter;
import jp.classicorange.utils.ExcelUtils;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.*;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * エクセルファイルのgrepツール。
 * <pre>
 * 指定したフォルダからエクセルファイルを再帰的に検索し、
 * 指定した文字列をgrepします。
 *
 * 設定する引数は以下の通り。
 * ・第一引数：検索対象のディレクトリパス  e.g.  .\testData\
 * ・第二引数：検索対象文字列  e.g. シュミレーション
 * ・第三引数：検索モード  FUZZY あいまい検索, STRICTLY 完全一致
 * ・第四引数：置換文字列  e.g. シミュレーション
 * </pre>
 *
 */
public class SearchExcel {

    private static final Logger log = LoggerFactory.getLogger(SearchExcel.class);
    private CheckParameter.SearchCond cond;

    private SXSSFWorkbook sxssfWorkbook ;
    private SXSSFSheet sxssfSheet ;
    private CreationHelper createHelper;


    /**
     * 本処理。
     *
     */
    public void search(String[] args) throws Exception {

        //検索条件 をセット
        cond = CheckParameter.checkParameter(args);

        //結果保存ファイル名
        String resultFileName = "result.xlsx"; //TODO ファイル存在時ナンバリング

        //Workbookファイル準備
        File f = new File(resultFileName);
        if(f.exists()) {
            FileInputStream fis = new FileInputStream(f);
            Workbook wb = WorkbookFactory.create(fis);
            if (wb instanceof XSSFWorkbook) {
                // xlsx → SXSSFWorkbookにラップして追記可能
                sxssfWorkbook = new SXSSFWorkbook((XSSFWorkbook) wb, 100);
            } else if (wb instanceof HSSFWorkbook) {
                // xls → SXSSFは不可、HSSFWorkbookのまま扱う
                System.out.println("警告: .xls ファイルは SXSSFWorkbook 非対応です。HSSFWorkbookを使用します。");
                // 必要なら変換処理を入れる
                // sxssfWorkbook = convertHssfToSxssf((HSSFWorkbook) wb);
            } else {
                throw new IllegalStateException("未知のWorkbook種類");
            }
        }else{
            sxssfWorkbook = new SXSSFWorkbook(100);
        }
        String formatted = LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyyMMdd_HHmmss"));
        // Excelのシート名に使える文字をセット。禁止文字は_に置換
        sxssfSheet = sxssfWorkbook.createSheet(WorkbookUtil.createSafeSheetName(args[1]+" "+formatted, '_'));
        // 一番左に移動
        sxssfWorkbook.setSheetOrder(sxssfSheet.getSheetName(), 0);
        // アクティブシートを新シートに
        int idx = sxssfWorkbook.getSheetIndex(sxssfSheet);
        sxssfWorkbook.setActiveSheet(idx);
        sxssfWorkbook.setSelectedTab(idx);
        createHelper = sxssfWorkbook.getCreationHelper();

        // シートにヘッダー生成
        setHeader();

        // 検索の実行
        searchDir(cond.searchDirPath());

        // 結果 保存
        try (FileOutputStream fos = new FileOutputStream(resultFileName)) {
            sxssfWorkbook.write(fos);
        }finally {
            // 結果ファイル閉じる
            sxssfWorkbook.close();
        }


    }


    /**
     * 指定したパスを検索する。
     * 注意：再帰的に使用される
     *
     * @param targetPath ディレクトリパス
     */
    private boolean searchDir(final String targetPath) throws Exception {

        File file = new File(targetPath);
        File[] listFiles = file.listFiles((aDir, aName) -> {
            // ドットで始まるファイルは対象外
            if (aName.startsWith(".")) {
                return false;
            }

            // 対象要素の絶対パスを取得
            String absolutePath = aDir.getAbsolutePath() + File.separator + aName;

            // エクセルファイルのみ対象とする。
            if (new File(absolutePath).isFile()
                && (absolutePath.endsWith(".xls") || absolutePath
                .endsWith(".xlsx"))) {
                return true;
            } else {
                // ディレクトリの場合、再び同一メソッドを呼出す。
                try {
                    return searchDir(absolutePath);
                } catch (Exception e) {
                    log.error("検索エラー : {}",e.getMessage());
                    return false;
                }
            }
        });

        if (listFiles == null) {
            log.info("対象ファイル収集 : {}", 0);
            return false;
        }
        log.info("対象ファイル収集 : {}", listFiles.length);

        // 検索を実行
        for (int i = 0; i < listFiles.length; i++) {
            File f = listFiles[i];
            if (f.isFile()) {
                searchWord(i,f);
            }
        }
        //System.out.println("検索が完了しました。");
        return true;
    }

    /**
     * 対象のエクセルシートから文字列を検索し、リストに格納します。
     *
     * @param file ファイルオブジェクト
     */
    private void searchWord(int fileIndex, File file) throws Exception {

        log.info("個別検索開始 : {}",file.getAbsolutePath());

        SearchMode searchMode = cond.searchMode();
        String searchWord = cond.searchWord() ;
        String replaceWord = cond.replaceWord() ;

        // Excelファイルの読込み
        InputStream inputStream = new FileInputStream(file);
        Workbook workbook;
        try {
            workbook = WorkbookFactory.create(inputStream);
        } catch (Exception ex) {
            return;
        }
        inputStream.close();

        // シート枚数を読込み
        int numberOfSheets = workbook.getNumberOfSheets();
        // シート枚数分ループ処理
        for (int sheetIndex = 0; sheetIndex < numberOfSheets; sheetIndex++) {

            String path = file.getAbsolutePath();
            // シート
            Sheet sheet = workbook.getSheetAt(sheetIndex);
            //シェープ
            searchShape(
                fileIndex, path,
                sheet, searchWord, replaceWord, searchMode);
            //セル
            searchCell(fileIndex, sheetIndex,
                sheet, searchWord, replaceWord, searchMode,
                path);


        }

        // 置換、上書き保存
        if(replaceWord != null){
            FileOutputStream outputStream = new FileOutputStream(file);
            workbook.write(outputStream);
            outputStream.close();
        }
    }



    public void searchCell(
        int fileIndex,
        int sheetIndex, Sheet sheet,
        String searchWord, String replaceWord, SearchMode searchMode,
        String absPath
    ) throws Exception {
        // シート名
        String sheetName = sheet.getSheetName();
        // シートの最終行
        int lastRowNum = sheet.getLastRowNum();
        // 対象ファイルの絶対パス

        int cellIndex=0;
        // 最終行までループ処理
        for (int j = 0; j <= lastRowNum; j++) {
            // 行の取得
            Row row = sheet.getRow(j);
            if (row == null) continue;

            // 行内の最後のセルの位置
            short lastCellNum = row.getLastCellNum();
            // 行内の最後のセルまでループ処理
            for (int k = 0; k < lastCellNum; k++) {
                // セルを取得
                Cell cell = row.getCell(k);
                if (cell == null) continue;
                // セルの値を文字列へ変換
                String original = ExcelUtils.getStringValue(cell);

                if (original.contains(searchWord)) {
                    // 置換処理を実行
                    String result=replaceWord(cell,searchWord,replaceWord, searchMode);
                    // 結果出力
                    appendRecord( fileIndex, sheetIndex, cellIndex++,
                        absPath, sheetName, ExcelUtils.convertCellPos(j, k),
                        original, result);
                }

//                if (!sb.isEmpty() && !sb.toString().endsWith("\n")) {
//                    sb.append("\n");
//                }
            }
        }
    }




    private void setHeader() {
        SXSSFRow row = sxssfSheet.createRow(sxssfSheet.getLastRowNum() + 1);
        int c=0;
        row.createCell(c++).setCellValue("fileIndex");
        row.createCell(c++).setCellValue("sheetIndex");
        row.createCell(c++).setCellValue("cellIndex");
        row.createCell(c++).setCellValue("filePath");
        row.createCell(c++).setCellValue("sheetName");
        row.createCell(c++).setCellValue("position");
        row.createCell(c++).setCellValue("value");
        row.createCell(c++).setCellValue("replaced");
        row.createCell(c++).setCellValue("link");
    }


    /**
     * 文字列バッファにファイルパス、シート名、セルの位置情報、値を設定して返却する。
     *
     * @param fileIndex fileIndex
     * @param sheetIndex sheetIndex
     * @param cellIndex cellIndex
     * @param filePath ファイルパス
     * @param sheetName シート名
     * @param position セルの位置情報
     * @param value セルの値
     */
    public void appendRecord(
        int fileIndex,
        int sheetIndex,
        int cellIndex,
        String filePath,
        String sheetName,
        String position,
        String value,
        String replaced
    ) {
        int c=0;
        SXSSFRow row = sxssfSheet.createRow(sxssfSheet.getLastRowNum() + 1);
        int r = row.getRowNum() + 1;
        row.createCell(c++).setCellValue(fileIndex);
        row.createCell(c++).setCellValue(sheetIndex);
        row.createCell(c++).setCellValue(cellIndex);
        row.createCell(c++).setCellValue(filePath);
        row.createCell(c++).setCellValue(sheetName);
        row.createCell(c++).setCellValue(position);
        row.createCell(c++).setCellValue(value);
        row.createCell(c++).setCellValue(replaced);
        Cell cell = row.createCell(c++);
        cell.setCellFormula(String.format("HYPERLINK(D%s & \"#'\" & E%s & \"'!\" & F%s, \"LINK\")",r,r,r));
    }



    /**
     *
     * 指定したシートのシェイプの文字列を検索して返す
     *
     * @param sheet Sheet
     * @param searchWord 検索ワード
     */
    public void searchShape(int fileIndex, String filePath, Sheet sheet, String searchWord, String replaceWord, SearchMode searchMode) throws Exception {

        if (sheet instanceof XSSFSheet xssfSheet) {
            XSSFDrawing drawing = xssfSheet.getDrawingPatriarch();
            if (drawing == null) return;

            for (Shape shape : drawing.getShapes()) {
                if (!(shape instanceof XSSFSimpleShape xshape)) continue;
                String text = xshape.getText();
                if (text == null || !text.contains(searchWord)) continue;

                String result = text;
                if( replaceWord !=null ){
                    if(searchMode == SearchMode.FUZZY){
                        //FUZZY
                        Pattern pattern = Pattern.compile(searchWord);
                        Matcher matcher = pattern.matcher(text);
                        result = matcher.replaceAll(replaceWord);
                    }else{
                        //STRICT
                        result = replaceWord;
                    }
                    xshape.setText(result);
                }

                ClientAnchor anchor = (XSSFClientAnchor) xshape.getAnchor();
                String sheetName = sheet.getSheetName();

                appendRecord(
                    fileIndex,
                    sheet.getWorkbook().getSheetIndex(sheetName),
                    -1,
                    filePath,
                    sheet.getSheetName(),
                    ExcelUtils.convertCellPos(anchor.getRow1(), anchor.getCol1()),
                    result,
                    replaceWord
                );
            }

        } else if (sheet instanceof HSSFSheet hssfSheet) {
            HSSFPatriarch patriarch = hssfSheet.getDrawingPatriarch();
            if (patriarch == null) return ;

            for (HSSFShape shape : patriarch.getChildren()) {

                if (!(shape instanceof HSSFSimpleShape simpleShape)) continue;

                HSSFRichTextString rText = simpleShape.getString();
                String text = (rText != null) ? rText.getString() : "";
                if (!text.contains(searchWord)) continue;

                String result = null;
                if( replaceWord!=null && !replaceWord.isEmpty()){
                    if(searchMode == SearchMode.FUZZY){
                        //FUZZY
                        Pattern pattern = Pattern.compile(searchWord);
                        Matcher matcher = pattern.matcher(text);
                        result = matcher.replaceAll(replaceWord);
                    }else{
                        //STRICT
                        result = replaceWord;
                    }
                    // テキストボックスに文字列をセット
                    simpleShape.setString(new HSSFRichTextString(result));
                }

                HSSFAnchor anchor = shape.getAnchor();
                if (!(anchor instanceof HSSFClientAnchor clientAnchor)) continue;

                int row = clientAnchor.getRow1();
                int col = clientAnchor.getCol1();

                String sheetName = sheet.getSheetName();
                appendRecord(
                    fileIndex,
                    sheet.getWorkbook().getSheetIndex(sheetName),
                    -1,
                    filePath,
                    sheet.getSheetName(),
                    ExcelUtils.convertCellPos(row, col),
                    result,
                    replaceWord
                );
            }
        }


    }

    /**
     *
     * @param cell Cell
     * @param searchWord String search word
     * @param replaceWord String replace word
     * @param searchMode SearchMode mode
     * @return String 置換した/しないときの実行した結果を返す
     */
    private String replaceWord(Cell cell, String searchWord, String replaceWord, SearchMode searchMode){

        //置換ワードがない場合、何もしない
        if(replaceWord == null)return "";

        String result = replaceWord;

        //モードで置換振り分け
        if(searchMode == SearchMode.FUZZY){
            //FUZZY
            Pattern pattern = Pattern.compile(searchWord);
            Matcher matcher = pattern.matcher(ExcelUtils.getStringValue(cell));
            result = matcher.replaceAll(replaceWord);
            cell.setCellValue(result);
        }else{
            //STRICT
            cell.setCellValue(result);
        }
        return result;
    }

}
