package jp.classicorange;

import jp.classicorange.types.SearchMode;
import jp.classicorange.utils.CheckParameter;
import jp.classicorange.utils.ExcelUtils;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.*;
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

    /**
     * 本処理。
     *
     */
    public void search(String[] args) throws Exception {

        //検索条件 をセット
        cond = CheckParameter.checkParameter(args);

        System.out.println("File Index / Sheet Index / Cell Index / File Path / Sheet Name / Cell Name / value");

        // 検索の実行
        searchDir(cond.searchDirPath());

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
     * @param aFile ファイルオブジェクト
     */
    private void searchWord(int fileIndex,File aFile) throws Exception {

        log.info("個別検索開始 : {}",aFile.getAbsolutePath());

        SearchMode searchMode = cond.searchMode();
        String searchWord = cond.searchWord() ;
        String replaceWord = cond.replaceWord() ;

        // Excelファイルの読込み
        InputStream inputStream = new FileInputStream(aFile);
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
            StringBuilder sb = new StringBuilder();
            // シート
            Sheet sheet = workbook.getSheetAt(sheetIndex);
            // シート名
            String sheetName = sheet.getSheetName();
            // シートの最終行
            int lastRowNum = sheet.getLastRowNum();
            // 対象ファイルの絶対パス
            String absPath = aFile.getAbsolutePath();

            //シェープ
            ArrayList<HitRecord> hits = searchShape(sheet, searchWord, replaceWord, searchMode);
            if(hits!=null && !hits.isEmpty()){
                for (HitRecord hit: hits) {
                    appendRecord(
                        sb, fileIndex,sheetIndex,-1,
                        absPath,sheetName,hit.cellPos(),
                        hit.text(),
                        hit.replaced()
                    );
                }
            }

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
                        appendRecord( sb, fileIndex, sheetIndex, cellIndex++,
                            absPath, sheetName, ExcelUtils.convertCellPos(j, k),
                            original, result);
                    }

                    if (!sb.isEmpty() && !sb.toString().endsWith("\n")) {
                        sb.append("\n");
                    }
                }
            }
            System.out.print(sb);

        }



        // 上書き保存
        if(replaceWord != null){
            FileOutputStream outputStream = new FileOutputStream(aFile);
            workbook.write(outputStream);
            outputStream.close();
        }
    }


    /**
     * 文字列バッファにファイルパス、シート名、セルの位置情報、値を設定して返却する。
     *
     * @param sb 文字列バッファ
     * @param aFilePath ファイルパス
     * @param aSheetName シート名
     * @param aPosition セルの位置情報
     * @param aValue セルの値
     */
    public void appendRecord(
        StringBuilder sb,
        String fileIndex,
        String sheetIndex,
        String cellIndex,
        String aFilePath,
        String aSheetName,
        String aPosition,
        String aValue,
        String aReplaced
    ) {
        sb.append(fileIndex);
        sb.append("\t");
        sb.append(sheetIndex);
        sb.append("\t");
        sb.append(cellIndex);
        sb.append("\t");
        sb.append(aFilePath);
        sb.append("\t");
        sb.append(aSheetName);
        sb.append("\t");
        sb.append(aPosition);
        sb.append("\t");
        sb.append(aValue);
        sb.append("\t");
        sb.append(aReplaced);
        sb.append("\n");
    }

    public void appendRecord(
        StringBuilder sb,
        int fileIndex,
        int sheetIndex,
        int cellIndex,
        String aFilePath,
        String aSheetName,
        String aPosition,
        String aValue,
        String aReplaced
    ) {
        appendRecord(sb,fileIndex+"",sheetIndex+"",cellIndex+"",aFilePath,aSheetName,aPosition,aValue,aReplaced);
    }


    /**
     *
     * 指定したシートのシェイプの文字列を検索して返す
     *
     * @param sheet Sheet
     * @param searchWord 検索ワード
     * @return 検索結果のリスト
     */
    public ArrayList<HitRecord> searchShape(Sheet sheet, String searchWord, String replaceWord, SearchMode searchMode) throws Exception {

        ArrayList<HitRecord> shapeList = new ArrayList<>();

        if (sheet instanceof XSSFSheet xssfSheet) {
            XSSFDrawing drawing = xssfSheet.getDrawingPatriarch();
            if (drawing == null) return shapeList;

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
                shapeList.add(new HitRecord(
                    "",
                    sheet.getSheetName(),
                    ExcelUtils.convertCellPos(anchor.getRow1(), anchor.getCol1()),
                    text,
                    result
                ));
            }

        } else if (sheet instanceof HSSFSheet hssfSheet) {
            HSSFPatriarch patriarch = hssfSheet.getDrawingPatriarch();
            if (patriarch == null) return shapeList;

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
                shapeList.add(new HitRecord(
                    "",
                    sheet.getSheetName(),
                    ExcelUtils.convertCellPos(row, col),
                    text,
                    result
                ));
            }
        }

        return shapeList;

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



    /** 検索した結果を格納するエンティティ */
    public record HitRecord(String path, String sheetName, String cellPos, String text, String replaced) {
    }

}
