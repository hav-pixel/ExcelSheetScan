import jp.classicorange.SearchExcel;

/**
 * ・第一引数：検索対象のディレクトリパス
 * ・第二引数：検索対象文字列
 * ・第三引数：検索モード(あいまい検索/完全一致)  FUZZY
 * ・第四引数：置換文字列
 */
public class Main {

    public static void main(String[] args) throws Exception {
        new SearchExcel().search(args);
    }

}