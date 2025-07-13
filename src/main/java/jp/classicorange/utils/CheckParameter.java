package jp.classicorange.utils;

import jp.classicorange.types.SearchMode;

public class CheckParameter {

    public record SearchCond(String searchDirPath, String searchWord, SearchMode searchMode, String replaceWord) {}

    /**
     * 渡された検索条件の配列をフィールドにセット
     *
     * @param args 検索条件
     * @return 検索パス
     */
    public static SearchCond checkParameter(String[] args){

        int argLen = args.length;
        if(argLen < 3){
            System.out.println("""
                 設定する引数は以下の通り。
                 ・第一引数：検索対象のディレクトリパス  e.g.  .\\testData\\
                 ・第二引数：検索対象文字列  e.g. Apple
                 ・第三引数：検索モード  FUZZY あいまい検索, STRICTLY 完全一致
                 ・第四引数：置換文字列  e.g. りんご
                
                例：
                .\\build\\install\\SearchDocs\\bin\\SearchDocs.bat .\\testData\\ "Apple" FUZZY
                """);

            System.exit(1);
        }


        // 引数から情報を取得
        String searchDirPath = args[0];
        String searchWord = args[1];
        String replaceWord = null;
        SearchMode searchMode = SearchMode.valueOf(args[2]);
        // 置換処理は任意
        if(args.length > 3){
            replaceWord = args[3];
        }

        // メッセージを表示
        System.out.println("以下の条件でgrep検索を実行します。");
        System.out.println("検索対象フォルダ：" + searchDirPath);
        System.out.println("検索文字列：" + searchWord);
        System.out.println("検索方法：" + searchMode);
        if(args.length > 3){
            System.out.println("置換文字列：" + replaceWord);
        }

        return new SearchCond(searchDirPath,searchWord, searchMode, replaceWord);
    }

}
