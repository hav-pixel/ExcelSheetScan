import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.TestFactory;

import java.io.ByteArrayOutputStream;
import java.io.PrintStream;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;


public class MainTest {

    @BeforeAll
    public static void before(){
        MakeTestExcel.main(null);
    }


    @Test
    public void testMainWithThreeArgs() throws Exception {
        String[] args = {".\\testData", "Jakarta", "FUZZY"};
        Main.main(args);
        // ※ System.out 出力を検証したいなら PrintStream をモック化する
    }


    /**
     * System.out で出されるログの検証
     * 置換するテスト
     */
    @Test
    public void testMainOutput() {
        // System.outを一時的に差し替え
        ByteArrayOutputStream outContent = new ByteArrayOutputStream();
        PrintStream originalOut = System.out;
        System.setOut(new PrintStream(outContent));

        try {
            //置換
            String[] args = {".\\testData", "Jakarta", "FUZZY", "●●●●●●"};
            Main.main(args);

            // 出力結果を検証
            String output = outContent.toString().trim();

            assertTrue(
                output.contains("以下の条件でgrep検索を実行します。") &&
                output.contains("検索対象フォルダ：.\\testData") &&
                output.contains("検索文字列：Jakarta") &&
                output.contains("検索方法：FUZZY"),
                "期待した文字列を含んだ、条件表示がありません。");

        } catch (Exception e) {
            throw new RuntimeException(e);
        } finally {
            // 元に戻す
            System.setOut(originalOut);
        }
    }

}