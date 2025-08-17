import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Test;


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


}