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


}