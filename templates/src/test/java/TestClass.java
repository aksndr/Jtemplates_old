import java.util.HashMap;

/**
 * User: a.arzamastsev Date: 31.07.13 Time: 14:20
 */
public class TestClass {
    public static void main(String[] args) throws Exception {
        String templateFilepath = "Регистрационная карточка.docx";
        HashMap<String, String> params = new HashMap<String, String>();
        params.put("type", "тест 1");
        params.put("name", "тест 2");
        params.put("k2864452_2", "Привет");
        params.put("k2864452_3", "test4");
        params.put("k2864452_5", "test5");
        params.put("k2864452_4", "дадада");
        params.put("k2864452_17", "test7");
        params.put("k2864452_11", "test8");
        params.put("k2864452_12", "test9");
        params.put("k2864452_13", "test10");
        params.put("k2864452_14", "test11");
        params.put("k2864452_15", "test12");
        params.put("barcode", "16");

        if ((CompleteTemplate.instance().completeTemplate(templateFilepath, params)).isEmpty()) {
            System.out.println("Fail");
        } else {
            System.out.println("Done");
        }
    }
}
