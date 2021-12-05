import java.io.FileInputStream;
import java.io.IOException;
import java.util.List;

/**
 * @author Leonid Cheremshantsev
 */
public class Main {

    public static final String PATH = System.getProperty("user.home") + "\\Desktop\\test.xlsx";

    public static void main(String[] args) throws IOException {
        XlsxWorker xlsxWorker = new XlsxWorker();

        xlsxWorker.write(PATH);

        List<String> strings = xlsxWorker.readFirstColumn(new FileInputStream(PATH));
        strings.forEach(System.out::println);
    }

}
