import lombok.val;
import org.junit.Test;

public class ExcelEqualiserTest {

    private String file = "/Users/admin/IdeaProjects/Cognos-test/excel-equalaser/src/test/resources/Workbook1.xlsx";

    @Test
    public void isEquals() {
        val systemOutInstance = ExcelEqualiser.getSystemOutInstance();

        systemOutInstance.isEquals(file, 10, 10);

    }
}