import checkers.nullness.quals.NonNull;
import lombok.SneakyThrows;
import lombok.val;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.PrintStream;

import static java.lang.Math.abs;

public interface ExcelEqualiser {

    void wrongType(int rowIndex, int cellIndex, CellType cell1TypeEnum, CellType cell2TypeEnum);

    void wrongValue(int rowIndex, int cellIndex, String cell1StringValue, String cell2StringValue);

    void wrongValue(int rowIndex, int cellIndex, boolean cell1BooleanValue, boolean cell2BooleanValue);

    void wrongValue(int rowIndex, int cellIndex, double cell1NumericValue, double cell2NumericValue);

    void wrongFormulaValue(int rowIndex, int cellIndex, String cell1FormulaValue, String cell2FormulaValue);

    static ExcelEqualiser getSystemOutInstance() {
        return getInstance(System.out);
    }

    static ExcelEqualiser getInstance(PrintStream out) {
        return new ExcelEqualiser() {
            @Override
            public void wrongValue(int rowIndex, int cellIndex, String cell1StringValue, String cell2StringValue) {
                out.println("В строке №" + rowIndex
                        + " ячейка с логическим значением №" + cellIndex
                        + " не совпадает по значению: в первой вкладке значение ячейки " + cell1StringValue
                        + ", а во второй - " + cell2StringValue);
            }

            @Override
            public void wrongValue(int rowIndex, int cellIndex, boolean cell1BooleanValue, boolean cell2BooleanValue) {
                out.println("В строке №" + rowIndex
                        + " ячейка с логическим значением №" + cellIndex
                        + " не совпадает по значению: в первой вкладке значение ячейки " + cell1BooleanValue
                        + ", а во второй - " + cell2BooleanValue);
            }

            @Override
            public void wrongValue(int rowIndex, int cellIndex, double cell1NumericValue, double cell2NumericValue) {
                out.println("В строке №" + rowIndex
                        + " ячейка с числом №" + cellIndex
                        + " не совпадает по значению: в первой вкладке значение ячейки " + cell1NumericValue
                        + ", а во второй - " + cell2NumericValue);
            }

            @Override
            public void wrongFormulaValue(int rowIndex, int cellIndex, String cell1FormulaValue, String cell2FormulaValue) {
                out.println("В строке №" + rowIndex
                        + " ячейка с формулой №" + cellIndex
                        + " не совпадает по значению: в первой вкладке значение ячейки " + cell1FormulaValue
                        + ", а во второй - " + cell2FormulaValue);
            }

            @Override
            public void wrongType(int rowIndex, int cellIndex, CellType cell1TypeEnum, CellType cell2TypeEnum) {
                out.println("В строке №" + rowIndex
                        + " ячейка №" + cellIndex
                        + " не совпадает по типу: в первой вкладке тип ячейки " + cell1TypeEnum
                        + ", а во второй - " + cell2TypeEnum);
            }
        };
    }

    default void isEquals(@NonNull String fileName, int maxColumn, int maxRow) {
        isEquals(new File(fileName), maxColumn, maxRow);
    }

    @SneakyThrows
    default void isEquals(@NonNull File file, int maxColumn, int maxRow) {

        assert file.exists();
        assert maxRow > 0;
        assert maxColumn > 0;

        try (val fileInputStream = new FileInputStream(file);
             val sheets = new XSSFWorkbook(fileInputStream)) {

            val sheet1 = sheets.getSheetAt(0);
            val sheet2 = sheets.getSheetAt(1);

            for (int rowIndex = 0; rowIndex <= maxRow; rowIndex++) {

                val row1 = sheet1.getRow(rowIndex);
                val row2 = sheet2.getRow(rowIndex);

                if (row1 != null && row2 != null)
                    for (int cellIndex = 0; cellIndex <= maxColumn; cellIndex++) {
                        val cell1 = row1.getCell(cellIndex);
                        val cell2 = row2.getCell(cellIndex);

                        if (cell1 != null && cell2 != null) {
                            val cellTypeEnum = cell1.getCellTypeEnum();

                            if (!cellTypeEnum.equals(cell2.getCellTypeEnum()))
                                wrongType(rowIndex, cellIndex, cell2.getCellTypeEnum(), cellTypeEnum);
                            else
                                switch (cell1.getCellTypeEnum()) {
                                    case STRING:
                                        checkStringCells(rowIndex, cellIndex, cell1, cell2);
                                        break;

                                    case NUMERIC:
                                        checkNumericCells(rowIndex, cellIndex, cell1, cell2);
                                        break;

                                    case BOOLEAN:
                                        checkBooleanCells(rowIndex, cellIndex, cell1, cell2);
                                        break;

                                    case FORMULA:
                                        checkFormulaCells(rowIndex, cellIndex, cell1, cell2);
                                        break;

                                    case ERROR:
                                        break;
                                }
                        }
                    }
            }
        }
    }

    default void checkFormulaCells(int rowIndex, int cellIndex, XSSFCell cell1, XSSFCell cell2) {
        val cell1FormulaValue = cell1.getCellFormula();
        val cell2FormulaValue = cell2.getStringCellValue();
        if (!cell1FormulaValue.equals(cell2FormulaValue))
            wrongFormulaValue(rowIndex, cellIndex, cell1FormulaValue, cell2FormulaValue);
    }

    default void checkBooleanCells(int rowIndex, int cellIndex, XSSFCell cell1, XSSFCell cell2) {
        val cell1BooleanValue = cell1.getBooleanCellValue();
        val cell2BooleanValue = cell2.getBooleanCellValue();
        if (cell1BooleanValue ^ cell2BooleanValue)
            wrongValue(rowIndex, cellIndex, cell1BooleanValue, cell2BooleanValue);
    }

    default void checkNumericCells(int rowIndex, int cellIndex, XSSFCell cell1, XSSFCell cell2) {
        val cell1NumericValue = cell1.getNumericCellValue();
        val cell2NumericValue = cell2.getNumericCellValue();
        if (abs(cell1NumericValue - cell2NumericValue) > 0.01)
            wrongValue(rowIndex, cellIndex, cell1NumericValue, cell2NumericValue);
    }

    default void checkStringCells(int rowIndex, int cellIndex, XSSFCell cell1, XSSFCell cell2) {
        val cell1StringValue = cell1.getStringCellValue();
        val cell2StringValue = cell2.getStringCellValue();
        if (!cell1StringValue.equals(cell2StringValue))
            wrongValue(rowIndex, cellIndex, cell1StringValue, cell2StringValue);
    }
}
