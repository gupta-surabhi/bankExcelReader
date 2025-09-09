import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;

import java.util.Comparator;

class RowsAgeComparator implements Comparator<Row> {

    private final int ageColIndex;

    public RowsAgeComparator(int ageColIndex) {
        this.ageColIndex = ageColIndex;
        if (ageColIndex == -1) throw new RuntimeException("Age column not found");
    }

    @Override
    public int compare(Row r1, Row r2) {
        Cell c1 = r1.getCell(ageColIndex);
        Cell c2 = r2.getCell(ageColIndex);
        Integer age1 = (c1 != null && c1.getCellType() != CellType.BLANK) ? Integer.valueOf(c1.getStringCellValue().replace("Day(s)","").trim()) : Integer.MIN_VALUE;
        Integer age2 = (c2 != null && c2.getCellType() != CellType.BLANK) ? Integer.valueOf(c2.getStringCellValue().replace("Day(s)","").trim()) : Integer.MIN_VALUE;
        return Integer.compare(age1, age2);
    }
}
