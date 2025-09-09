import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;

import java.util.Comparator;

class RowsBankNameComparator implements Comparator<Row> {

    private final int bankNameColIndex;

    public RowsBankNameComparator(int bankNameColIndex) {
        this.bankNameColIndex = bankNameColIndex;
        if (bankNameColIndex == -1) throw new RuntimeException("Bank Name column not found");
    }

    @Override
    public int compare(Row r1, Row r2) {

        Cell c1 = r1.getCell(bankNameColIndex);
        Cell c2 = r2.getCell(bankNameColIndex);

        if ((c1 != null && c1.getCellType() != CellType.BLANK) && (c2 != null && c2.getCellType() == CellType.BLANK)) {
            return c1.getStringCellValue().compareTo(c2.getStringCellValue());
        } else if (c1 != null && c1.getCellType() != CellType.BLANK) {
            return 1;
        } else if ((c2 != null && c2.getCellType() != CellType.BLANK)) {
            return -1;
        }
        return 0;
    }
}
