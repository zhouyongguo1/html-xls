package lequ.util.xls;

import org.apache.poi.ss.util.CellRangeAddress;

public class RangeAddress {
    private Integer rowIndex;
    private Integer colIndex;
    private Integer rowSpan;
    private Integer colSpan;

    public RangeAddress(Integer rowIndex, Integer colIndex, Integer rowSpan, Integer colSpan) {
        this.colIndex = colIndex;
        this.colSpan = colSpan;
        this.rowIndex = rowIndex;
        this.rowSpan = rowSpan;
    }

    public CellRangeAddress mergedRegion(int left, int top) {
        return new CellRangeAddress(rowIndex + top, rowIndex + top + rowSpan - 1,
                colIndex + left, colIndex + left + colSpan - 1);
    }

}
