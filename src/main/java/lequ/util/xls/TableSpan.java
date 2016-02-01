package lequ.util.xls;

public class TableSpan {
    private int rowSpan;
    private int colSpanl;

    public TableSpan(int rowSpan, int colSpanl) {
        this.rowSpan = rowSpan;
        this.colSpanl = colSpanl;
    }

    public int pop() {
        if (rowSpan == 0) {
            return 0;
        } else {
            rowSpan--;
            return colSpanl;
        }
    }

}
