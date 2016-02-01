package lequ.util.xls;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.jsoup.nodes.Element;

import java.util.ArrayList;
import java.util.List;

public class HtmlRow {
    private List<HtmlCell> cells = new ArrayList<HtmlCell>();
    private HtmlTable table;
    private Integer index = 0; //SUPPRESS
    private int height = 400; //SUPPRESS
    private int colIndex = 0; //SUPPRESS

    public void setHeight(Integer height) {
        this.height = height;
    }

    public HtmlTable getTable() {
        return table;
    }

    public void setTable(HtmlTable table) {
        this.table = table;
    }

    public HtmlRow(HtmlTable table, Element element, Integer index) {
        this.table = table;
        this.index = index;
        colIndex = 0;
        for (Element childElement : element.select("td,th")) {
            createNullTd();
            int rowSpan = 1;
            String strRowSpan = childElement.attr("rowspan");
            if (StringUtils.isNotBlank(strRowSpan) &&
                    StringUtils.isNumeric(strRowSpan)) {
                rowSpan = Integer.parseInt(strRowSpan);
            }
            int colSpan = 1;
            String strColSpan = childElement.attr("colspan");
            if (StringUtils.isNotBlank(strColSpan) &&
                    StringUtils.isNumeric(strColSpan)) {
                colSpan = Integer.parseInt(strColSpan);
            }
            if (rowSpan > 1 || colSpan > 1) {
                this.table.regSpan(colIndex, rowSpan - 1, colSpan);
                table.regMergedRegion(index, colIndex, rowSpan, colSpan);
            }

            for (int i = 0; i < colSpan; i++) {
                if (i == 0) {
                    cells.add(new HtmlCell(this, childElement, colIndex));
                } else {
                    cells.add(new HtmlCell(this, colIndex));
                }
                colIndex++;
            }
        }
    }

    private void createNullTd() {
        int leng = this.table.popColSpan(colIndex);
        if (leng > 0) {
            for (int i = 0; i < leng; i++) {
                cells.add(new HtmlCell(this, colIndex));
                colIndex++;
            }
            createNullTd();
        }
    }

    public void exportXlsRow(HSSFSheet sheet) {
        int top = table.getDiv().getTop();
        HSSFRow row = sheet.createRow(index + top);
        for (HtmlCell cell : cells) {
            cell.exportXlsCell(row);
        }
        if (height > 0) {
            row.setHeight((short) height);
        }
    }
}
