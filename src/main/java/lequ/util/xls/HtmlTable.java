package lequ.util.xls;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.jsoup.nodes.Element;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class HtmlTable {

    private List<HtmlRow> rows = new ArrayList<HtmlRow>();
    private Map<Integer, TableSpan> spanMap = new HashMap<Integer, TableSpan>();
    private List<RangeAddress> mergedRegions = new ArrayList<RangeAddress>();
    private HtmlDiv div;

    public HtmlTable(HtmlDiv div, Element element, Integer index) {
        this.div = div;
        Integer rowIndex = index;
        for (Element childElement : element.select("tr")) {
            HtmlRow row = new HtmlRow(this, childElement, rowIndex);
            rows.add(row);
            rowIndex++;
        }
    }

    public void regSpan(int colIndex, int rowSpan, int colSpan) {
        spanMap.put(colIndex, new TableSpan(rowSpan, colSpan));
    }

    public void regMergedRegion(Integer rowIndex, Integer colIndex, Integer rowSpan, Integer colSpan) {
        RangeAddress rangeAddress = new RangeAddress(rowIndex, colIndex, rowSpan, colSpan);
        mergedRegions.add(rangeAddress);
    }

    public int popColSpan(int colIndex) {
        TableSpan tableSpan = spanMap.get(colIndex);
        if (tableSpan != null) {
            return tableSpan.pop();
        } else {
            return 0; //SUPPRESS
        }
    }

    public void exportXlsTable(HSSFSheet sheet) {
        for (HtmlRow row : rows) {
            row.exportXlsRow(sheet);
        }
        for (RangeAddress rangeAddress : mergedRegions) {
            sheet.addMergedRegion(rangeAddress.mergedRegion(div.getLeft(), div.getTop()));
        }
    }

    public List<HtmlRow> getRows() {
        return rows;
    }

    public HtmlDiv getDiv() {
        return div;
    }

    public void setDiv(HtmlDiv div) {
        this.div = div;
    }
}
