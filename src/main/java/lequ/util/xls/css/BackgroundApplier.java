package lequ.util.xls.css;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellStyle;

import java.util.HashMap;
import java.util.Map;

public class BackgroundApplier implements CssApplier {


    public Map<String, String> parse(Map<String, String> style) {
        Map<String, String> mapRtn = new HashMap<String, String>();
        String bg = style.get(BACKGROUND);
        String bgColor = null;
        if (StringUtils.isNotBlank(bg)) {
            for (String bgAttr : bg.split("(?<=\\)|\\w|%)\\s+(?=\\w)")) {
                bgColor = HtmlUtils.processColor(bgAttr);
                if (bgColor != null) {
                    mapRtn.put(BACKGROUND_COLOR, bgColor);
                    break;
                }
            }
        }
        bg = style.get(BACKGROUND_COLOR);
        bgColor = HtmlUtils.processColor(bg);
        if (StringUtils.isNotBlank(bg) && (bgColor != null)) {
            mapRtn.put(BACKGROUND_COLOR, bgColor);
        }
        if (bgColor != null) {
            bgColor = mapRtn.get(BACKGROUND_COLOR);
            if ("#ffffff".equals(bgColor)) {
                mapRtn.remove(BACKGROUND_COLOR);
            }
        }
        return mapRtn;
    }


    public void apply(HSSFWorkbook workbook, HSSFCellStyle cellStyle, Map<String, String> style) {
        String bgColor = style.get(BACKGROUND_COLOR);
        if (StringUtils.isNotBlank(bgColor)) {
            cellStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
            cellStyle.setFillForegroundColor(
                    HtmlUtils.parseColor(workbook, bgColor).getIndex());
        }
    }
}
