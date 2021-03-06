package lequ.util.xls.css;

import org.apache.commons.lang3.ArrayUtils;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellStyle;

import java.util.HashMap;
import java.util.Map;

public class AlignApplier implements CssApplier {
    public Map<String, String> parse(Map<String, String> style) {
        Map<String, String> mapRtn = new HashMap<String, String>();
        String align = style.get(TEXT_ALIGN);
        if (!ArrayUtils.contains(new String[]{LEFT, CENTER, RIGHT, JUSTIFY}, align)) {
            align = LEFT;
        }
        mapRtn.put(TEXT_ALIGN, align);
        align = style.get(VETICAL_ALIGN);
        if (!ArrayUtils.contains(new String[]{TOP, MIDDLE, BOTTOM}, align)) {
            align = MIDDLE;
        }
        mapRtn.put(VETICAL_ALIGN, align);
        return mapRtn;
    }


    public void apply(HSSFWorkbook workbook, HSSFCellStyle cellStyle, Map<String, String> style) {
        // text align
        String align = style.get(TEXT_ALIGN);
        short sAlign = CellStyle.ALIGN_LEFT;
        if (RIGHT.equals(align)) {
            sAlign = CellStyle.ALIGN_RIGHT;
        } else if (CENTER.equals(align)) {
            sAlign = CellStyle.ALIGN_CENTER;
        } else if (JUSTIFY.equals(align)) {
            sAlign = CellStyle.ALIGN_JUSTIFY;
        }
        cellStyle.setAlignment(sAlign);
        // vertical align
        align = style.get(VETICAL_ALIGN);
        sAlign = CellStyle.VERTICAL_CENTER;
        if (TOP.equals(align)) {
            sAlign = CellStyle.VERTICAL_TOP;
        } else if (BOTTOM.equals(align)) {
            sAlign = CellStyle.VERTICAL_BOTTOM;
        } else if (JUSTIFY.equals(align)) {
            sAlign = CellStyle.VERTICAL_JUSTIFY;
        }
        cellStyle.setVerticalAlignment(sAlign);
    }
}
