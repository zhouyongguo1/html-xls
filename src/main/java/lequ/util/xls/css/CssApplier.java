package lequ.util.xls.css;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.util.Map;

public interface CssApplier {
    String PATTERN_LENGTH = "\\d*\\.?\\d+\\s*(?:em|ex|cm|mm|q|in|pt|pc|px)?";
    String STYLE = "style";
    // direction
    String TOP = "top";
    String RIGHT = "right";
    String BOTTOM = "bottom";
    String LEFT = "left";
    String WIDTH = "width";
    String HEIGHT = "height";
    String COLOR = "color";
    String BORDER = "border";
    String CENTER = "center";
    String JUSTIFY = "justify";
    String MIDDLE = "middle";
    String FONT = "font";
    String FONT_STYLE = "font-style";
    String FONT_WEIGHT = "font-weight";
    String FONT_SIZE = "font-size";
    String FONT_FAMILY = "font-family";
    String ITALIC = "italic";
    String BOLD = "bold";
    String NORMAL = "normal";
    String TEXT_ALIGN = "text-align";
    String VETICAL_ALIGN = "vertical-align";
    String BACKGROUND = "background";
    String BACKGROUND_COLOR = "background-color";

    Map<String, String> parse(Map<String, String> style);

    void apply(HSSFWorkbook workbook, HSSFCellStyle cellStyle, Map<String, String> style);
}
