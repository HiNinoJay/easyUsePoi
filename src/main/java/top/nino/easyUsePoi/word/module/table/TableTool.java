package top.nino.easyUsePoi.word.module.table;


import lombok.extern.slf4j.Slf4j;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.springframework.stereotype.Component;
import top.nino.easyUsePoi.word.module.data.WordModule;

import java.util.List;

/**
 * @Author：zengzhj
 * @Date：2023/10/30 15:14
 */
@Component
@Slf4j
public class TableTool {

    /**
     * 表格的默认 宽
     */
    private final static String TABLE_DEFAULT_WIDTH = "4381";

    /**
     * 表格的默认 高
     */
    private final static int TABLE_DEFAULT_HEIGHT = 547;

    public void drawTable(XWPFDocument xwpfDocument, int rows, int cols, List<List<String>> text) {
        if (CollectionUtils.isEmpty(text)) {
            return;
        }

        XWPFTable table = xwpfDocument.createTable(rows, cols);
        int dataRows = text.size();

        for (int i = 0; i < dataRows; i++) {
            // 得到第i行
            List<String> rowTextList = text.get(i);
            if(CollectionUtils.isEmpty(rowTextList)) {
                continue;
            }
            XWPFTableRow row = table.getRow(i);
            row.setHeight(TABLE_DEFAULT_HEIGHT);
            for (int j = 0; j < text.get(i).size(); j++) {
                // 得到第i行,第j个单元格
                String cellText = rowTextList.get(j);
                row.getCell(j).setWidth(TABLE_DEFAULT_WIDTH);
                row.getCell(j).setText(cellText);
                row.getCell(j).setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
            }
        }
    }


    public void drawTable(XWPFDocument xwpfDocument, WordModule wordModule) {
        drawTable(xwpfDocument, wordModule.getRows(), wordModule.getCols(), wordModule.getTableContent());
    }
}
