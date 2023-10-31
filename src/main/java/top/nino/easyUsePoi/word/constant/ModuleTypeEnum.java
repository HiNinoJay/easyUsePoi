package top.nino.easyUsePoi.word.constant;

import lombok.AllArgsConstructor;
import lombok.Getter;

/**
 * @Author：zengzhj
 * @Date：2023/10/31 11:18
 */
@Getter
@AllArgsConstructor
public enum ModuleTypeEnum {

    PARAGRAPH("paragraph"),
    TABLE("table"),
    BAR_CHART("barChart"),
    Pi_CHART("piChart");

    private final String name;

}
