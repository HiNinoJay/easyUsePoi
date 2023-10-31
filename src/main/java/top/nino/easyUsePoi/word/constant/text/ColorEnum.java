package top.nino.easyUsePoi.word.constant.text;

import lombok.AllArgsConstructor;
import lombok.Getter;

/**
 * @Author：zengzhj
 * @Date：2023/10/30 12:31
 */
@Getter
@AllArgsConstructor
public enum ColorEnum {

    BLACK("黑色", "000000"),
    WHITE("白色", "FFFFFF"),
    CUSTOM_COLOR("青蓝色", "08abac");

    private final String name;
    private final String hexCode;

}
