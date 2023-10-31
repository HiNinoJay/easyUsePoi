package top.nino.easyUsePoi.word.constant.text;

import lombok.AllArgsConstructor;
import lombok.Getter;

/**
 * @author zengzhongjie
 */
@Getter
@AllArgsConstructor
public enum FontSizeEnum {

    SMALL_PRIMARY("小初", 36, 12.70),
    ONE("一号", 26, 9.17),
    SMALL_ONE("小一", 24, 8.47),
    TWO("二号", 22, 7.76),
    SMALL_TWO("小二", 18, 6.35),
    THREE("三号", 16, 5.64),
    SMALL_THREE("小三", 15, 5.29),
    FOUR("四号", 14, 4.94),
    SMALL_FOUR("小四", 12, 4.32),
    SMALL_FIVE("小五", 9, 3.18);

    /**
     * 字号
     */
    private final String name;
    /**
     * 字体磅
     */
    private final Integer sizeInPoints;

    /**
     * 毫米数
     */
    private final Double sizeInMillimeters;

}

