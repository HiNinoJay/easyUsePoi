package top.nino.easyUsePoi.word.constant.paragraph;

import org.apache.poi.xwpf.usermodel.ParagraphAlignment;

/**
 * 段落的默认设置
 * @Author：zengzhj
 * @Date：2023/10/30 12:37
 */
public class ParagraphDefaultSetting {

    /**
     * 首行缩进
     */
    public static final Integer DEFAULT_FIRST_LINE_INDENT = 0;

    /**
     * 行间距
     */
    public static final Double DEFAULT_SPACING_BETWEEN = 1.5;


    /**
     * 正文默认朝向 靠左
     */
    public static final ParagraphAlignment DEFAULT_MAIN_BODY_ALIGN = ParagraphAlignment.LEFT;

    /**
     * 小标题默认朝向 靠左
     */
    public static final ParagraphAlignment DEFAULT_TITLE_ALIGN = ParagraphAlignment.LEFT;

    /**
     * 公告/封面 等等 默认朝向 居中
     */
    public static final ParagraphAlignment DEFAULT_ANNOUNCEMENT_ALIGN = ParagraphAlignment.CENTER;


    /**
     * 返回poi的对齐方向类
     * @param align
     * @return
     */
    public static ParagraphAlignment getAlignByString(String align) {
        if("left".equals(align)) {
            return ParagraphAlignment.LEFT;
        }
        if("center".equals(align)) {
            return ParagraphAlignment.CENTER;
        }
        return ParagraphAlignment.RIGHT;
    }

}
