package top.nino.easyUsePoi.word.module.data;

import lombok.Data;

import java.util.List;

/**
 * @Author：zengzhj
 * @Date：2023/10/31 10:44
 */
@Data
public class WordModule {

    /**
     * word 组件 的类型：paragraph | table | barChart | piChart
     */
    private String type;

    /**
     * 对齐位置：left | center | right
     */
    private String align;

    /**
     * 首行缩进
     */
    private String firstLineIndet;

    /**
     * 行间距
     */
    private String spacingBetween;

    /**
     * 是否分页
     */
    private Boolean pageBreak;

    /**
     * 段落里的文本
     */
    private List<WordText> textList;

    /**
     * 表格
     */
    private Integer rows;
    private Integer cols;
    private List<List<String>> tableContent;

    /**
     * 图
     */
    private String chartTitle;
    private List<String> chartFieldName;

}
