package top.nino.easyUsePoi.word.module.data;

import lombok.Data;

import java.util.List;

/**
 * @Author：zengzhj
 * @Date：2023/10/31 10:46
 */
@Data
public class WordText {
    private String fontFamily;
    private String fontSize;
    private String boldFlag;
    private String color;
    private List<String> content;
}
