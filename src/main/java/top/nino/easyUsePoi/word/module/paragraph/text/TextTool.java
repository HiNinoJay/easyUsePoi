package top.nino.easyUsePoi.word.module.paragraph.text;


import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xwpf.usermodel.BreakType;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.springframework.stereotype.Component;
import top.nino.easyUsePoi.word.constant.text.ColorEnum;
import top.nino.easyUsePoi.word.constant.text.FontFamilyEnum;
import top.nino.easyUsePoi.word.constant.text.FontSizeEnum;
import top.nino.easyUsePoi.word.module.data.WordText;

import java.util.HashMap;

/**
 * @Author：zengzhj
 * @Date：2023/10/30 13:32
 */
@Component
public class TextTool {


    /**
     * 给段落 增加一段文字 并且 设置该段文字 的样式
     *
     * @param paragraph  该 word 文件 的 一个段落
     * @param text       该段落的文本信息
     * @param fontFamily 字体类型
     * @param fontSize   字体大小
     * @param color      颜色
     * @param bold       是否加粗
     * @param returnNum  回车数量
     * @param pageEnd    是否开启下一页
     */
    public void drawText(XWPFParagraph paragraph, String text,
                         String fontFamily, Integer fontSize, String color, Boolean bold,
                         Integer returnNum, boolean pageEnd) {
        XWPFRun run = paragraph.createRun();
        //段落内容
        run.setText(text);
        for (Integer i = 0; i < returnNum; i++) {
            run.addBreak();
        }
        run.setFontFamily(fontFamily);
        run.setFontSize(fontSize);
        run.setBold(bold);
        run.setColor(color);
        //默认段后间距为10榜
//        paragraph.setSpacingAfter(200);
        if (pageEnd) {
            run.addBreak(BreakType.PAGE);
        }
    }

    /**
     * 默认 正文 文本格式： 宋体，四号字体大小，黑色，不加粗，默认一个回车，结束不分页
     * @param paragraph
     * @param text
     */
    public void drawDefaultMainBodyText(XWPFParagraph paragraph, String text) {
        drawText(paragraph, text, FontFamilyEnum.SONG.getName(), FontSizeEnum.FOUR.getSizeInPoints(), ColorEnum.BLACK.getHexCode(), false, 1, false);
    }

    /**
     * 默认 小标题 文本格式： 黑体，三号字体大小，黑色，加粗，默认一个回车，结束不分页
     * @param paragraph
     * @param text
     */
    public void drawDefaultTitleText(XWPFParagraph paragraph, String text) {
        drawText(paragraph, text, FontFamilyEnum.BLACK.getName(), FontSizeEnum.THREE.getSizeInPoints(), ColorEnum.BLACK.getHexCode(), false, 1, false);
    }

    /**
     * 默认 居中公告标题 文本格式： 黑体，三号字体大小，黑色，加粗，默认一个回车，结束不分页
     * @param paragraph
     * @param text
     */
    public void drawDefaultAnnouncementText(XWPFParagraph paragraph, String text) {
        drawText(paragraph, text, FontFamilyEnum.BLACK.getName(), FontSizeEnum.THREE.getSizeInPoints(), ColorEnum.BLACK.getHexCode(), false, 1, false);
    }

    public void drawAnnouncementTextColorAndSize(XWPFParagraph paragraph, String text, String colorHex, Integer fontSize, boolean boldFlag) {
        drawText(paragraph, text, FontFamilyEnum.BLACK.getName(), fontSize, colorHex, boldFlag, 1, false);
    }


    public void drawText(XWPFParagraph paragraph, WordText textPoList, HashMap<String, String> preData) {
        textPoList.getContent().forEach(preText -> {
            String newText = handleText(preText, preData);
            drawText(paragraph, newText,
                    textPoList.getFontFamily(), Integer.parseInt(textPoList.getFontSize()), textPoList.getColor(), Boolean.parseBoolean(textPoList.getBoldFlag()),
                    1, false);
        });
    }

    /**
     * 要去做替换
     * @param preText
     * @param preData
     * @return
     */
    private String handleText(String preText, HashMap<String, String> preData) {
        int leftIndex = -1;
        int rightIndex = -1;
        boolean startFlag = false;
        StringBuilder newText = new StringBuilder();
        for(int i = 0; i < preText.length(); i++) {
            char c = preText.charAt(i);
            if(c == '{') {
                leftIndex = i;
                startFlag = true;
            } else if(c == '}') {
                rightIndex = i;
                String preString = preText.substring(leftIndex + 1, rightIndex);
                String newString = preData.get(preString);
                if(StringUtils.isNotBlank(newString)) {
                    newText.append(newString);
                } else {
                    newText.append("{").append(preString).append("}");
                }
                startFlag = false;
            } else {
                if(!startFlag) {
                    newText.append(c);
                }
            }
        }
        return newText.toString();
    }
}
