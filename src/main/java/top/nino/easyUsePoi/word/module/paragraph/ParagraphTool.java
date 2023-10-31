package top.nino.easyUsePoi.word.module.paragraph;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Component;
import top.nino.easyUsePoi.word.constant.paragraph.ParagraphDefaultSetting;
import top.nino.easyUsePoi.word.module.data.WordModule;
import top.nino.easyUsePoi.word.module.data.WordText;
import top.nino.easyUsePoi.word.module.paragraph.picture.PictureTool;
import top.nino.easyUsePoi.word.module.paragraph.text.TextTool;

import java.util.HashMap;

/**
 * @Author：zengzhj
 * @Date：2023/10/30 12:48
 */
@Component
@Slf4j
public class ParagraphTool {


    @Autowired
    private PictureTool pictureTool;

    @Autowired
    private TextTool textTool;


    /**
     * 创建一个 段落 并设置该段落的 样式
     *
     * @param xwpfDocument       该 word 文件
     * @param paragraphAlignment 该段落的 对齐方向，如 居中 ParagraphAlignment.CENTER
     * @param firstLineIndent    首行缩进 如400
     * @param spacingBetween     行间距 如1.5
     */
    public XWPFParagraph drawParagraph(XWPFDocument xwpfDocument,
                                       ParagraphAlignment paragraphAlignment,
                                       Integer firstLineIndent, Double spacingBetween) {
        // 创建段落
        XWPFParagraph paragraph = xwpfDocument.createParagraph();
        // 对齐方式
        paragraph.setAlignment(paragraphAlignment);
        // 首行缩进   400 即是0.71cm   2字符
        paragraph.setFirstLineIndent(firstLineIndent);
        // 行间距1.5
        paragraph.setSpacingBetween(spacingBetween);
        // paragraphX.setSpacingLineRule(LineSpacingRule.AT_LEAST);
        return paragraph;
    }

    /**
     * 返回一个正文的默认段落 靠左，首行缩进为0，行间距为1.5
     * @param xwpfDocument
     * @return
     */
    public XWPFParagraph drawDefaultMainBodyParagraph(XWPFDocument xwpfDocument) {
        return drawParagraph(xwpfDocument, ParagraphDefaultSetting.DEFAULT_MAIN_BODY_ALIGN,
                ParagraphDefaultSetting.DEFAULT_FIRST_LINE_INDENT, ParagraphDefaultSetting.DEFAULT_SPACING_BETWEEN);
    }

    /**
     * 返回一个公告标题的默认段落 居中，首行缩进为0，行间距为1.5
     * @param xwpfDocument
     * @return
     */
    public XWPFParagraph drawDefaultAnnouncementParagraph(XWPFDocument xwpfDocument) {
        return drawParagraph(xwpfDocument, ParagraphDefaultSetting.DEFAULT_ANNOUNCEMENT_ALIGN,
                ParagraphDefaultSetting.DEFAULT_FIRST_LINE_INDENT, ParagraphDefaultSetting.DEFAULT_SPACING_BETWEEN);
    }

    /**
     * 给段落 增加一张图片
     *
     * @param paragraph
     * @param pictureUrl
     */
    public void drawDefaultPngPicture(XWPFParagraph paragraph, String pictureUrl) {
        pictureTool.drawDefaultPngPicture(paragraph, pictureUrl);
    }

    /**
     * 默认 正文 文本格式： 宋体，四号字体大小，黑色，不加粗，默认一个回车，结束不分页
     * @param paragraph
     * @param text
     */
    public void drawDefaultMainBodyText(XWPFParagraph paragraph, String text) {
        textTool.drawDefaultMainBodyText(paragraph, text);
    }

    /**
     * 默认 小标题 文本格式： 黑体，三号字体大小，黑色，加粗，默认一个回车，结束不分页
     * @param paragraph
     * @param text
     */
    public void drawDefaultTitleText(XWPFParagraph paragraph, String text) {
        textTool.drawDefaultTitleText(paragraph, text);
    }

    /**
     * 默认 居中公告标题 文本格式： 黑体，三号字体大小，黑色，加粗，默认一个回车，结束不分页
     * @param paragraph
     * @param text
     */
    public void drawDefaultAnnouncementText(XWPFParagraph paragraph, String text) {
        textTool.drawDefaultAnnouncementText(paragraph, text);
    }

    /**
     * 可变化 公告文字 的颜色 和 字体大小
     * @param paragraph
     * @param text
     * @param colorHex
     * @param fontSize
     * @param boldFlag
     */
    public void drawAnnouncementTextColorAndSize(XWPFParagraph paragraph, String text, String colorHex, Integer fontSize, boolean boldFlag) {
        textTool.drawAnnouncementTextColorAndSize(paragraph, text, colorHex, fontSize, boldFlag);
    }

    /**
     * 通过 wordModule 生成 一个 段落
     * @param xwpfDocument
     * @param wordModule
     * @return
     */
    public XWPFParagraph drawParagraph(XWPFDocument xwpfDocument, WordModule wordModule) {
        return drawParagraph(xwpfDocument, ParagraphDefaultSetting.getAlignByString(wordModule.getAlign()),
                Integer.parseInt(wordModule.getFirstLineIndet()), Double.parseDouble(wordModule.getSpacingBetween()));
    }

    /**
     * 在 段落 中 添加多段文字
     * @param paragraph
     * @param preData
     */
    public void drawText(XWPFParagraph paragraph, WordText textPo, HashMap<String, String> preData) {
        textTool.drawText(paragraph, textPo, preData);
    }
}
