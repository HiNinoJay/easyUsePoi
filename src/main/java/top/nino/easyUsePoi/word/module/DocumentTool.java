package top.nino.easyUsePoi.word.module;


import lombok.extern.slf4j.Slf4j;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.lang3.ObjectUtils;
import org.apache.poi.xddf.usermodel.PresetColor;
import org.apache.poi.xwpf.usermodel.BreakType;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Component;
import top.nino.easyUsePoi.word.constant.ModuleTypeEnum;
import top.nino.easyUsePoi.word.constant.text.ColorEnum;
import top.nino.easyUsePoi.word.constant.text.FontSizeEnum;
import top.nino.easyUsePoi.word.module.chart.BarChartTool;
import top.nino.easyUsePoi.word.module.chart.PiChartTool;
import top.nino.easyUsePoi.word.module.data.WordJsonVo;
import top.nino.easyUsePoi.word.module.data.WordModule;
import top.nino.easyUsePoi.word.module.data.WordText;
import top.nino.easyUsePoi.word.module.paragraph.ParagraphTool;
import top.nino.easyUsePoi.word.module.table.TableTool;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigInteger;
import java.time.LocalDateTime;
import java.util.HashMap;
import java.util.List;

/**
 * @Author：zengzhj
 * @Date：2023/10/30 13:02
 */
@Component
@Slf4j
public class DocumentTool {

    /**
     * A4纸张 的 宽
     */
    private final static long A4_WIDTH = 12242L;

    /**
     * A4纸张 的 高
     */
    private final static long A4_HEIGHT = 15842L;


    @Autowired
    private ParagraphTool paragraphTool;

    @Autowired
    private TableTool tableTool;

    @Autowired
    private BarChartTool barChartTool;

    @Autowired
    private PiChartTool piChartTool;


    /**
     * 创建一个 word 自定义 宽 和 高
     * @param width
     * @param height
     * @return
     */
    public XWPFDocument createNewWord(Long width, Long height) {
        XWPFDocument docxDocument = new XWPFDocument();
        CTSectPr sectPr = docxDocument.getDocument().getBody().addNewSectPr();
//        CTPageSz pgsz = sectPr.isSetPgSz() ? sectPr.getPgSz() : sectPr.addNewPgSz();
//        pgsz.setW(BigInteger.valueOf(width));
//        pgsz.setH(BigInteger.valueOf(height));
        return docxDocument;
    }

    /**
     * 创建一个word, 设置该word 的整体默认参数
     * sd
     * 比如：页面大小
     *
     * @return
     */
    public XWPFDocument createA4Word() {
        return createNewWord(A4_WIDTH, A4_HEIGHT);
    }


    /**
     * 生成一个默认首页
     * @param xwpfDocument
     * @param preData
     */
    private void constructDefaultIndexPage(XWPFDocument xwpfDocument, HashMap<String, String> preData) {

        String companyName = preData.get("companyName");
        String productName = preData.get("productName");
        String sampleNum = preData.get("collectedSampleCount");
        String url = preData.get("surfacePicture");

        String text2 = productName + "数据分析结题报告";
        String text3 = "easy Use Poi";
        String text4 = sampleNum + "例数据分析";
        String text5 = "Nino | " + LocalDateTime.now().getYear() +  "年" + LocalDateTime.now().getMonthValue() + "月";

        XWPFParagraph paragraph = drawDefaultAnnouncementParagraph(xwpfDocument);
        drawDefaultPngPicture(paragraph, url);
        drawDefaultAnnouncementText(paragraph, companyName);
        drawAnnouncementTextColorAndSize(paragraph, text2, ColorEnum.CUSTOM_COLOR.getHexCode(), FontSizeEnum.TWO.getSizeInPoints(), true);
        drawDefaultMainBodyText(paragraph, text3);
        drawDefaultMainBodyText(paragraph, text4);
        drawNewBreak(xwpfDocument, 5);
        drawDefaultMainBodyText(paragraph, text5);

        drawNewPage(xwpfDocument);
    }


    /**
     * 新开一页
     * @param docxDocument
     */
    public void drawNewPage(XWPFDocument docxDocument) {
        docxDocument.createParagraph().createRun().addBreak(BreakType.PAGE);
    }


    /**
     * 添加回车
     * @param docxDocument
     * @param breakNum
     */
    public void drawNewBreak(XWPFDocument docxDocument, Integer breakNum) {
        XWPFRun run = docxDocument.createParagraph().createRun();
        for(int i = 0; i < breakNum; i++) {
            run.addBreak();
        }
    }


    /**
     * 根据 wordModule 生成 一个段落
     * @param xwpfDocument
     * @param wordModule
     * @return
     */
    public XWPFParagraph drawParagraph(XWPFDocument xwpfDocument, WordModule wordModule) {
        return paragraphTool.drawParagraph(xwpfDocument, wordModule);
    }

    /**
     * 在 段落 中 添加多段文字
     * @param paragraph
     * @param textList
     * @param preData
     */
    public void drawText(XWPFParagraph paragraph, List<WordText> textList, HashMap<String, String> preData) {
        textList.forEach(textPo -> {
            paragraphTool.drawText(paragraph, textPo, preData);
        });
    }

    /**
     * 根据 wordModule 生成一个表格
     * @param xwpfDocument
     * @param wordModule
     */
    public void drawTable(XWPFDocument xwpfDocument, WordModule wordModule) {
        tableTool.drawTable(xwpfDocument, wordModule);
    }

    /**
     * 根据 wordModule 生成 一个 柱状图
     * @param xwpfDocument
     * @param wordModule
     */
    public void drawBarChart(XWPFDocument xwpfDocument, WordModule wordModule) {
        barChartTool.drawBarChart(xwpfDocument, wordModule);
    }

    /**
     * 根据 wordModule 生成 一个 饼状图
     * @param xwpfDocument
     * @param wordModule
     */
    public void drawPiChart(XWPFDocument xwpfDocument, WordModule wordModule) {
        piChartTool.drawPieChart(xwpfDocument, wordModule);
    }


    /**
     * 提供一个默认的正文段落：靠左，首行缩进为 0， 行间距为1.5
     * @param xwpfDocument
     * @return
     */
    public XWPFParagraph drawDefaultMainBodyParagraph(XWPFDocument xwpfDocument) {
        return paragraphTool.drawDefaultMainBodyParagraph(xwpfDocument);
    }

    /**
     * 提供一个默认的公告文字段落：居中，首行缩进为 0， 行间距为1.5
     * @param xwpfDocument
     * @return
     */
    public XWPFParagraph drawDefaultAnnouncementParagraph(XWPFDocument xwpfDocument) {
        return paragraphTool.drawDefaultAnnouncementParagraph(xwpfDocument);
    }

    /**
     * 在 段落中 新增一段 默认正文格式的文字：宋体，四号字体大小，黑色，不加粗，默认一个回车，结束不分页
     * @param paragraph
     * @param text
     */
    public void drawDefaultMainBodyText(XWPFParagraph paragraph, String text) {
        paragraphTool.drawDefaultMainBodyText(paragraph, text);
    }


    /**
     * 默认 小标题 文本格式： 黑体，三号字体大小，黑色，加粗，默认一个回车，结束不分页
     * @param paragraph
     * @param text
     */
    public void drawDefaultTitleText(XWPFParagraph paragraph, String text) {
        paragraphTool.drawDefaultTitleText(paragraph, text);
    }

    /**
     * 默认 居中公告标题 文本格式： 黑体，三号字体大小，黑色，加粗，默认一个回车，结束不分页
     * @param paragraph
     * @param text
     */
    public void drawDefaultAnnouncementText(XWPFParagraph paragraph, String text) {
        paragraphTool.drawDefaultAnnouncementText(paragraph, text);
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
        paragraphTool.drawAnnouncementTextColorAndSize(paragraph, text, colorHex, fontSize, boldFlag);
    }

    /**
     * 给段落 增加一张图片
     *
     * @param paragraph
     * @param pictureUrl
     */
    public void drawDefaultPngPicture(XWPFParagraph paragraph, String pictureUrl) {
        paragraphTool.drawDefaultPngPicture(paragraph, pictureUrl);
    }


    /**
     * 绘画一个表格
     * @param xwpfDocument
     * @param rows
     * @param cols
     * @param text
     */
    public void drawTable(XWPFDocument xwpfDocument, int rows, int cols, List<List<String>> text) {
        tableTool.drawTable(xwpfDocument, rows, cols, text);
    }

    /**
     * 绘画一个饼状图
     * @param docxDocument
     * @param charTitle
     * @param xAxisData
     * @param yAxisData
     */
    public void drawPieChart(XWPFDocument docxDocument, String charTitle, String[] xAxisData, Integer[] yAxisData) {
        piChartTool.drawPieChart(docxDocument, charTitle, xAxisData, yAxisData);
    }

    /**
     * 绘画一个 默认样式 柱状图
     * @param docxDocument
     * @param chartTile
     * @param xAxisName
     * @param xAxisData
     * @param yAxisName
     * @param yAxisData
     * @param titleName
     * @param color
     */
    public void drawDefaultBarChart(XWPFDocument docxDocument, String chartTile, String xAxisName, String[] xAxisData, String yAxisName, Integer[] yAxisData, String titleName, PresetColor color) {
        barChartTool.drawDefaultBarChart(docxDocument, chartTile, xAxisName, xAxisData, yAxisName, yAxisData, titleName, color);
    }




    /**
     * 通过 json 识别后的 vo 自动生成 word
     * @param wordJsonVo
     * @return
     */
    public XWPFDocument constructWordByVo(WordJsonVo wordJsonVo) {

        HashMap<String, String> preData = wordJsonVo.getPreData();
        if(CollectionUtils.isEmpty(wordJsonVo.getWordBody())) {
            return null;
        }

        XWPFDocument xwpfDocument = createA4Word();

        constructDefaultIndexPage(xwpfDocument, preData);

        wordJsonVo.getWordBody().forEach(wordModule -> {

            if(wordModule.getType().equals(ModuleTypeEnum.PARAGRAPH.getName())) {
                XWPFParagraph paragraph = drawParagraph(xwpfDocument, wordModule);
                // todo 画图
                drawText(paragraph, wordModule.getTextList(), preData);
            }

            if(wordModule.getType().equals(ModuleTypeEnum.TABLE.getName())) {
                drawTable(xwpfDocument, wordModule);
            }

            if(wordModule.getType().equals(ModuleTypeEnum.BAR_CHART.getName())) {
                drawBarChart(xwpfDocument, wordModule);
            }

            if(wordModule.getType().equals(ModuleTypeEnum.Pi_CHART.getName())) {
                drawPiChart(xwpfDocument, wordModule);
            }

            if(ObjectUtils.isNotEmpty(wordModule.getPageBreak()) && wordModule.getPageBreak()) {
                drawNewPage(xwpfDocument);
            }
        });
        return xwpfDocument;
    }


    /**
     * 根据传入的文件名生成word
     * @param xwpfDocument
     * @param fileName
     */
    public void exportWord(XWPFDocument xwpfDocument, String fileName) {
        String path = fileName + ".docx";
        File file = new File(path);
        FileOutputStream stream = null;
        try {
            stream = new FileOutputStream(file);
            xwpfDocument.write(stream);
        } catch (IOException e) {
            throw new RuntimeException(e);
        } finally {
            try {
                stream.close();
            } catch (IOException e) {
                log.error("请检查你是否已经打开该word文件！如果已经打开，请关闭！");
            }
        }
        log.info("word生成完成！");
    }

    /**
     * 生成该word
     *
     * @param docxDocument
     */
    public void exportWord(XWPFDocument docxDocument) {
        exportWord(docxDocument, "json生成word");
    }
}


