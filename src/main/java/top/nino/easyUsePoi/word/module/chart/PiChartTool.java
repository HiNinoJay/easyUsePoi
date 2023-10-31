package top.nino.easyUsePoi.word.module.chart;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xwpf.usermodel.XWPFChart;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTDLbls;
import org.springframework.stereotype.Component;
import top.nino.easyUsePoi.word.module.data.WordModule;

import java.io.IOException;

/**
 * @Author：zengzhj
 * @Date：2023/10/30 15:25
 */
@Component
public class PiChartTool {

    /*
     * @Description:diy饼状图,与模板设计稍有出入,但是关键数据展示成功
     * @Param: XWPFDocument docxDocument, String titleText : 饼状图名称, String xAxisName : x轴名称, String[] xAxisData : x轴数据, String yAxisName, Integer[] yAxisData, String titleName, PresetColor color : 柱状图的颜色
     * @Return:
     * @DateTime: 11:28 2023/10/30
     * @author: dingchy
     */
    public void drawPieChart(XWPFDocument docxDocument, String charTitle, String[] xAxisData, Integer[] yAxisData) {
        XWPFChart chart = null;
        try {
            chart = docxDocument.createChart(15 * Units.EMU_PER_CENTIMETER, 10 * Units.EMU_PER_CENTIMETER);
        } catch (InvalidFormatException e) {
            throw new RuntimeException(e);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
        chart.setTitleText(charTitle);
        chart.setTitleOverlay(false);
        XDDFChartLegend orAddLegend = chart.getOrAddLegend();
        orAddLegend.setPosition(LegendPosition.TOP_RIGHT);

        XDDFCategoryDataSource xAxisSource = XDDFDataSourcesFactory.fromArray(xAxisData); // 设置分类数据
        XDDFNumericalDataSource<Integer> yAxisSource = XDDFDataSourcesFactory.fromArray(yAxisData); // 设置值数据

        // 7、创建饼图对象,饼状图不需要X,Y轴,只需要数据集即可
        XDDFPieChartData pieChart = (XDDFPieChartData) chart.createData(ChartTypes.PIE, null, null);

        // 8、加载饼图数据集
        XDDFPieChartData.Series pieSeries = (XDDFPieChartData.Series) pieChart.addSeries(xAxisSource, yAxisSource);
//        pieSeries.setTitle("粉丝数", null); // 系列提示标题
        // 9、绘制饼图
        chart.plot(pieChart);

        CTDLbls dLbls = chart.getCTChart().getPlotArea().getPieChartArray(0).getSerArray(0).addNewDLbls();
        dLbls.addNewShowVal().setVal(false);//不显示值
        dLbls.addNewShowLegendKey().setVal(false);
        dLbls.addNewShowCatName().setVal(true);//类别名称
        dLbls.addNewShowSerName().setVal(false);//不显示系列名称
        dLbls.addNewShowPercent().setVal(true);//显示百分比
        dLbls.addNewShowLeaderLines().setVal(true); //显示引导线
    }

    public void drawPieChart(XWPFDocument xwpfDocument, WordModule wordModule) {
    }
}
