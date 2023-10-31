package top.nino.easyUsePoi.word.module.chart;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xddf.usermodel.PresetColor;
import org.apache.poi.xddf.usermodel.XDDFColor;
import org.apache.poi.xddf.usermodel.XDDFSolidFillProperties;
import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xwpf.usermodel.XWPFChart;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.springframework.stereotype.Component;
import top.nino.easyUsePoi.word.module.data.WordModule;

import java.io.IOException;

/**
 * @Author：zengzhj
 * @Date：2023/10/30 15:21
 */
@Component
public class BarChartTool {
    /*
     * @Description:重载创建柱状图方法
     * @Param: XWPFDocument docxDocument
     *         String chartTile : 图标名称,
     *         String xAxisName : x轴名称,
     *         String[] xAxisData : x轴数据,
     *         String yAxisName, Integer[] yAxisData : xy轴数据,
     *         String titleName,
     *         PresetColor color : 柱状图的颜色
     *         LegendPosition location:图例位置(上下左右)
     *         AxisPosition valueAxis : 值轴
     *         AxisPosition classificationAxis : 分类轴
     *         AxisCrossBetween axisCrossBetween : 图柱位置 eg:居中
     *         BarDirection barDirection : 柱状图方向
     * @Return:
     * @DateTime: 15:30 2023/10/30
     * @author: dingchy
     */
    public void drawBarChart(XWPFDocument docxDocument, String chartTile, String xAxisName, String[] xAxisData, String yAxisName, Integer[] yAxisData, String titleName, PresetColor color,
                               LegendPosition location, AxisPosition classificationAxis, AxisPosition valueAxis, AxisCrossBetween axisCrossBetween, BarDirection barDirection) {
        // 1、创建chart图表对象,抛出异常

        XWPFChart chart = null;
        try {
            chart = docxDocument.createChart(15 * Units.EMU_PER_CENTIMETER, 10 * Units.EMU_PER_CENTIMETER);
        } catch (InvalidFormatException | IOException e) {
            throw new RuntimeException(e);
        }
        // 2、图表相关设置

        chart.setTitleText(chartTile); // 图表标题

        // 3、图例设置

        XDDFChartLegend legend = chart.getOrAddLegend();
        legend.setPosition(location); // 图例位置:上下左右

        // 4、X轴(分类轴)相关设置

        XDDFCategoryAxis xAxis = chart.createCategoryAxis(AxisPosition.BOTTOM); // 创建X轴,并且指定位置
        xAxis.setTitle(xAxisName); // x轴标题
        XDDFCategoryDataSource xAxisSource = XDDFDataSourcesFactory.fromArray(xAxisData); // 设置X轴数据

        // 5、Y轴(值轴)相关设置

        XDDFValueAxis yAxis = chart.createValueAxis(AxisPosition.LEFT); // 创建Y轴,指定位置
        yAxis.setTitle(yAxisName); // Y轴标题
        yAxis.setCrossBetween(AxisCrossBetween.BETWEEN); // 设置图柱的位置:BETWEEN居中
        XDDFNumericalDataSource<Integer> yAxisSource = XDDFDataSourcesFactory.fromArray(yAxisData); // 设置Y轴数据

        // 6、创建柱状图对象

        XDDFBarChartData barChart = (XDDFBarChartData) chart.createData(ChartTypes.BAR, xAxis, yAxis);
        barChart.setBarDirection(BarDirection.COL); // 设置柱状图的方向:BAR横向,COL竖向,默认是BAR

        // 7、加载柱状图数据集

        XDDFBarChartData.Series barSeries = (XDDFBarChartData.Series) barChart.addSeries(xAxisSource, yAxisSource);
        barSeries.setTitle(titleName, null); // 图例标题
        barSeries.setFillProperties(new XDDFSolidFillProperties(XDDFColor.from(color)));

        // 8、绘制柱状图

        chart.plot(barChart);
    }


    /*
     * @Description:重载创建柱状图方法
     * @Param: XWPFDocument docxDocument, String chartTile : 图标名称, String xAxisName : x轴名称, String[] xAxisData : x轴数据, String yAxisName, Integer[] yAxisData, String titleName, PresetColor color : 柱状图的颜色
     * @Return:
     * @DateTime: 19:30 2023/10/27
     * @author: dingchy
     */
    public void drawDefaultBarChart(XWPFDocument docxDocument, String chartTile, String xAxisName, String[] xAxisData, String yAxisName, Integer[] yAxisData, String titleName, PresetColor color) {
        //默认柱状图
        drawBarChart(docxDocument, chartTile, xAxisName, xAxisData, yAxisName, yAxisData, titleName, PresetColor.BLUE, LegendPosition.TOP, AxisPosition.BOTTOM, AxisPosition.LEFT, AxisCrossBetween.BETWEEN, BarDirection.COL);
    }

    public void drawBarChart(XWPFDocument xwpfDocument, WordModule wordModule) {
    }
}
