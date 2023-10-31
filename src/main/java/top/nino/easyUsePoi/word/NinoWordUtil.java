package top.nino.easyUsePoi.word;



import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson.JSONObject;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.ClassPathResource;
import org.springframework.stereotype.Component;
import top.nino.easyUsePoi.word.module.DocumentTool;
import top.nino.easyUsePoi.word.module.data.WordJsonVo;

import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.net.URL;
import java.net.URLConnection;

/**
 * 初步成果：手动传入参数，绘制出我们要的word样式
 * 最终成功：只需传入定义好的json，即可自动绘制出我们要的word样式
 *
 * @author zengzhongjie
 */
@Slf4j
@Component
public class NinoWordUtil {

    @Autowired
    private DocumentTool documentTool;

    /**
     * 读取json文件
     * @param json
     * @return
     */
    public WordJsonVo startFromJson(String json){
        // 通过URL去访问服务器上的资源
        URL url = null;
        try {
            url = new URL(json);
            URLConnection urlCon = url.openConnection();
            urlCon.connect();         //获取连接
            InputStream is = urlCon.getInputStream();
            BufferedReader buffer = new BufferedReader(new InputStreamReader(is));
            StringBuffer bs = new StringBuffer();
            String l = null;
            while((l = buffer.readLine()) != null){
                bs.append(l);
            }
            return JSONObject.parseObject(bs.toString(), WordJsonVo.class);
        } catch (IOException e) {
            e.printStackTrace();
        }
        return null;
    }

    public WordJsonVo readFromLocalJson() {
        try {
            ClassPathResource resource = new ClassPathResource("wordJson.json");
            BufferedReader buffer = new BufferedReader(new InputStreamReader(resource.getInputStream()));
            StringBuffer bs = new StringBuffer();
            String l = null;
            while((l = buffer.readLine()) != null){
                bs.append(l);
            }
            WordJsonVo wordInfo = JSON.parseObject(bs.toString(), WordJsonVo.class);
            return wordInfo;
        } catch (Exception e) {
            // 处理异常
            e.printStackTrace();
            return null;
        }
    }

    /**
     * 根据 WordJsonVo 生成数据
     * @param wordJsonVo
     */
    public XWPFDocument constructWordByVo(WordJsonVo wordJsonVo) {
        return documentTool.constructWordByVo(wordJsonVo);
    }


    /**
     * 自定义文件名
     * @param xwpfDocument
     * @param fileName
     */
    public void exportWord(XWPFDocument xwpfDocument, String fileName) {
        documentTool.exportWord(xwpfDocument, fileName);
    }

    /**
     * 生成该word, 文件名默认
     *
     * @param docxDocument
     */
    public void exportWord(XWPFDocument docxDocument) {
        documentTool.exportWord(docxDocument);
    }
}
