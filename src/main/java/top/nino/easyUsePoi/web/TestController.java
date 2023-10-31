package top.nino.easyUsePoi.web;

import com.alibaba.fastjson.JSON;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.ClassPathResource;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;
import top.nino.easyUsePoi.word.NinoWordUtil;
import top.nino.easyUsePoi.word.module.data.WordJsonVo;

import java.io.InputStreamReader;

/**
 * @Author：zengzhj
 * @Date：2023/10/31 20:20
 */
@RequestMapping("/nino")
@RestController
public class TestController {

    @Autowired
    private NinoWordUtil ninoWordUtil;


    @GetMapping("/testWord")
    public void test() {
        WordJsonVo wordJsonVo = ninoWordUtil.readFromLocalJson();
        XWPFDocument xwpfDocument = ninoWordUtil.constructWordByVo(wordJsonVo);
        ninoWordUtil.exportWord(xwpfDocument);
    }



}
