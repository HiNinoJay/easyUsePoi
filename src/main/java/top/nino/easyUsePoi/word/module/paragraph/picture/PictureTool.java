package top.nino.easyUsePoi.word.module.paragraph.picture;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.Document;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.springframework.stereotype.Component;

import java.io.BufferedInputStream;
import java.io.InputStream;
import java.net.URL;

/**
 * @Author：zengzhj
 * @Date：2023/10/30 13:13
 */
@Component
@Slf4j
public class PictureTool {

    private final static int PICTURE_WIDTH = Units.toEMU(425);
    private final static int PICTURE_HEIGHT = Units.toEMU(242);
    private final static String PICTURE_NAME = "默认图片名称";

    /**
     * 给段落 增加一张图片
     *
     * @param paragraph
     * @param pictureUrl
     */
    public void drawPicture(XWPFParagraph paragraph, String pictureUrl,
                            int pictureType, String fileName,
                            int width, int height) {
        XWPFRun run = paragraph.createRun();
        try {
            URL url = new URL(pictureUrl);
            InputStream in = new BufferedInputStream(url.openStream());
            run.addPicture(in, pictureType, fileName, width, height);
        } catch (Exception e) {
            log.error("该word插入图片失败：{}", pictureUrl);
        }
    }

    /**
     * 给段落 增加一张图片
     *
     * @param paragraph
     * @param pictureUrl
     */
    public void drawDefaultPngPicture(XWPFParagraph paragraph, String pictureUrl) {
        drawPicture(paragraph, pictureUrl, Document.PICTURE_TYPE_PNG, PICTURE_NAME + ".png", PICTURE_WIDTH, PICTURE_HEIGHT);
    }
}
