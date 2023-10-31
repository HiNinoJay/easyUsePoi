package top.nino.easyUsePoi.word.module.data;

import lombok.Data;

import java.util.HashMap;
import java.util.List;

/**
 * @Author：zengzhj
 * @Date：2023/10/31 10:38
 */
@Data
public class WordJsonVo {

    /**
     * 一些提前准备的数据
     */
    private HashMap<String, String> preData;


    /**
     * 该word的组成
     */
    private List<WordModule> wordBody;

}
