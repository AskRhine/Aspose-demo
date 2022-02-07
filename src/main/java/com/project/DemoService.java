package com.project;

import com.aspose.words.Document;

/**
 * @Author Ask
 * @Date 2022/2/7
 * @Describe
 */
public interface DemoService {
    /**
     * 创建一个表格
     *
     * @throws Exception
     */
    void createTable() throws Exception;

    /**
     * 读取表格内容并将其返回为Aspose.Word.Document对象
     *
     * @return
     * @throws Exception
     */
    Document readTable() throws Exception;

    /**
     * 对表格内容进行修改
     *
     * @param doc
     * @throws Exception
     */
    void altTable(Document doc) throws Exception;
}
