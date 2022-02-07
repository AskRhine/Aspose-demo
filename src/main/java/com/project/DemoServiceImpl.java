package com.project;

import com.aspose.words.*;
import org.springframework.stereotype.Service;

import javax.swing.filechooser.FileSystemView;
import java.io.InputStream;


/**
 * @Author Ask
 * @Date 2022/2/7
 * @Describe
 */
@Service
public class DemoServiceImpl implements DemoService {
    private final static String USER_HOME_PATH = FileSystemView.getFileSystemView().getHomeDirectory().getPath();

    @Override
    public void createTable() throws Exception {

//        InputStream inputStream = this.getClass().getClassLoader().getResourceAsStream("TableTest.docx");
//        Document doc = new Document(inputStream);

        Document doc = new Document();

        //在当前word下创建一个表格
        Table table = new Table(doc);
        //将表格附加至word文件中
        doc.getFirstSection().getBody().appendChild(table);


        //row是指当前单元格的第一列
        Row firstRow = new Row(doc);
        table.appendChild(firstRow);
        Row secondRow = new Row(doc);
        table.appendChild(secondRow);


        //cell指当前单元格某一行中的第几个单元格
        Cell firstCell1c1 = new Cell(doc);
        Cell secondCell1c2 = new Cell(doc);
        Cell firstCell2c1 = new Cell(doc);
        Cell secondCell2c2 = new Cell(doc);


        firstRow.appendChild(firstCell1c1);
        firstRow.appendChild(secondCell1c2);
        secondRow.appendChild(firstCell2c1);
        secondRow.appendChild(secondCell2c2);


        //paragraph是单元格内容中的段落
        Paragraph paragraph1t1 = new Paragraph(doc);
        firstCell1c1.appendChild(paragraph1t1);
        Paragraph paragraph1t2 = new Paragraph(doc);

        firstCell1c1.appendChild(paragraph1t2);
        Paragraph paragraph2t1 = new Paragraph(doc);
        secondCell1c2.appendChild(paragraph2t1);

        Paragraph paragraph3t1 = new Paragraph(doc);
        firstCell2c1.appendChild(paragraph3t1);
        Paragraph paragraph3t2 = new Paragraph(doc);
        secondCell2c2.appendChild(paragraph3t2);


        //run是指单个单元格中一个段落中的内容
        Run run1r1 = new Run(doc, "这是第一个单元格,第一行内容");
        paragraph1t1.appendChild(run1r1);
        Run run1r2 = new Run(doc, "这是第一个单元格，第二行内容");
        paragraph1t2.appendChild(run1r2);
        Run run2r2 = new Run(doc, "这是第二个单元格，第一行内容");
        paragraph2t1.appendChild(run2r2);
        Run run3r1 = new Run(doc, "这是第三个单元格，第一行内容");
        paragraph3t1.appendChild(run3r1);
        Run run3r2 = new Run(doc, "这是第四个单元格，第一行内容");
        paragraph3t2.appendChild(run3r2);

        System.out.println(USER_HOME_PATH);

        doc.save(USER_HOME_PATH + "/表格测试.docx");

    }


    @Override
    public Document readTable() throws Exception {
        //获取文件
        InputStream inputStream = this.getClass().getClassLoader().getResourceAsStream("TableTest.docx");
        Document doc = new Document(inputStream);
        //获取该word下所有表格
        TableCollection tables = doc.getFirstSection().getBody().getTables();
        for (int i = 0; i < tables.getCount(); i++) {
            System.out.println("第" + (i + 1) + "个表格\n");
            RowCollection rows = tables.get(i).getRows();
            for (int j = 0; j < rows.getCount(); j++) {
                System.out.println("表格第" + (j + 1) + "行");
                CellCollection cells = rows.get(j).getCells();
                for (int k = 0; k < cells.getCount(); k++) {
                    String cellText = cells.get(k).toString(SaveFormat.TEXT).trim();
                    System.out.println("表格第" + (j + 1) + "行第" + (k + 1) + "个单元,内容:\n" + cellText + "\n");
                }
            }
        }
        return doc;
    }


    @Override
    public void altTable(Document doc) throws Exception {
        // 在文件中创建一个新的表格
        Table outerTable = createTable(doc, 3, 4, "Outer Table");
        doc.getFirstSection().getBody().appendChild(outerTable);
        //在第一个单元格内创建一个新的表格
        Table innerTable = createTable(doc, 2, 2, "Inner Table");
        outerTable.getFirstRow().getFirstCell().appendChild(innerTable);
        doc.save(USER_HOME_PATH + "/修改后文件.docx");
    }


    /**
     * 快速创建一个表格
     * 实际情况下cellText可以更改为String数组或是集合进行存储，方便快速进行生成表格
     *
     * @param doc
     * @param rowCount
     * @param cellCount
     * @param cellText
     * @return
     * @throws Exception
     */
    private Table createTable(Document doc, int rowCount, int cellCount, String cellText) throws Exception {
        DocumentBuilder documentBuilder = new DocumentBuilder(doc);
        //创建一个新的表格
        Table table = documentBuilder.startTable();
        for (int rowId = 1; rowId <= rowCount; rowId++) {
            for (int cellId = 1; cellId <= cellCount; cellId++) {
                documentBuilder.insertCell();
                documentBuilder.writeln(cellText);
            }
            //结束创建表格
            documentBuilder.endRow();
        }
        //结束创建表格
        documentBuilder.endTable();
        return table;
    }


}

