package com.project.table;


import com.aspose.words.*;
import com.project.table.DemoService;
import org.springframework.stereotype.Service;

import java.awt.*;
import java.io.InputStream;
import java.util.Objects;


/**
 * @Author Ask
 * @Description TODO
 * @date 2022/2/7 13:52
 */

@Service
public class DemoServiceImpl implements DemoService {

    private final static String HOME = "C:\\Users\\49928\\Desktop\\";

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
        doc.save(HOME + "表格测试.docx");
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
    public void delOneCell(int row, int cellCount) throws Exception {
        Document doc = new Document(Objects.requireNonNull(this.getClass().getClassLoader().getResourceAsStream("TableTest.docx")));
        //获取表格
        TableCollection table = doc.getFirstSection().getBody().getTables();
        for (int i = 0; i < table.getCount(); i++) {
            Table table1 = table.get(i);
            Row row1 = (Row) table1.getChild(6, row - 1, true);
            Cell cell = row1.getCells().get(cellCount);
            cell.remove();
        }
        doc.save(HOME + "删除单个单元格文件.docx");
    }


    @Override
    public void margeOneCell(int rowNum) throws Exception {
        Document doc = readTable();
        TableCollection tables = doc.getFirstSection().getBody().getTables();
        //合并单元格
        Row row = (Row) tables.get(0).getChild(NodeType.ROW, rowNum - 1, true);
        Row row1 = (Row) tables.get(0).getChild(NodeType.ROW, rowNum , true);
        //此时进行合并会导致后后单元格数据丢失，需要将单元格数据提取出来复制拷贝至主单元格，此处不进行赘述
        mergeCells(row.getFirstCell(), row1.getFirstCell());
        doc.save(HOME + "合并单元格后文件.docx");
    }


    /**
     * 合并现有文件的指定的单元格
     *
     * @param startCell 起始单元格
     * @param endCell   结束单元格
     */
    public static void mergeCells(Cell startCell, Cell endCell) {
        Table parentTable = startCell.getParentRow().getParentTable();
        Point startCellPos = new Point(startCell.getParentRow().indexOf(startCell), parentTable.indexOf(startCell.getParentRow()));
        Point endCellPos = new Point(endCell.getParentRow().indexOf(endCell), parentTable.indexOf(endCell.getParentRow()));
        Rectangle mergeRange = new Rectangle(Math.min(startCellPos.x, endCellPos.x), Math.min(startCellPos.y, endCellPos.y), Math.abs(endCellPos.x - startCellPos.x) + 1,
                Math.abs(endCellPos.y - startCellPos.y) + 1);
        for (Row row : parentTable.getRows()) {

            for (Cell cell : row.getCells()) {
                Point currentPos = new Point(row.indexOf(cell), parentTable.indexOf(row));
                if (mergeRange.contains(currentPos)) {
                    if (currentPos.x == mergeRange.x) {
                        cell.getCellFormat().setHorizontalMerge(CellMerge.FIRST);
                    } else {
                        cell.getCellFormat().setHorizontalMerge(CellMerge.PREVIOUS);
                    }

                    if (currentPos.y == mergeRange.y) {
                        cell.getCellFormat().setVerticalMerge(CellMerge.FIRST);
                    } else {

                        cell.getCellFormat().setVerticalMerge(CellMerge.PREVIOUS);
                    }
                }
            }
        }
    }


    @Override
    public void addNewCell() throws Exception {

    }


    @Override
    public void altTable(Document doc) throws Exception {
        // 在文件中创建一个新的表格
        Table outerTable = createTable(doc, 3, 4, "Outer Table");
        doc.getFirstSection().getBody().appendChild(outerTable);
        //在第一个单元格内创建一个新的表格
        Table innerTable = createTable(doc, 2, 2, "Inner Table");
        outerTable.getFirstRow().getFirstCell().appendChild(innerTable);
        doc.save(HOME + "修改后文件.docx");
    }


    /**
     * 创建一个表格
     *
     * @param doc       操作的文档
     * @param rowCount  表格行数
     * @param cellCount 每行表格的格数
     * @param cellText  单个表格中文字
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

