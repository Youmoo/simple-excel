package com.github.youmoo.excel;

import jxl.Workbook;
import jxl.format.Alignment;
import jxl.format.Border;
import jxl.format.BorderLineStyle;
import jxl.format.Colour;
import jxl.write.*;

import java.io.*;
import java.lang.Number;
import java.util.List;

/**
 * Excel表格生成类
 * <p/>
 * excel表格坐标从0开始，所有接收col或row参数的方法，都是绝对坐标
 *
 * @author youmoo
 * @since 2013年12月1日
 */
public class ExcelFile {

    protected WritableWorkbook wwb = null;//工作薄
    protected WritableSheet ws = null;//工作表,只提供一个,可以自己扩展
    protected String fileName;//文件名,工作表名
    protected File deliveryFile = null;//临时文件
    protected int row = 0;//第几行，从0开始
    protected boolean freeze = true;//标题冻结
    protected int headRow = 0;//冻结的行
    protected int columnWidth = 25;//列宽
    private WritableCellFormat HEAD_FORMAT = headFormat();//head样式
    private WritableCellFormat BODY_FORMAT = bodyFormat();//body样式
    public static boolean NO_FREEZE = false;

    /**
     * @param fileName 文件名称
     */
    public ExcelFile(String fileName) throws IOException, WriteException {
        init(fileName, this.columnWidth, this.freeze);
    }

    /**
     * @param fileName    文件名称
     * @param columnWidth 列宽
     * @throws Exception
     */
    public ExcelFile(String fileName, int columnWidth) throws IOException, WriteException {
        init(fileName, columnWidth, this.freeze);
    }

    /**
     * @param fileName 文件名
     * @param freeze   是否对标题进行冻结
     * @throws Exception
     */
    public ExcelFile(String fileName, boolean freeze) throws IOException, WriteException {
        init(fileName, this.columnWidth, freeze);
    }

    /**
     * @param fileName 文件名
     * @param freeze   是否对标题进行冻结
     * @throws Exception
     */
    public ExcelFile(String fileName, int columnWidth, boolean freeze) throws IOException, WriteException {
        init(fileName, columnWidth, freeze);
    }

    /**
     * 初始化excel组件
     *
     * @param fileName    文件名
     * @param columnWidth 列宽
     * @param freeze      是否对标题进行冻结
     * @throws java.io.IOException
     */
    protected void init(String fileName, int columnWidth, boolean freeze) throws IOException {
        this.fileName = fileName;
        this.columnWidth = columnWidth;
        this.freeze = freeze;
        this.deliveryFile = File.createTempFile(fileName, ".xls");
        this.wwb = Workbook.createWorkbook(deliveryFile);
        this.ws = this.wwb.createSheet(fileName, 0);
    }

    /**
     * 写入head
     *
     * @param titles 列标题
     * @return this
     * @throws jxl.write.biff.RowsExceededException
     * @throws jxl.write.WriteException
     */
    public ExcelFile writeHead(String... titles) throws WriteException {
        writeHead(titles, defaultColspans(titles.length));
        return this;
    }

    /**
     * 写入表头head
     *
     * @param titles   列标题数组
     * @param colspans 每个列标题所跨列数数组
     * @return this
     * @throws jxl.write.biff.RowsExceededException
     * @throws jxl.write.WriteException
     */
    public ExcelFile writeHead(String[] titles, int[] colspans) throws WriteException {
        if (titles.length != colspans.length) throw new RuntimeException("两个参数的长度必须相等！！");
        int start = 0;
        for (int i = 0, len = titles.length; i < len; i++) {
//            if (titles[i].equals("")) continue;//空白单元格
            Label label = new Label(start, row, titles[i], headFormat());
            this.ws.addCell(label);
            this.ws.mergeCells(start, row, start + colspans[i] - 1, row);/*2,4,3*/
            this.ws.setColumnView(i, this.columnWidth);
            start += colspans[i];
        }
        this.row++;
        this.headRow++;
        return this;
    }

    /**
     * 列标题默认不跨列
     *
     * @param len 列标题长度
     * @return int[]
     */
    private int[] defaultColspans(int len) {
        int[] result = new int[len];
        while (len-- != 0) {
            result[len] = 1;
        }
        return result;
    }

    /**
     * 写入多行
     *
     * @param list      行数据集合
     * @param rowReader 行数据读取者
     * @param <T>       行数据提供者(通常是一个bean或一个model)
     * @return this
     * @throws jxl.write.WriteException
     * @throws jxl.write.biff.RowsExceededException
     */
    public <T> ExcelFile writeRows(List<T> list, RowReader<T> rowReader) throws WriteException {
        for (T aList : list) {
            writeRow(aList, rowReader);
        }
        return this;
    }

    /**
     * 写入单行
     *
     * @param data      行数据提供者
     * @param rowReader 行数据读取者
     * @param <T>       行数据提供者的类型
     * @return this
     * @throws jxl.write.WriteException
     * @throws jxl.write.biff.RowsExceededException
     */
    public <T> ExcelFile writeRow(T data, RowReader<T> rowReader)
            throws WriteException {
        List<Object> rowData = rowReader.read(data);
        writeRow(rowData);
        return this;
    }

    /**
     * 写入单行
     *
     * @param rowData 行数据
     * @return this
     * @throws jxl.write.WriteException
     * @throws jxl.write.biff.RowsExceededException
     */
    public ExcelFile writeRow(List<Object> rowData)
            throws WriteException {
        for (int i = 0, len = rowData.size(); i < len; i++) {
            WritableCell cell;
            Object object = rowData.get(i);
            if (object == null) {
                cell = new Label(i, row, "", BODY_FORMAT);
            } else if (object instanceof Number) {
                Number number = (Number) object;
                cell = new jxl.write.Number(i, row, number.doubleValue(), BODY_FORMAT);
            } else {
                cell = new Label(i, row, object.toString(), BODY_FORMAT);
            }
            this.ws.addCell(cell);

        }
        row++;
        return this;
    }

    /**
     * (没有数据时)进行提醒
     *
     * @param msg 提示文本
     * @return this
     * @throws jxl.write.WriteException
     */
    public ExcelFile tip(String msg) throws WriteException {
        ws.addCell(new Label(0, row, msg));
        row++;
        return this;
    }

    /**
     * 返回生成的excel文件
     *
     * @return
     * @throws Exception
     */
    public File end() throws Exception {
        freeze();
        closeWritableWorkbook();
        return this.deliveryFile;
    }

    /**
     * 标题冻结
     */
    private void freeze() {
        if (freeze) {
            ws.getSettings().setVerticalFreeze(headRow);
        }
    }

    /**
     * 写入并关闭
     *
     * @throws Exception
     */
    protected void closeWritableWorkbook() throws IOException, WriteException {
        if (wwb != null) {
            wwb.write();
            wwb.close();
        }
    }

    /**
     * 表格表头样式
     *
     * @return
     * @throws jxl.write.WriteException
     */
    private static WritableCellFormat headFormat() throws WriteException {
        //设置字体
        WritableFont font = new WritableFont(WritableFont.createFont("微软雅黑"), 12, WritableFont.BOLD, false);
        font.setColour(Colour.BLUE);
        WritableCellFormat cellFormat = new WritableCellFormat(font);
        //设置边框
        cellFormat.setBorder(Border.LEFT, BorderLineStyle.THIN);
        cellFormat.setBorder(Border.TOP, BorderLineStyle.THIN);
        cellFormat.setBorder(Border.RIGHT, BorderLineStyle.THIN);
        cellFormat.setBorder(Border.BOTTOM, BorderLineStyle.THIN);
        //设置文本居中
        cellFormat.setAlignment(Alignment.CENTRE);
        cellFormat.setVerticalAlignment(VerticalAlignment.CENTRE);
        //设置背景色
        cellFormat.setBackground(Colour.PALE_BLUE);
        return cellFormat;
    }

    /**
     * 表格主体样式
     *
     * @return
     * @throws jxl.write.WriteException
     */
    private static WritableCellFormat bodyFormat() throws WriteException {
        //设置字体
        WritableFont font = new WritableFont(WritableFont.createFont("微软雅黑"), 12, WritableFont.NO_BOLD, false);
        WritableCellFormat cellFormat = new WritableCellFormat(font);
        //设置边框
        cellFormat.setBorder(Border.LEFT, BorderLineStyle.THIN);
        cellFormat.setBorder(Border.TOP, BorderLineStyle.THIN);
        cellFormat.setBorder(Border.RIGHT, BorderLineStyle.THIN);
        cellFormat.setBorder(Border.BOTTOM, BorderLineStyle.THIN);
        //设置文本居中
        cellFormat.setAlignment(Alignment.CENTRE);
        cellFormat.setVerticalAlignment(VerticalAlignment.CENTRE);
        return cellFormat;
    }

    /**
     * 工具方法,将文件写到输出流
     *
     * @param file 文件
     * @param os   输入流
     * @throws java.io.FileNotFoundException
     * @throws java.io.IOException
     */
    public static void write(File file, OutputStream os) throws IOException {
        InputStream is = new FileInputStream(file);
        byte[] b = new byte[1024];
        int len;
        while ((len = is.read(b)) > 0) {
            os.write(b, 0, len);
        }
        is.close();
        if (os != null) {
            os.flush();
            os.close();
        }
    }

}