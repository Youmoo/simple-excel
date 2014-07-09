import com.github.youmoo.excel.ExcelFile;
import com.github.youmoo.excel.RowReader;
import org.junit.Before;
import org.junit.Test;

import java.io.File;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Random;

/**
 * @autor youmoo
 * @since 2014-07-09 下午9:51
 */
public class TestExcelFile {

    List<DataBean> dataBeans;
    String[] titles = "基本信息,保密".split(",");
    int[] titleColSpans = {2, 1};
    String[] subTitles = "用户名,年龄,生日".split(",");

    @Before
    public void setUp() {
        dataBeans = new ArrayList<DataBean>();
        Random random = new Random();

        for (int i = 0; i < 10; i++) {
            dataBeans.add(new DataBean("name " + i, random.nextInt(100), new Date()));
        }
    }

    @Test
    public void generate() throws Exception {
        File file =
                new ExcelFile("测试文件生成")
                        .writeHead(titles, titleColSpans)
                        .writeHead(subTitles)
                        .writeRows(dataBeans, new RowReader<DataBean>() {
                            @Override
                            public List<Object> read(DataBean dataBean) {
                                List<Object> row = new ArrayList<Object>();
                                row.add(dataBean.getUsername());
                                row.add(dataBean.getAge());
                                row.add(dataBean.getBirthday());
                                return row;
                            }
                        }).end();

        ExcelFile.write(file,new FileOutputStream("./"+file.getName()));

        //do something with the generated file here
    }

}
