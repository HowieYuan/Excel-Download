package com.howie.excelandfile;

import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.FileInputStream;
import java.io.OutputStream;
import java.util.List;

/**
 * Created with IntelliJ IDEA
 *
 * @Author yuanhaoyue swithaoy@gmail.com
 * @Description
 * @Date 2018-03-22
 * @Time 21:24
 */
@RestController
public class Controller {
    private final PersonFactroy personFactroy;

    @Autowired
    public Controller(PersonFactroy personFactroy) {
        this.personFactroy = personFactroy;
    }

    /**
     * 将数据全部导入到一张表格
     */
    @RequestMapping(value = "/getExcel/onlyOne", method = RequestMethod.GET)
    public void createBoxListExcelOnlyOne(HttpServletResponse response) throws Exception {
        //创建文件本地文件
        String filePath = "人员数据.xls";
        File dbfFile = new File(filePath);
        //首先要使用Workbook类的工厂方法创建一个可写入的工作薄(Workbook)对象
        WritableWorkbook wwb = Workbook.createWorkbook(dbfFile);
        if (!dbfFile.exists() || dbfFile.isDirectory()) {
            dbfFile.createNewFile();
        }
        List<Person> list = personFactroy.getPersonList();
        WritableSheet ws = wwb.createSheet("列表 1", 0);  //创建一个可写入的工作表

        //添加excel表头
        ws.addCell(new Label(0, 0, "序号"));
        ws.addCell(new Label(1, 0, "姓名"));
        ws.addCell(new Label(2, 0, "年龄"));
        ws.addCell(new Label(3, 0, "性别"));
        int index = 0;
        for (Person person : list) {
            //将生成的单元格添加到工作表中
            //（这里需要注意的是，在Excel中，第一个参数表示列，第二个表示行）
            ws.addCell(new Label(0, index + 1, String.valueOf(person.getId())));
            ws.addCell(new Label(1, index + 1, person.getName()));
            ws.addCell(new Label(2, index + 1, String.valueOf(person.getAge())));
            ws.addCell(new Label(3, index + 1, person.getGender()));
            index++;
        }
        wwb.write();//从内存中写入文件中
        wwb.close();//关闭资源，释放内存
        String fileName = new String("人员信息.xlsx".getBytes(), "ISO-8859-1");
        response.addHeader("Content-Disposition", "filename=" + fileName);
        OutputStream os = response.getOutputStream();
        FileInputStream fis = new java.io.FileInputStream(filePath);
        byte[] b = new byte[1024];
        int j;
        while ((j = fis.read(b)) > 0) {
            os.write(b, 0, j);
        }
        fis.close();
        os.flush();
        os.close();
    }

    /**
     * 将数据导入到多个表格
     */
    @RequestMapping(value = "/getExcel", method = RequestMethod.GET)
    public void createBoxListExcel(HttpServletResponse response) throws Exception {
        //创建文件本地文件
        String filePath = "人员数据.xls";
        File dbfFile = new File(filePath);
        //首先要使用Workbook类的工厂方法创建一个可写入的工作薄(Workbook)对象
        WritableWorkbook wwb = Workbook.createWorkbook(dbfFile);
        if (!dbfFile.exists() || dbfFile.isDirectory()) {
            dbfFile.createNewFile();
        }
        List<Person> list = personFactroy.getPersonList();
        int mus = 2;//每个工作表格最多存储2条数据（注：excel表格一个工作表可以存储65536条）
        int totle = list.size();//获取List集合的size
        int avg = totle / mus;
        for (int i = 0; i < avg + 1; i++) {
            WritableSheet ws = wwb.createSheet("列表" + (i + 1), i);  //创建一个可写入的工作表

            //添加excel表头
            ws.addCell(new Label(0, 0, "序号"));
            ws.addCell(new Label(1, 0, "姓名"));
            ws.addCell(new Label(2, 0, "年龄"));
            ws.addCell(new Label(3, 0, "性别"));
            int num = i * mus;
            int index = 0;
            for (int m = num; m < list.size(); m++) {
                if (index == mus) {//判断index == mus的时候跳出当前for循环
                    break;
                }
                Person person = list.get(m);

                //将生成的单元格添加到工作表中
                //（这里需要注意的是，在Excel中，第一个参数表示列，第二个表示行）
                ws.addCell(new Label(0, index + 1, String.valueOf(person.getId())));
                ws.addCell(new Label(1, index + 1, person.getName()));
                ws.addCell(new Label(2, index + 1, String.valueOf(person.getAge())));
                ws.addCell(new Label(3, index + 1, person.getGender()));
                index++;
            }
        }
        wwb.write();//从内存中写入文件中
        wwb.close();//关闭资源，释放内存
        String fileName = new String("人员信息.xls".getBytes(), "ISO-8859-1");
        response.addHeader("Content-Disposition", "filename=" + fileName);
        OutputStream outputStream = response.getOutputStream();
        FileInputStream fileInputStream = new FileInputStream(filePath);
        byte[] b = new byte[1024];
        int j;
        while ((j = fileInputStream.read(b)) > 0) {
            outputStream.write(b, 0, j);
        }
        fileInputStream.close();
        outputStream.flush();
        outputStream.close();
    }


}
