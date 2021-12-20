package com.fz.test;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.FileOutputStream;
import java.sql.*;

public class Test {

    /**
     * 将数据库dname中的表名为tname的数据表写入Excel表格
     * @param dname
     * @param tname
     */

    public static String writeDbtoExcel(String dname,String tname){
        String path="D:/axls/"+tname+".xls";

        HSSFWorkbook book=new HSSFWorkbook();
        HSSFSheet sheet=book.createSheet("表");
        try {
            Class.forName("com.mysql.cj.jdbc.Driver");//根据mysql版本（我的8.0；  mysql5版本为：com.mysql.cj.jdbc.Driver）
            Connection con;
            //mysql5 的url为：jdbc:mysql://localhost:3306/"+dname+"?useUnicode=true&characterEncoding=utf-8","root","12345"
            con = DriverManager.getConnection("jdbc:mysql://localhost:3306/"+dname+"?useUnicode=true&characterEncoding=utf-8&useSSL=false&serverTimezone=GMT","root","123456");
            Statement st=con.createStatement();
            String sql="select * from "+tname;
            ResultSet rs=st.executeQuery(sql);
            System.out.println("rs:"+rs);
            ResultSetMetaData rsmd=rs.getMetaData();//得到结果集的字段名
            System.out.println("结果集的字段名:"+rsmd);
            int c=rsmd.getColumnCount();//得到数据表的结果集的字段的数量
            //生成表单的第一行，即表头
            HSSFRow row0=sheet.createRow(0);//创建第一行
            for(int i=0;i<c;i++){
                HSSFCell cel=row0.createCell(i);//创建第一行的第i列
                cel.setCellValue(rsmd.getColumnName(i+1));
//				cel.setCellStyle(style);
            }
            //将数据表中的数据按行导入进Excel表中
            int r=1;
            while(rs.next()){
                HSSFRow row=sheet.createRow(r++);//创建非第一行的其他行
                for(int i=0;i<c;i++){//仍然是c列，导入第r行的第i列
                    HSSFCell cel=row.createCell(i);
                    //以下两种写法均可
//					cel.setCellValue(rs.getString(rsmd.getColumnName(i+1)));
                    cel.setCellValue(rs.getString(i+1));
                }
            }
            //用文件输出流类创建名为table的Excel表格
            FileOutputStream out=new FileOutputStream(path);
            book.write(out);//将HSSFWorkBook中的表写入输出流中
            book.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
        return path;
    }


    public static void main(String[] args) {

        String dname = "test";
        String tname = "student";
        Test test1 = new Test();
        test1.writeDbtoExcel(dname,tname);
    }

}
