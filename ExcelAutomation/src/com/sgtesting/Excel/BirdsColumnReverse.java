package com.sgtesting.Excel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;

public class BirdsColumnReverse
{
    public static void main(String[] args)
    {
        reverseBirds();
    }
    private static void reverseBirds()
    {
        FileInputStream fin=null;
        FileOutputStream fout=null;
        Workbook wb=null;
        Sheet sh1=null;
        Sheet sh2=null;
        Row rowsh1=null;
        Row rowsh2=null;
        Cell cellsh1=null;
        Cell cellsh2=null;
        try
        {
            fin=new FileInputStream("G:\\Excel\\Sheet11.xlsx");
            wb=new XSSFWorkbook(fin);
            sh1=wb.getSheet("Sheet1");
            sh2=wb.getSheet("Sheet2");
            if(sh2==null)
            {
                sh2=wb.createSheet("Sheet2");
            }
            int rc=sh1.getPhysicalNumberOfRows();
            for(int i=0;i<rc;i++)
            {
                rowsh1=sh1.getRow(i);
                rowsh2=sh2.getRow(rc-1-i);
                if(rowsh2==null)
                {
                    rowsh2=sh2.createRow(rc-1-i);
                }
                    cellsh1=rowsh1.getCell(4);
                    String data=cellsh1.getStringCellValue();
                    System.out.printf("%-15s",data);
                    System.out.println();

                    cellsh2=rowsh2.getCell(4);
                    if(cellsh2==null)
                    {
                        cellsh2=rowsh2.createCell(4);
                    }
                    cellsh2.setCellValue(data);
                }

            fout=new FileOutputStream("G:\\Excel\\Sheet11.xlsx");
            wb.write(fout);
        }catch(Exception e)
        {
            e.printStackTrace();
        }
        finally
        {
            try
            {
                fout.close();
                wb.close();
                fin.close();
            } catch(Exception e)
            {
                e.printStackTrace();
            }
        }
    }
}
