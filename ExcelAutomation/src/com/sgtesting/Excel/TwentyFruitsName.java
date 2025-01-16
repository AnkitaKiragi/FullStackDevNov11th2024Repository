package com.sgtesting.Excel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;

public class TwentyFruitsName
{
    public static void main(String[] args) {
        Fruits20();
    }
    private static void Fruits20()
    {
        FileOutputStream fout=null;
        Workbook wb=null;
        Sheet sh=null;
        Row row=null;
        Cell cell=null;

        try
        {
            wb=new XSSFWorkbook();
            sh=wb.createSheet("Fruits");
            row=sh.createRow(0);
            for(int i=0;i<20;i++)
            {
                cell=row.createCell(i);
                cell.setCellValue("Fruits"+(i+1));
            }

            fout = new FileOutputStream("G:\\Excel\\Fruit.xlsx");
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
            }
            catch(Exception e)
            {
                e.printStackTrace();
            }
        }
    }
}
