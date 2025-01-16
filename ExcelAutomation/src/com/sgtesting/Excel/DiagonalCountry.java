package com.sgtesting.Excel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;

public class DiagonalCountry
{
    public static void main(String[] args) {
        countryDiagonal();
    }
    private static void countryDiagonal()
    {
        FileOutputStream fout=null;
        Workbook wb=null;
        Sheet sh=null;
        Row row=null;
        Cell cell=null;
        try
        {
            wb=new XSSFWorkbook();
            sh=wb.createSheet("Countries");
            for(int i=0;i<20;i++)
            {
                row=sh.createRow(i);
                cell=row.createCell(i);
                cell.setCellValue("Countries"+(i+1));
            }
            fout=new FileOutputStream("G:\\Excel\\Countries.xlsx");
            wb.write(fout);

        } catch (Exception e)
        {
            e.printStackTrace();
        }
        finally
        {
            try {
                    fout.close();
                    wb.close();
            }catch (Exception e)
            {
                e.printStackTrace();
            }
        }
    }
}
