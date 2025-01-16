package com.sgtesting.Excel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;

public class TwentyFlowers
{
    public static void main(String[] args) {
        flowersExcel();
    }
    private static void flowersExcel()
    {
        FileOutputStream fout = null;
        Workbook wb = null;
        Sheet sh = null;
        Row row=null;
        Cell cell=null;

        try
        {
            wb = new XSSFWorkbook();
            sh = wb.createSheet("Flowers");

            for (int i = 0; i < 20; i++)
            {
                 row = sh.createRow(i);
                 cell = row.createCell(0);
                cell.setCellValue("Flower"+(i+1));
            }

            fout = new FileOutputStream("G:\\Excel\\Assignment.xlsx");
            wb.write(fout);

        } catch (Exception e) {
            e.printStackTrace();
        }

        finally {
            try
            {
                fout.close();
                wb.close();
            }catch(Exception e)
            {
                e.printStackTrace();
            }
        }
    }
}
