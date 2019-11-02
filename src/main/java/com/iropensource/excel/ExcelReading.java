package com.iropensource.excel;

import com.iropensource.model.User;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

public class ExcelReading {


    public static List<User> readExcel(final String filePath) throws IOException, InvalidFormatException {
        File file;
        if (filePath == null || !(file = new File(filePath)).exists() || !file.isFile())
            throw new IllegalArgumentException("Input Data has some problems");

        List<User> users = new ArrayList<>();

        Workbook wb = null;
        InputStream is = null;
        try {
            is = new FileInputStream(file);
            wb = new HSSFWorkbook(is);

            // صفحه اکسل رو بگیر
            Sheet sheet = wb.getSheet("top-sheet");

            // برای استفاده حلقه یه iterator بگیر
            Iterator<Row> iterator = sheet.iterator();

            //چون ردیف اول مرتبطه با هدینگ دیتا و ما میخواهیم آبجکت یوزر رو پر کنیم از این ردیف صرف نظر میکنیم
            // این کد مکان نمای حلقه رو یدونه میندازه جلو یعنی سطر اول هدینگ توی حلقه ی پایین دیگه نیست
            Row currentRow = iterator.next();

            // مادامی که سطری وجود دارد تکرار شو
            while (iterator.hasNext()) {
                // سطر جاری
                currentRow = iterator.next();

                // برای استفاده حلقه روی ستون ها
                Iterator<Cell> cellIterator = currentRow.iterator();


                String firstName = null, lastName = null, username = null, emailAddress = null;
                // مادامی که ستونی وجود دارد تکرار شو
                while (cellIterator.hasNext()) {

                    //ستونه جاری
                    Cell currentCell = cellIterator.next();

                    //مقدار ستون جاری
                    String value = currentCell.getStringCellValue();

                    // طبق ایندکس تون میفهمیم که هرکدام برای چه مقداریه
                    switch (currentCell.getColumnIndex()) {
                        case 0:
                            firstName = value;
                            break;
                        case 1:
                            lastName = value;
                            break;
                        case 2:
                            username = value;
                            break;
                        case 3:
                            emailAddress = value;
                            break;
                    }
                }

                // آبجکت ستون مورد نظر را می سازیم
                User rowUser = new User(firstName, lastName, username, emailAddress);

                //آبجکت رو اضافه میکنیم
                users.add(rowUser);

            }


        } finally {
            // استریم هارو ببند - این کار برای جلوگیری از مموری لیک و
            // اجازه به گاربیج کالکشن برای حذف دیتاهای اضافی باید گذاشته شود
            if (wb != null)
                wb.close();

            if (is != null)
                is.close();
        }


        return users;
    }
}
