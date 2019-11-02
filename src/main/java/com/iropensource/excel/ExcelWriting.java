package com.iropensource.excel;

import com.iropensource.model.User;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.List;

public class ExcelWriting {

    public static void writeExcel(final List<User> users, final String filePath) throws IOException {
        File file = new File(filePath);

        if (users == null || users.isEmpty() || filePath == null)
            throw new IllegalArgumentException("Input Data has some problems");

        if(!file.exists() && !file.createNewFile())
            throw new IllegalArgumentException("Path is wrong");


        //یه ورک بوک برای شما درست میکنه
        Workbook wb = null;
        OutputStream os = null;
        try {
            wb = new HSSFWorkbook();

            // پر کردن محتوای اکسل
            wb = createContent(wb, users);

            // آماده سازی stream برای نوشتن روی فایل
            os = new FileOutputStream(file);

            // نوشتن بر روی فایلا
            wb.write(os);
        } finally {
            // استریم هارو ببند - این کار برای جلوگیری از مموری لیک و
            // اجازه به گاربیج کالکشن برای حذف دیتاهای اضافی باید گذاشته شود
            if (wb != null)
                wb.close();

            if (os != null)
                os.close();
        }


    }

    private static Workbook createContent(final Workbook wb, final List<User> users) {
        if (wb == null)
            throw new IllegalArgumentException("Data is missing");

        // صفحه ی Excel با نام top-sheet
        Sheet mainSheet = wb.createSheet("top-sheet");

        //درست کردن سرفصله صفحه
        String[] heading = {"First Name", "Last Name", "Username", "Email Address"};
        Row headingRow = mainSheet.createRow(0);
        for (int i = 0; i < heading.length; i++) {
            Cell cell = headingRow.createCell(i);
            cell.setCellValue(heading[i]);

            //اضافه کردن استایل برای هدینگ صفحه
            CellStyle cs = wb.createCellStyle();

            // تنظیم فونت و رنگش
            Font font = wb.createFont();
            font.setBold(true);
            font.setColor(HSSFColor.HSSFColorPredefined.DARK_RED.getIndex());

            // وسط چین کردن سر صفحه
            cs.setAlignment(HorizontalAlignment.CENTER);
            cs.setFont(font);

            cell.setCellStyle(cs);
        }

        //درست کردن محتوای فایل Excel
        for (int row = 0; row < users.size(); row++) {
            Row contentRow = mainSheet.createRow((row + 1));//چون سطر اول هدینگ صفحه هستش و باید از سطر بعدی شروع کنیم
            User user = users.get(row);

            //برای راحت حلقه زدن روی ستون ها بهتره به این شکل درش بیاریم ابتدا
            String[] contents = {user.getFirstName(), user.getLastName(), user.getUsername(), user.getEmailAddress()};
            for (int col = 0; col < contents.length; col++) {
                Cell cell = contentRow.createCell(col);
                cell.setCellValue(contents[col]);
            }
        }

        return wb;
    }
}
