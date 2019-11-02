package com.iropensource.excel;

import com.iropensource.model.User;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.UUID;

public class TestExcel {

    private static final String FILE_NAME = "iropensource";
    private static final String DESKTOP_PATH = String.format("%s%s%s%s", System.getProperty("user.home"), File.separator, "Desktop", File.separator);
    // آدرس فایل اکسل رو صفحه ی دسکتاپ
    public static final String FILE_PATH = DESKTOP_PATH + FILE_NAME + ".xlsx";

    public static void main(String[] args) throws IOException, InvalidFormatException {
        // هزار تا یوزر فیک بساز
        List<User> users = dummyUsers(1000);
        // همرو ذخیره کن تو فایل اکسل
        ExcelWriting.writeExcel(users, FILE_PATH);

        //فایل اکسل ذخیره شده رو بخون و محتواش رو ذخیره کن توی لیستی از یوزر
        List<User> readUsers = ExcelReading.readExcel(FILE_PATH);

        if (readUsers == null)
            return;

        //نمایش ۱۰ تای اول اگر لیست بزرگتر از ۱۰ تا بود
        int length = readUsers.size() <= 10 ? readUsers.size() : 10;
        for (int i = 0; i < length; i++) {
            System.out.println(readUsers.get(i));
        }
    }

    private static List<User> dummyUsers(final int num) {
        final List<User> users = new ArrayList<>();
        for (int i = 0; i < num; i++) {
            users.add(new User(randStr(5), randStr(8), randStr(6), String.format("%s@%s.%s", randStr(5), randStr(5), randStr(3))));
        }

        return users;
    }

    private static String randStr(int length) {
        return UUID.randomUUID().toString().substring(0, length);
    }
}
