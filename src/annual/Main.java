/*
 * Decompiled with CFR 0_115.
 * 
 * Could not load the following classes:
 *  org.apache.poi.ss.usermodel.Sheet
 *  org.apache.poi.xssf.usermodel.XSSFWorkbook
 */
package annual;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Main {
    static ArrayList<String> sheetNames = new ArrayList();

    public static void getSheets() throws FileNotFoundException, IOException {
        String excelFilePath = "C:\\Users\\talha\\Desktop\\deneme.xlsx";
        FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
        XSSFWorkbook workbook = new XSSFWorkbook((InputStream)inputStream);
        for (Sheet sheet : workbook) {
            sheetNames.add(sheet.getSheetName());
        }
        workbook.close();
        inputStream.close();
    }

    public static void main(String[] args) throws IOException {
    }
}
