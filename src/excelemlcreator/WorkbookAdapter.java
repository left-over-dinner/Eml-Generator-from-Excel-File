package excelemlcreator;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

public class WorkbookAdapter {
    public static Iterator<Row> createRowIterator(String sourcePath){
        try {
            File filePath = new File(sourcePath);
            FileInputStream filePathStream = new FileInputStream(filePath);
            Workbook workbook = WorkbookFactory.create(filePathStream);
            //by default use first sheet
            Sheet sheet = workbook.getSheetAt(0);
            return sheet.iterator();
        }
        catch (FileNotFoundException e){
            System.out.println("File not found at: "+sourcePath);
        }
        catch (IOException e){
            e.printStackTrace();
        }
        return null;
    }
}
