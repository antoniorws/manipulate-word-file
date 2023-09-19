package com.wordfile;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

public class Main {
    public static void main(String[] args) {

        String filePath = System.getProperty("user.dir") + File.separator + "oldFile.docx";
        String oldName = "oldName";
        String newName = "newName";

        try {
            FileInputStream fis = new FileInputStream(filePath);
            XWPFDocument doc = new XWPFDocument(fis);
            for (XWPFParagraph paragraph : doc.getParagraphs()) {
                for (XWPFRun run : paragraph.getRuns()) {
                    tryUpdateText(oldName, newName, run);
                }
            }
            for (XWPFTable table : doc.getTables()) {
                for (XWPFTableRow row : table.getRows()) {
                    for (XWPFTableCell cell : row.getTableCells()) {
                        for (XWPFParagraph paragraph : cell.getParagraphs()) {
                            for (XWPFRun run : paragraph.getRuns()) {
                                tryUpdateText(oldName, newName, run);
                            }
                        }
                    }
                }
            }
            FileOutputStream fos = new FileOutputStream(
                    System.getProperty("user.dir") + File.separator + "newFile.docx");
            doc.write(fos);

            fis.close();
            fos.close();
            doc.close();

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * The code is trying to update the text in a run XWPFRun.
     * It starts by getting the text at index 0 of the run, which is "Hello World".
     * Then it checks if there is an old name that matches what it's trying to
     * replace with newName.
     * If so, then it replaces oldName with newName and sets the updated text back
     * into index 0 of the run.
     * The code is used to update the text of a run.
     *
     * @param run
     * @param oldName
     * @param newName
     * 
     * @return void [return description]
     */
    private static void tryUpdateText(String oldName, String newName, XWPFRun run) {
        String text = run.getText(0);
        if (text != null && text.contains(oldName)) {
            text = text.replace(oldName, newName);
            run.setText(text, 0);
        }
    }
}