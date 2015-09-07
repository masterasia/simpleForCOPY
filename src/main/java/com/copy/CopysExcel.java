package com.copy;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.*;
import java.util.*;
import java.util.Map.Entry;

public class CopysExcel {

    private static int max = 0;

    public void exAll() {
        List<Papers> map = readTxt();
        writeExcel(map);
    }

    public static void main(String[] args) {
        CopysExcel c = new CopysExcel();
        c.exAll();
    }

    public static List<Papers> readTxt() {
        List<Papers> map = new ArrayList<Papers>();
        File direct = new File("D:\\eat");
        if (null == direct || !direct.isDirectory()) {
            System.out.println();
            return null;
        }
        File[] files = direct.listFiles();
        Map<String, File> maps = new TreeMap<String, File>();
        for (File file : files) {
            maps.put(file.getName(), file);
        }
        Set<Entry<String, File>> set = maps.entrySet();

        for (Entry<String, File> entry : set) {
            Papers p = new Papers();
            p.setName(entry.getKey());
            read(entry.getValue(), p);
            map.add(p);
        }

        return map;
    }

    public static void read(File file, Papers p) {
        try {
            BufferedReader br = new BufferedReader(new FileReader(file));
            String temp = br.readLine();
            List<String> list = p.getKey_value();
            while (temp != null) {
                if (temp.trim().length() < 1 || temp.contains(",") || temp.contains("V")) {
                    temp = br.readLine();
                    continue;
                }
                list.add(temp.trim());
                temp = br.readLine();
            }

            max = max > list.size() ? max : list.size();
        } catch (Exception e) {

        }
    }

    public static boolean writeExcel(List<Papers> papers) {
        boolean flag = false;
        boolean tORt = false;
        HSSFRow row = null;
        HSSFCell cell = null;
        HSSFSheet sheet = null;
        FileOutputStream outputStream = null;
        HSSFWorkbook workbook = new HSSFWorkbook();
        sheet = workbook.createSheet();
        try {
            System.setOut(new PrintStream(new File("D:\\log.txt")));
        } catch (FileNotFoundException e1) {
            // TODO Auto-generated catch block
            e1.printStackTrace();
        }

        for (int i = 0; i < max; i++) {
            row = sheet.createRow((short) i);
            int cellCount = 0;
            for (int j = 0; j < papers.size(); j++) {
                Papers paper = papers.get(j);
                if (0 == i) {
                    cell = row.createCell(cellCount++);
                    cell.setCellValue("");
                    if (tORt){
                        cell = row.createCell(cellCount++);
                        cell.setCellValue("");
                    }else {
                        if (paper.getKey_value().get(0).contains("\t"))
                            tORt = true;
                        else
                            tORt = false;
                    }
                    cell = row.createCell(cellCount++);
                    cell.setCellValue(paper.getName());
                } else {
                    List<String> list = paper.getKey_value();
                    if (i >= list.size()) {
                        if (tORt) {
                            cellCount = cellCount + 3;
                        }
                        else {
                            cellCount = cellCount + 2;
                        }

                        continue;
                    }
                    String[] s;
                    if (list.get(i - 1).contains("\t")) {
                        s = list.get(i - 1).split("\t");
                    }
                    else {
                        s = list.get(i - 1).split(" ");
                    }
                    for (String string : s) {
                        if (null == string || string.isEmpty()
                                || string.trim().length() < 1)
                            continue;
                        if (i > 230 || cellCount > 230) {
                            System.out.println(paper.getName() + " "
                                    + cellCount);
                            System.out.println(i + " " + string);
                        }
                        cell = row.createCell(cellCount++);
                        cell.setCellValue(string);
                    }
                }
            }
        }

        try {
            outputStream = new FileOutputStream(ExcelUnit.getNumbers("sumTxt"));
            workbook.write(outputStream);
            flag = true;
        } catch (FileNotFoundException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        } catch (IOException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        } finally {
            try {
                outputStream.close();
            } catch (IOException e) {
                outputStream = null;
                e.printStackTrace();
            }
        }
        return flag;
    }
}
