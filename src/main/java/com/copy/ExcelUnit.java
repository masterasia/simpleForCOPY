package com.copy;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.DateUtil;

import java.io.*;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.List;
import java.util.Properties;

public class ExcelUnit {
    public static String numbers = getNumbers("numbers");

    public static String extype = getNumbers("extype");

    public static void main(String[] args) {
        System.out.println(numbers);
        if (extype.equals("1")) {
            CopysExcel c = new CopysExcel();
            c.exAll();
        } else {
            changeExcel();
            // read2003Excel(new File(getNumbers("sourceEXCEL")));
        }
    }

    public static String getNumbers(String key) {
        Properties properties = new Properties();
        try {
            properties.load(new FileInputStream(new File(
                    "d:/resource.properties")));
        } catch (FileNotFoundException e1) {
            System.out.println(" 保存循环条数的文件丢了 ");
            e1.printStackTrace();
        } catch (IOException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
        return (String) properties.get(key);
    }

    public static void changeExcel() {
        List<List<Object>> lists = readExcel(new File(getNumbers("sourceEXCEL")));
        if (writeExcel(lists)) {
            System.out.println(" 操作成功 ");
        } else {
            System.out.println(" 操作失败 ");
        }
    }

    public static boolean writeExcel(List<List<Object>> results) {
        boolean flag = false;
        HSSFRow row = null;
        HSSFCell cell = null;
        HSSFSheet sheet = null;
        FileOutputStream outputStream = null;
        List<Object> rowList = null;
        HSSFWorkbook workbook = new HSSFWorkbook();
        sheet = workbook.createSheet();
        int count = 0;
        for (int i = 0; i < results.size(); i++) {
            if (count == Integer.parseInt(numbers)) {
                sheet = workbook.createSheet();
                count = 0;
            }
            System.err.println(count);
            rowList = results.get(i);
            row = sheet.createRow((short) count);
            for (int j = 0; j < rowList.size(); j++) {
                cell = row.createCell(j);
                cell.setCellValue(rowList.get(j).toString());
            }
            count++;
        }
        try {
            outputStream = new FileOutputStream(getNumbers("targetEXCEL"));
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

    public static List<List<Object>> readExcel(File file) {
        String fileName = file.getName();
        String extension = -1 == fileName.lastIndexOf(".") ? "" : fileName
                .substring(fileName.lastIndexOf(".") + 1);
        if ("xls".equals(extension)) {
            System.out.println(" this is a 2003 excel ");
            return read2003Excel(file);
        } else if ("xlsx".equals(extension)) {
            System.out.println(" this is a 2007 excel ");
            return read2007Excel(file);
        } else {
            System.out.println(" 不是标准的excel格式 ");
            return null;
        }
    }

    public static List<List<Object>> read2003Excel(File file) {
        HSSFRow row = null;
        HSSFCell cell = null;
        Object value = null;
        List<List<Object>> results = new ArrayList<List<Object>>();
        try {
            System.out.println(file.getName());
            FileInputStream fis = new FileInputStream(file);
            HSSFWorkbook workbook = new HSSFWorkbook(fis);
            HSSFSheet sheet = workbook.getSheetAt(0);
            for (int i = sheet.getFirstRowNum(); i <= sheet
                    .getPhysicalNumberOfRows(); i++) {
                row = sheet.getRow(i);
                if (null == row) {
                    continue;
                }
                List<Object> rowList = new ArrayList<Object>();
                for (int j = row.getFirstCellNum(); j <= row.getLastCellNum(); j++) {
                    cell = row.getCell(j);
                    if (null == cell) {
                        continue;
                    }
                    DecimalFormat df = new DecimalFormat("0");// 格式化 number
                    // String 字符
                    SimpleDateFormat sdf = new SimpleDateFormat(
                            "yyyy-MM-dd HH:mm:ss");// 格式化日期字符串
                    DecimalFormat nf = new DecimalFormat("0.00");// 格式化数字
                    switch (cell.getCellType()) {
                        case HSSFCell.CELL_TYPE_STRING:
                            System.out.println(i + "行" + j + " 列 is String type");
                            value = cell.getStringCellValue();
                            break;
                        case HSSFCell.CELL_TYPE_NUMERIC:
                            System.out.println(i + "行" + j
                                    + " 列 is Number type ; DateFormt:"
                                    + cell.getCellStyle().getDataFormatString());
                            if ("@".equals(cell.getCellStyle()
                                    .getDataFormatString())) {
                                value = df.format(cell.getNumericCellValue());
                            } else if ("General".equals(cell.getCellStyle()
                                    .getDataFormatString())) {
                                value = nf.format(cell.getNumericCellValue());
                            } else if (DateUtil.isCellDateFormatted(cell)) {
                                value = sdf.format(cell.getDateCellValue());
                            } else {
                                value = nf.format(cell.getNumericCellValue());
                            }
                            break;
                        case HSSFCell.CELL_TYPE_BOOLEAN:
                            System.out.println(i + "行" + j + " 列 is Boolean type");
                            value = cell.getBooleanCellValue();
                            break;
                        case HSSFCell.CELL_TYPE_BLANK:
                            System.out.println(i + "行" + j + " 列 is Blank type");
                            value = "";
                            break;
                        default:
                            System.out.println(i + "行" + j + " 列 is default type");
                            value = cell.toString();
                    }

                    rowList.add(value);
                }
                results.add(rowList);
            }

        } catch (FileNotFoundException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        } catch (IOException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
        System.out.println(results);
        return results;
    }

    public static List<List<Object>> read2007Excel(File file) {
        return null;
    }
}
