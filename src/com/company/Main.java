package com.company;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;

public class Main {


    public static void main(String[] args) throws IOException {
//        String path = "d:/";
        String path = "./";
//        String fileName = "test";
        String fileName = "城市列表";
        String fileType = "xls";
//        writer(path, fileName, fileType);
        read(path, fileName, fileType);

    }


    private static void writer(String path, String fileName, String fileType) {
        //创建工作文档对象
        Workbook wb = null;
        try {
            if (fileType.equals("xls")) {
                wb = new HSSFWorkbook();

            } else if (fileType.equals("xlsx")) {
//                wb = new XSSFWorkbook();

            } else {
                System.out.println("您的文档格式不正确！");

            }
            //创建sheet对象
            Sheet sheet1 = (Sheet) wb.createSheet("sheet1");
            //循环写入行数据
            for (int i = 0; i < 5; i++) {
                Row row = (Row) sheet1.createRow(i);
                //循环写入列数据
                for (int j = 0; j < 8; j++) {
                    Cell cell = row.createCell(j);
                    cell.setCellValue("测试" + j);

                }

            }
            //创建文件流
            OutputStream stream = new FileOutputStream(path + fileName + "." + fileType);
            //写入数据
            wb.write(stream);
            //关闭文件流
            stream.close();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            if (wb != null) {
                try {
                    wb.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }

    }


    public static void read(String path, String fileName, String fileType)


    {
        Workbook wb = null;
        PrintWriter fileWriter = null;
        try {
            InputStream stream = new FileInputStream(path + fileName + "." + fileType);
            fileWriter = new PrintWriter(new File("all_cities.txt"));
            if (fileType.equals("xls")) {
                wb = new HSSFWorkbook(stream);

            } else if (fileType.equals("xlsx")) {
                wb = new XSSFWorkbook(stream);

            } else {
                System.out.println("您输入的excel格式不正确");

            }
            Sheet sheet1 = wb.getSheetAt(0);
            for (Row row : sheet1) {
                StringBuilder builder = new StringBuilder();
                for (Cell cell : row) {
//                    case 0:
//                        return "numeric";
//                    case 1:
//                        return "text";
//                    case 2:
//                        return "formula";
//                    case 3:
//                        return "blank";
//                    case 4:
//                        return "boolean";
//                    case 5:
//                        return "error";
//                    default:
//                        return "#unknown cell type (" + cellTypeCode + ")#";
//                    if (cell.getCellType() == 0) {
//
//                        System.out.print(cell.getNumericCellValue() + "  ");
//                    }
                    String s = null;
                    if (cell.getColumnIndex() > 5) {
                        continue;
                    }
                    switch (cell.getCellType()) {
                        case 0:
                            s = String.valueOf((int) cell.getNumericCellValue()).trim();
                            break;
//                            return "numeric";
                        case 1:
                            s = cell.getStringCellValue().trim();
                            if (s.equals("√")) {
                                s = "1";
                            }
                            break;
//                            return "text";
                        case 2:
                            s = cell.getCellFormula().trim();
                            break;
//                            return "formula";
                        case 3:
                            s = " ";
                            break;
//                            return "blank";
                        case 4:
                            s = String.valueOf(cell.getBooleanCellValue());
                            break;
//                            return "boolean";
                        case 5:
                            s = String.valueOf(cell.getErrorCellValue());
                            break;
//                            return "error";
                    }
                    if (cell.getColumnIndex() != 5) {
                        s = String.format("%s&", s);
                    }
                    builder.append(s);
//                    System.out.print(s);

                }
                fileWriter.println(builder.toString());
//                builder.append("/n");
//                System.out.println();

            }
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            if (wb != null) {
                try {
                    wb.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
            if (fileWriter != null) {
                fileWriter.close();
            }
        }

    }

}
