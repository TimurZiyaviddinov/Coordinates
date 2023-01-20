package org.example;


import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;

public class Main {
    public static void main(String[] args) throws IOException {
        String path = "C:\\Users\\Тимур\\Desktop\\test1.txt";
        File file = new File(path);
        String input;
        String[] arr = new String[1000];
        int i = 0;
        BufferedReader bufferedReader = new BufferedReader(new FileReader(file));
        while ((input = bufferedReader.readLine()) != null) {
            if (!input.equals("")) {
                input = input.substring(0,7);
                input = new StringBuilder(input).insert(2,"°").insert(5, ".").toString();
//                input = input + "'";
                arr[i] = input.replace("B", "8")
                        .replace("C","0")
                        .replace("E", "8")
                        .replace("'", "")
                        .replace("*", "")
                        .replace("^", "");
                i++;
            }
        }
//        System.out.println(i);
        String[] arrN = new String[(i) / 2];
        System.arraycopy(arr, 0, arrN, 0, (i) / 2);
        String[] arrE = new String[(i) / 2];
        System.arraycopy(arr, i / 2, arrE, 0, (i) / 2);

//             Joining two coordinates into one line
//        for (i = 0; i < arrE.length; i++) {
//            arrN[i] = arrN[i] + " " + arrE[i];
//        }
//        Arrays.stream(arrN).forEach(System.out::println);
        Main main = new Main();
        main.writeDataToExcel(arrN, arrE, "20");
        System.out.println(i%2==0);
    }

    public void writeDataToExcel(String[] arr1, String [] arr2, String fileName) {

        // workbook object
        XSSFWorkbook workbook = new XSSFWorkbook();

        // spreadsheet object
        XSSFSheet spreadsheet
                = workbook.createSheet("Coordinates");

        // creating a row object
        XSSFRow row;

        // This data needs to be written (Object[])


        // writing the data into the sheets...

        row = spreadsheet.createRow(0);
        Cell cell = row.createCell(1);
        cell.setCellValue("longitude");
        Cell cell0 = row.createCell(0);
        cell0.setCellValue("latitude");
        for (int i = 0; i < arr1.length; i++) {
            row = spreadsheet.createRow(i + 1);
            Cell cell1 = row.createCell(0);
            cell1.setCellValue(arr1[i]);
            Cell cell2 = row.createCell(1);
            cell2.setCellValue(arr2[i]);

        }


        // .xlsx is the format for Excel Sheets...
        // writing the workbook into the file...
        try (
                FileOutputStream out = new FileOutputStream(
                        new File("C:\\Users\\Тимур\\Desktop\\new excel\\" + fileName + ".xlsx"))) {

            workbook.write(out);
            System.out.printf("File '%s' has been created! \n", fileName);

        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }
}