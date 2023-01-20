package org.example;


import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;

public class Main {
    public static void main(String[] args) throws IOException {
        String path = "C:\\Users\\ADmin\\Desktop\\test1.txt";
        File file = new File(path);
        String input;
        String[] arr = new String[1000];
        int i = 0;
        BufferedReader bufferedReader = new BufferedReader(new FileReader(file));
        while ((input = bufferedReader.readLine()) != null) {
            if (!input.equals("")) {
                char[] inputChar = input.toCharArray();
                inputChar[2] = ' ';
                inputChar[5] = '.';
                input = String.valueOf(inputChar);
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

//             Joining two coordinates into ine line
        for (i = 0; i < arrE.length; i++) {
            arrN[i] = arrN[i] + " " + arrE[i];
        }
//        Arrays.stream(arrN).forEach(System.out::println);
        Main main = new Main();
        main.writeDataToExcel(arrN, "19Layer");
    }

    public void writeDataToExcel(String[] arr1, String fileName) {

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
        Cell cell = row.createCell(0);
        cell.setCellValue("Coordinates");
        for (int i = 0; i < arr1.length; i++) {
            row = spreadsheet.createRow(i + 1);
            Cell cell1 = row.createCell(0);
            cell1.setCellValue(arr1[i]);

        }


        // .xlsx is the format for Excel Sheets...
        // writing the workbook into the file...
        try (
                FileOutputStream out = new FileOutputStream(
                        new File("C:\\Users\\ADmin\\Desktop\\new excel\\" + fileName + ".xlsx"))) {

            workbook.write(out);
            System.out.printf("File '%s' has been created!", fileName);

        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }
}