package org.example;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.example.model.Employee;

public class Main {

  public static void main(String[] args) {
    // Writing some data to Excel file
    // Prepare data
    Employee employee1 = new Employee(1, "Donald", "Duck");
    Employee employee2 = new Employee(2, "Mickey", "Mouse");
    Employee employee3 = new Employee(3, "Minnie", "Mouse");

    Map<Integer, List<Object>> data = new TreeMap<>();
    data.put(0, List.of("ID", "NAME", "LASTNAME"));
    data.put(1, List.of(employee1.getId(), employee1.getName(), employee1.getLastName()));
    data.put(2, List.of(employee2.getId(), employee2.getName(), employee2.getLastName()));
    data.put(3, List.of(employee3.getId(), employee3.getName(), employee3.getLastName()));

    // Creating a workbook
    try (Workbook workbook = new XSSFWorkbook()) {
      // Create a sheet
      Sheet sheet = workbook.createSheet("Employee Data");
      // Iterate through the data and write it to sheet
      Set<Integer> keyset = data.keySet();
      for (Integer key : keyset) {
        Row row = sheet.createRow(key);
        List<Object> rowValue = data.get(key);

        for (int cellnum = 0; cellnum < rowValue.size(); cellnum++) {
          Object object = rowValue.get(cellnum);
          Cell cell = row.createCell(cellnum);
          if (object instanceof String) {
            cell.setCellValue((String) object);
          } else if (object instanceof Integer) {
            cell.setCellValue((Integer) object);
          }
        }
      }

      // write a workbook to the file
      try (FileOutputStream outputStream = new FileOutputStream("src/main/resources/emp_data.xlsx")) {
        workbook.write(outputStream);
        System.out.println("Write to the file successful");
      }
    } catch (IOException e) {
      System.out.println("Problem closing a workbook: " + e.getMessage());
    }
  }
}