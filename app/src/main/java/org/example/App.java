package org.example;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class App {

  public static List<String> searchExcelFile(String filePath, String searchString) {
    List<String> cellAddresses = new ArrayList<>();
    File targetFile = new File(filePath);

    try (FileInputStream fis = new FileInputStream(targetFile);
        Workbook workbook = WorkbookFactory.create(fis)) {

      for (Sheet sheet : workbook) {
        for (Row row : sheet) {
          for (Cell cell : row) {
            String cellValue = getCellValueAsString(cell);
            if (cellValue != null && cellValue.contains(searchString)) {
              String cellAddress = formatCellAddress(targetFile, sheet, cell, cellValue);
              cellAddresses.add(cellAddress);
            }
          }
        }
      }

    } catch (IOException e) {
      e.printStackTrace();
    }

    return cellAddresses;
  }

  private static String getCellValueAsString(Cell cell) {
    switch (cell.getCellType()) {
      case FORMULA:
        return cell.getCellFormula();
      case STRING:
        return cell.getStringCellValue();
      case NUMERIC:
        return String.valueOf(cell.getNumericCellValue());
      case BOOLEAN:
        return String.valueOf(cell.getBooleanCellValue());
      default:
        return null;
    }
  }

  private static String formatCellAddress(
      File targetFile, Sheet sheet, Cell cell, String cellValue) {
    return String.format(
        "%s : %s!%s : %s",
        targetFile.getAbsolutePath(),
        sheet.getSheetName(),
        cell.getAddress().formatAsString(),
        cellValue);
  }

  public static List<String> searchInDirectory(String directoryPath, String searchString) {
    List<String> results = new ArrayList<>();

    try {
      Files.walk(Paths.get(directoryPath))
          .filter(Files::isRegularFile)
          .filter(
              path ->
                  Files.isRegularFile(path)
                      && Files.isReadable(path)
                      && path.toString().matches(".*\\.(xlsx|xls|xlsm)$")
                      && !path.toString().contains("~$"))
          .forEach(
              path -> {
                List<String> cellAddresses = searchExcelFile(path.toString(), searchString);
                if (!cellAddresses.isEmpty()) {
                  results.addAll(cellAddresses);
                }
              });
    } catch (IOException e) {
      e.printStackTrace();
    }

    return results;
  }

  public static void main(String[] args) {
    if (args.length != 2) {
      System.out.println("Usage: java ExcelSearcher <directory-path> <search-string>");
      return;
    }

    String directoryPath = args[0];
    String searchString = args[1];

    List<String> results = searchInDirectory(directoryPath, searchString);
    if (results.isEmpty()) {
      System.out.println("No cells found containing the specified string.");
    } else {
      results.stream().sorted().forEach(System.out::println);
    }
  }
}
