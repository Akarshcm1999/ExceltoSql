package com.example.ExcelToSql.service;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.*;
import java.util.Iterator;
@Service
public class ExcelToUpdateSqlConverter {

    public  void ExcelToSqlFileWriter() {
        String excelFilePath = "C:\\Users\\CM20202812\\Documents\\Excel\\Book 2.xlsx"; // Path to your Excel file
        String sqlFilePath = "C:\\Users\\CM20202812\\Desktop\\springboot\\ExcelToSql\\SQL path\\script.sql"; // Path to your SQL update script file

        try (FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
             Workbook  workbook = new XSSFWorkbook(inputStream);
             FileOutputStream outputStream = new FileOutputStream(sqlFilePath)) {


            Sheet  sheet = workbook.getSheetAt(0); // Assuming you're reading the first sheet

            Iterator<Row> iterator = sheet.iterator();
            while (iterator.hasNext()) {
                Row currentRow = iterator.next();

                // Assuming your primary key (ID column) is in the first column (index 0)
                Cell idCell = currentRow.getCell(0);
                if (idCell == null || idCell.getCellType() != CellType.NUMERIC) {
                    // Skip rows with invalid ID (or handle differently)
                    continue;
                }

                // Constructing SQL UPDATE statement
                StringBuilder sqlStatement = new StringBuilder("UPDATE add_alias SET ");

                Iterator<Cell> cellIterator = currentRow.cellIterator();
                boolean firstCell = true;
                while (cellIterator.hasNext()) {
                    Cell currentCell = cellIterator.next();

                    // Skip ID column in the update statement
                    if (currentCell.getColumnIndex() == 0) {
                        continue;
                    }

                    if (!firstCell) {
                        sqlStatement.append(", ");
                    } else {
                        firstCell = false;
                    }

                   //Take out the cloulumn name such as cmd_id/souce id
                    String columnName = sheet.getRow(0).getCell(currentCell.getColumnIndex()).getStringCellValue(); // Assuming header row
                    sqlStatement.append(columnName).append(" = ");

                    switch (currentCell.getCellType()) {
                        case STRING:
                            sqlStatement.append("'").append(currentCell.getStringCellValue()).append("'");
                            break;
                        case NUMERIC:
                            if (DateUtil.isCellDateFormatted(currentCell)) {
                                sqlStatement.append("'").append(currentCell.getLocalDateTimeCellValue()).append("'");
                            } else {
                                sqlStatement.append(currentCell.getNumericCellValue());
                            }
                            break;
                        case BOOLEAN:
                            sqlStatement.append(currentCell.getBooleanCellValue());
                            break;
                        default:
                            sqlStatement.append("NULL");
                            break;
                    }
                }

                // WHERE clause based on ID
                sqlStatement.append(" WHERE cmd_id = ").append((int) idCell.getNumericCellValue()).append(";\n");

                outputStream.write(sqlStatement.toString().getBytes());
            }

            System.out.println("Update SQL script generated successfully!");

        } catch (IOException | EncryptedDocumentException e) {
            e.printStackTrace();
        }
    }
}
