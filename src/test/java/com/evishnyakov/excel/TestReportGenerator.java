package com.evishnyakov.excel;

import lombok.AllArgsConstructor;
import lombok.Getter;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.awt.Color;
import java.io.ByteArrayOutputStream;
import java.time.LocalDate;
import java.time.Month;
import java.util.List;

import static java.util.Arrays.asList;

public class TestReportGenerator {
    @Getter
    @AllArgsConstructor
    private static class Employee {

        private String idNumber;
        private String firstName;
        private String lastName;
        private LocalDate dateOfBirth;

    }

    public static class EmployeesReportGenerator {

        private final StyleCell header, cellLeft, cellCenter, cellDate;

        {
            StyleFont headerFont = StyleFont
                    .builder().bold(true).fontHeight(11).fontName(FontName.CALIBRI).build();
            StyleFont cellFont = StyleFont
                    .builder().fontHeight(11).fontName(FontName.CALIBRI).build();
            Color grey = new Color(217, 217, 217);

            header = StyleCell.builder()
                    .font(headerFont)
                    .horizontalAlignment(HorizontalAlignment.LEFT)
                    .color(grey)
                    .build();
            cellLeft = StyleCell.builder()
                    .font(cellFont)
                    .horizontalAlignment(HorizontalAlignment.LEFT)
                    .build();
            cellCenter = StyleCell.builder()
                    .font(cellFont)
                    .horizontalAlignment(HorizontalAlignment.CENTER)
                    .dataFormat(0)
                    .build();
            cellDate = StyleCell.builder()
                    .font(cellFont)
                    .horizontalAlignment(HorizontalAlignment.LEFT)
                    .formatPattern("DD/MM/YYYY")
                    .build();
        }

        public void generate(XSSFWorkbook workbook, List<Employee> employees) {
            XSSFSheet sheet = workbook.createSheet("Employees");
            SheetBuilder sheetBuilder = new SheetBuilder();
            sheetBuilder
                    .columnWidth(0, 256*20)
                    .columnWidth(1, 256*20)
                    .columnWidth(2, 256*20)
                    .columnWidth(3, 256*20)
                    .cell(c -> c.row(0).column(0).value("ID Number").style(header))
                    .cell(c -> c.row(0).column(1).value("First name").style(header))
                    .cell(c -> c.row(0).column(2).value("Last name").style(header))
                    .cell(c -> c.row(0).column(3).value("Date of birth").style(header));
            for(int i = 0; i < employees.size(); i++) {
                Employee employee = employees.get(i);
                int row = i + 1;
                sheetBuilder
                        .cell(c -> c.row(row).column(0).value(employee.getIdNumber()).style(cellCenter))
                        .cell(c -> c.row(row).column(1).value(employee.getFirstName()).style(cellLeft))
                        .cell(c -> c.row(row).column(2).value(employee.getLastName()).style(cellLeft))
                        .cell(c -> c.row(row).column(3).value(employee.getDateOfBirth()).style(cellDate));
            }
            sheetBuilder.build(sheet);
        }
    }

    @Test
    public void generateReport() throws Exception {
        try(XSSFWorkbook workbook = new XSSFWorkbook()) {
            new EmployeesReportGenerator().generate(workbook, asList(
                    new Employee("1X90RE", "Ivan", "Ivanov", LocalDate.of(1985, Month.JUNE, 15)),
                    new Employee("23TY78", "Petr", "Petrov", LocalDate.of(1980, Month.JULY, 20)),
                    new Employee("85WS41", "Vasay", "Sidorov", LocalDate.of(1975, Month.DECEMBER, 25))
            ));

            ByteArrayOutputStream bout = new ByteArrayOutputStream();
            workbook.write(bout);
            //IOUtils.write(bout.toByteArray(), new FileOutputStream("report.xlsx"));
        }
    }
}
