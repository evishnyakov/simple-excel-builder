# Simple Excel builder

A set of classes which wrap Apache POI API for providing fluent API for building complex Excel reports.

## Entry point

Class `SheetBuilder` is a main entry point for building excel document. 

It has two categories of methods:
1. Settings for group of cells:
    1. `defaultRowHeightInPoints`
    2. `columnWidth`
    3. `row`
2. Settings for only one cell:
    1. `cell`

## Example

The process of building excel document consist of following steps 

1. **Define styles of your excel document.** The excel caches and reuses styles. 
   That's why defining it only once and referring to them later may significant
   reduce size of your excel document and amount of time for loading it.  
2. **Create new instance of `SheetBuilder`.**
3. **Invoke method `SheetBuilder#build`.**

#### Note:
*Apache POI has a bug! You can't define several cells with different horizontal alignment in one row. 
But there is a workaround of it, you should just set up  data format for such cells, e.g. `StyleCell#dataFormat`
Look at Example where this situation is described.*

##### Full sources of working example can be found in test `TestReportGenerator`

```java
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
        public void generate(Workbook workbook, List<Employee> employees) {
            Sheet sheet = workbook.createSheet("Employees");
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
```

   