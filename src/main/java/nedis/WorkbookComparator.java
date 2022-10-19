/*
 * Copyright (c) 2022
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package nedis;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.Locale;
import java.util.Objects;

public final class WorkbookComparator {

    private final boolean trim;

    private final Workbook workbook1;

    private final Workbook workbook2;

    public WorkbookComparator(boolean trim, Workbook workbook1, Workbook workbook2) {
        this.trim = trim;
        this.workbook1 = workbook1;
        this.workbook2 = workbook2;
    }

    public void compare() {
        assertEquals(workbook1.getNumberOfSheets(), workbook2.getNumberOfSheets(), "NumberOfSheets");
        boolean displayCell = false;
        for (int sheetIndex = 0; sheetIndex < workbook1.getNumberOfSheets(); sheetIndex++) {
            Sheet sheet1 = workbook1.getSheetAt(sheetIndex);
            Sheet sheet2 = workbook2.getSheetAt(sheetIndex);
            for (int rowIndex = 0; ; rowIndex++) {
                Row row1 = sheet1.getRow(rowIndex);
                Row row2 = sheet2.getRow(rowIndex);
                if (row1 == null || row2 == null) {
                    if (row1 != null) {
                        System.err.printf("Expected: row with %s index, actual: no row %n", rowIndex);
                    } else if (row2 != null) {
                        System.err.printf("Expected: no row, actual: row with %s index %n", rowIndex);
                    }
                    System.out.println("Last row: " + (rowIndex));
                    break;
                }
                for (int colIndex = 0; ; colIndex++) {
                    Cell cell1 = row1.getCell(colIndex);
                    Cell cell2 = row2.getCell(colIndex);
                    if (cell1 == null || cell2 == null) {
                        if (!displayCell) {
                            if (cell1 != null) {
                                System.err.printf("Expected: col with %s index, actual: no col %n", getColNum(colIndex));
                            } else if (cell2 != null) {
                                System.err.printf("Expected: no col, actual: col with %s index %n", getColNum(colIndex));
                            }
                            displayCell = true;
                            System.out.println("Last col: " + getColNum(colIndex - 1));
                        }
                        break;
                    }

                    CellType cellType1 = cell1.getCellType();
                    CellType cellType2 = cell2.getCellType();
                    if (cellType1 == CellType.BLANK && cellType2 == CellType.BLANK) {
                        // do nothing
                    } else if (cellType1 == CellType.BLANK && cellType2 == CellType.STRING && cell2.getStringCellValue().isBlank()) {
                        // do nothing
                    } else if (cellType1 == CellType.STRING && cellType2 == CellType.BLANK && cell1.getStringCellValue().isBlank()) {
                        // do nothing
                    } else {
                        String cellPosition = String.format("[Row=%s, Col=%s]", rowIndex + 1, getColNum(colIndex));

                        if (cellType1 == CellType.STRING && cellType2 == CellType.STRING) {
                            assertEquals(toStringValue(cell1), toStringValue(cell2), cellPosition);
                        } else if (cellType1 == CellType.NUMERIC && cellType2 == CellType.NUMERIC) {
                            assertEquals(cell1.getNumericCellValue(), cell2.getNumericCellValue(), cellPosition);
                        } else if (cellType1 == CellType.STRING && cellType2 == CellType.NUMERIC) {
                            assertEquals(cell1.getStringCellValue(), cell2.getNumericCellValue(), cellPosition);
                        } else if (cellType1 == CellType.NUMERIC && cellType2 == CellType.STRING) {
                            assertEquals(cell1.getNumericCellValue(), cell2.getStringCellValue(), cellPosition);
                        } else {
                            assertEquals(cellType1, cellType2, cellPosition);
                            System.err.println("Cell type is: " + cellType1);
                        }
                    }
                }
            }
        }
    }

    private Object toStringValue(Cell cell) {
        String stringCellValue = cell.getStringCellValue();
        return trim ? stringCellValue.trim() : stringCellValue;
    }

    private void assertEquals(Object o1, Object o2, String cellPosition) {
        if (!Objects.equals(o1, o2)) {
            System.err.printf("Expected: `%s` (%s), actual: `%s` (%s) for %s cell%n", o1, typeOf(o1), o2, typeOf(o2), cellPosition);
        }
    }
    
    private String typeOf(Object o) {
        return o == null ? "null" : o.getClass().getName().replace("java.lang.", "").toLowerCase(Locale.ENGLISH);
    }

    private String getColNum(int value) {
        String s = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

        int high = value / s.length();
        int low = value - high * s.length();
        if (high == 0) {
            return String.valueOf(s.charAt(low));
        } else {
            return String.valueOf(s.charAt(high - 1)) + s.charAt(low);
        }
    }

}

