/**
 * Copyright (C) 2006-2014 phloc systems (www.phloc.com)
 * Copyright (C) 2014 Philip Helger (www.helger.com)
 * philip[at]helger[dot]com
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *         http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package com.helger.poi.excel;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.assertNull;
import static org.junit.Assert.assertTrue;

import javax.annotation.Nonnull;

import org.apache.poi.ss.formula.IStabilityClassifier;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Test;

import com.helger.commons.io.resource.ClassPathResource;

/**
 * Test class for class {@link ExcelReadUtils}.
 * 
 * @author Philip Helger
 */
public final class ExcelReadUtilsTest
{
  private static final String TEST1_XLS = "excel/test1.xls";
  private static final String TEST1_XLSX = "excel/test1.xlsx";

  /**
   * Validate reference sheets
   * 
   * @param aWB
   *        Workbook to use
   */
  private void _validateWorkbook (@Nonnull final Workbook aWB)
  {
    final Sheet aSheet1 = aWB.getSheet ("Sheet1");
    assertNotNull (aSheet1);
    assertNotNull (aWB.getSheet ("Sheet2"));
    final Sheet aSheet3 = aWB.getSheet ("Sheet3");
    assertNotNull (aSheet3);
    assertNull (aWB.getSheet ("Sheet4"));

    Cell aCell = aSheet1.getRow (0).getCell (0);
    assertNotNull (aCell);
    assertEquals (Cell.CELL_TYPE_STRING, aCell.getCellType ());
    assertEquals ("A1", aCell.getStringCellValue ());

    aCell = aSheet1.getRow (1).getCell (1);
    assertNotNull (aCell);
    assertEquals (Cell.CELL_TYPE_STRING, aCell.getCellType ());
    assertEquals ("B2", aCell.getStringCellValue ());

    aCell = aSheet1.getRow (2).getCell (2);
    assertNotNull (aCell);
    assertEquals (Cell.CELL_TYPE_STRING, aCell.getCellType ());
    assertEquals ("C\n3", aCell.getStringCellValue ());

    aCell = aSheet1.getRow (3).getCell (3);
    assertNotNull (aCell);
    assertEquals (Cell.CELL_TYPE_NUMERIC, aCell.getCellType ());
    assertEquals (0.00001, 4.4, aCell.getNumericCellValue ());

    for (int i = 0; i < 6; ++i)
    {
      aCell = aSheet3.getRow (i).getCell (i);
      assertNotNull (aCell);
      assertEquals (Cell.CELL_TYPE_NUMERIC, aCell.getCellType ());
      assertEquals (0.00001, i + 1, aCell.getNumericCellValue ());
    }

    // ="abc"
    aCell = aSheet1.getRow (4).getCell (0);
    assertNotNull (aCell);
    assertEquals (Cell.CELL_TYPE_FORMULA, aCell.getCellType ());
    assertEquals ("\"abc\"", aCell.getCellFormula ());
    assertEquals ("abc", aCell.getStringCellValue ());
    CellValue aEvaluated = new ExcelFormulaEvaluator (aWB, IStabilityClassifier.TOTALLY_IMMUTABLE).evaluate (aCell);
    assertEquals (Cell.CELL_TYPE_STRING, aEvaluated.getCellType ());
    assertEquals ("abc", aEvaluated.getStringValue ());

    // =4711
    aCell = aSheet1.getRow (5).getCell (1);
    assertNotNull (aCell);
    assertEquals (Cell.CELL_TYPE_FORMULA, aCell.getCellType ());
    assertEquals ("4711", aCell.getCellFormula ());
    assertEquals (0.00001, 4711, aCell.getNumericCellValue ());
    aEvaluated = new ExcelFormulaEvaluator (aWB, IStabilityClassifier.TOTALLY_IMMUTABLE).evaluate (aCell);
    assertEquals (Cell.CELL_TYPE_NUMERIC, aEvaluated.getCellType ());
    assertEquals (0.00001, 4711, aEvaluated.getNumberValue ());

    // =TRUE
    aCell = aSheet1.getRow (6).getCell (2);
    assertNotNull (aCell);
    assertEquals (Cell.CELL_TYPE_FORMULA, aCell.getCellType ());
    assertEquals ("TRUE", aCell.getCellFormula ());
    assertTrue (aCell.getBooleanCellValue ());
    aEvaluated = new ExcelFormulaEvaluator (aWB, IStabilityClassifier.TOTALLY_IMMUTABLE).evaluate (aCell);
    assertEquals (Cell.CELL_TYPE_BOOLEAN, aEvaluated.getCellType ());
    assertTrue (aEvaluated.getBooleanValue ());

    // Refers to cell at 6/2
    aCell = aSheet1.getRow (7).getCell (3);
    assertNotNull (aCell);
    assertEquals (Cell.CELL_TYPE_FORMULA, aCell.getCellType ());
    assertEquals ("C7", aCell.getCellFormula ());
    assertTrue (aCell.getBooleanCellValue ());
    aEvaluated = new ExcelFormulaEvaluator (aWB, IStabilityClassifier.TOTALLY_IMMUTABLE).evaluate (aCell);
    assertEquals (Cell.CELL_TYPE_BOOLEAN, aEvaluated.getCellType ());
    assertTrue (aEvaluated.getBooleanValue ());
  }

  @Test
  public void testReadWorkbookFromInputStream ()
  {
    // XLS
    Workbook aWB = ExcelReadUtils.readWorkbookFromInputStream (new ClassPathResource (TEST1_XLS));
    assertNotNull (aWB);
    _validateWorkbook (aWB);

    // XLSX
    aWB = ExcelReadUtils.readWorkbookFromInputStream (new ClassPathResource (TEST1_XLSX));
    assertNotNull (aWB);
    _validateWorkbook (aWB);

    // No such file
    aWB = ExcelReadUtils.readWorkbookFromInputStream (new ClassPathResource ("no-such-file.txt"));
    assertNull (aWB);

    // No Excel file
    aWB = ExcelReadUtils.readWorkbookFromInputStream (new ClassPathResource ("excel/nonexcel.txt"));
    assertNull (aWB);
  }

  @Test
  public void testGetCellValueObject ()
  {
    for (final EExcelVersion eVersion : EExcelVersion.values ())
    {
      final Workbook aWB = eVersion.createWorkbook ();
      final Sheet aSheet = aWB.createSheet ();
      final Row aRow = aSheet.createRow (0);
      final Cell aCell = aRow.createCell (0);

      // boolean
      aCell.setCellValue (true);
      assertEquals (Boolean.TRUE, ExcelReadUtils.getCellValueObject (aCell));

      // int
      aCell.setCellValue (4711);
      assertEquals (Integer.valueOf (4711), ExcelReadUtils.getCellValueObject (aCell));

      // long
      aCell.setCellValue (Long.MAX_VALUE);
      assertEquals (Long.valueOf (Long.MAX_VALUE), ExcelReadUtils.getCellValueObject (aCell));

      // double
      aCell.setCellValue (3.14159);
      assertEquals (Double.valueOf (3.14159), ExcelReadUtils.getCellValueObject (aCell));

      // String
      aCell.setCellValue ("Anyhow");
      assertEquals ("Anyhow", ExcelReadUtils.getCellValueObject (aCell));

      // Rich text string
      final Font aFont = aWB.createFont ();
      aFont.setItalic (true);
      final RichTextString aRTS = eVersion.createRichText ("Anyhow");
      aRTS.applyFont (1, 3, aFont);
      aCell.setCellValue (aRTS);
      assertEquals ("Anyhow", ExcelReadUtils.getCellValueObject (aCell));
    }
  }

  @Test
  public void testReadXLS ()
  {
    // XLS
    Workbook aWB = EExcelVersion.XLS.readWorkbook (ClassPathResource.getInputStream (TEST1_XLS));
    assertNotNull (aWB);
    _validateWorkbook (aWB);

    // XLSX
    aWB = EExcelVersion.XLSX.readWorkbook (ClassPathResource.getInputStream (TEST1_XLSX));
    assertNotNull (aWB);
    _validateWorkbook (aWB);
  }
}
