/**
 * Copyright (C) 2014-2017 Philip Helger (www.helger.com)
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

import java.io.File;
import java.io.IOException;
import java.io.OutputStream;
import java.math.BigDecimal;
import java.math.BigInteger;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.ZonedDateTime;
import java.util.Calendar;
import java.util.Date;
import java.util.GregorianCalendar;

import javax.annotation.Nonnegative;
import javax.annotation.Nonnull;
import javax.annotation.Nullable;
import javax.annotation.WillClose;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.helger.commons.ValueEnforcer;
import com.helger.commons.datetime.PDTFactory;
import com.helger.commons.io.file.FileHelper;
import com.helger.commons.io.stream.StreamHelper;
import com.helger.commons.state.ESuccess;
import com.helger.poi.excel.style.ExcelStyle;
import com.helger.poi.excel.style.ExcelStyleCache;

/**
 * A utility class for creating very simple Excel workbooks.
 *
 * @author Philip Helger
 */
public final class WorkbookCreationHelper
{
  /** The BigInteger for the largest possible long value */
  private static final BigInteger BIGINT_MAX_LONG = BigInteger.valueOf (Long.MAX_VALUE);

  /** The BigInteger for the smallest possible long value */
  private static final BigInteger BIGINT_MIN_LONG = BigInteger.valueOf (Long.MIN_VALUE);

  private static final Logger s_aLogger = LoggerFactory.getLogger (WorkbookCreationHelper.class);

  private final Workbook m_aWB;
  private final CreationHelper m_aCreationHelper;
  private final ExcelStyleCache m_aStyleCache = new ExcelStyleCache ();
  private Sheet m_aLastSheet;
  private int m_nLastSheetRowIndex = 0;
  private Row m_aLastRow;
  private int m_nLastRowCellIndex = 0;
  private Cell m_aLastCell;
  private int m_nMaxCellIndex = 0;
  private int m_nCreatedCellStyles = 0;

  public WorkbookCreationHelper (@Nonnull final EExcelVersion eVersion)
  {
    this (eVersion.createWorkbook ());
  }

  public WorkbookCreationHelper (@Nonnull final Workbook aWB)
  {
    m_aWB = ValueEnforcer.notNull (aWB, "Workbook");
    m_aCreationHelper = aWB.getCreationHelper ();
  }

  @Nonnull
  public Workbook getWorkbook ()
  {
    return m_aWB;
  }

  /**
   * Create a new font in the passed workbook.
   *
   * @return The created font.
   */
  @Nonnull
  public Font createFont ()
  {
    return m_aWB.createFont ();
  }

  /**
   * @return A new sheet with a default name
   */
  @Nonnull
  public Sheet createNewSheet ()
  {
    return createNewSheet (null);
  }

  /**
   * Create a new sheet with an optional name
   *
   * @param sName
   *        The name to be used. May be <code>null</code>.
   * @return The created workbook sheet
   */
  @Nonnull
  public Sheet createNewSheet (@Nullable final String sName)
  {
    m_aLastSheet = sName == null ? m_aWB.createSheet () : m_aWB.createSheet (sName);
    m_nLastSheetRowIndex = 0;
    m_aLastRow = null;
    m_nLastRowCellIndex = 0;
    m_aLastCell = null;
    m_nMaxCellIndex = 0;
    return m_aLastSheet;
  }

  /**
   * @return A new row in the current sheet.
   */
  @Nonnull
  public Row addRow ()
  {
    if (m_aLastSheet == null)
      throw new IllegalStateException ("A sheet needs to be created before a row can be added! Call createNewSheet");
    m_aLastRow = m_aLastSheet.createRow (m_nLastSheetRowIndex++);
    m_nLastRowCellIndex = 0;
    m_aLastCell = null;
    return m_aLastRow;
  }

  /**
   * @return The number of rows in the current sheet, 0-based.
   */
  @Nonnegative
  protected int getRowIndex ()
  {
    return m_nLastSheetRowIndex - 1;
  }

  /**
   * @return The number of rows in the current sheet, 0-based.
   */
  @Nonnegative
  public int getRowCount ()
  {
    return m_nLastSheetRowIndex;
  }

  /**
   * @return A new cell in the current row of the current sheet
   */
  @Nonnull
  public Cell addCell ()
  {
    if (m_aLastRow == null)
      throw new IllegalStateException ("A row needs to be created before a cell can be added! Call addRow");
    m_aLastCell = m_aLastRow.createCell (m_nLastRowCellIndex++);

    // Check for the maximum cell index in this sheet
    if (m_nLastRowCellIndex > m_nMaxCellIndex)
      m_nMaxCellIndex = m_nLastRowCellIndex;
    return m_aLastCell;
  }

  /**
   * @param bValue
   *        The value to be set.
   * @return A new cell in the current row of the current sheet with the passed
   *         value
   */
  @Nonnull
  public Cell addCell (final boolean bValue)
  {
    final Cell aCell = addCell ();
    aCell.setCellType (CellType.BOOLEAN);
    aCell.setCellValue (bValue);
    return aCell;
  }

  /**
   * @param aValue
   *        The value to be set. May be <code>null</code>.
   * @return A new cell in the current row of the current sheet with the passed
   *         value
   */
  @Nonnull
  public Cell addCell (@Nullable final Calendar aValue)
  {
    final Cell aCell = addCell ();
    aCell.setCellType (CellType.NUMERIC);
    if (aValue != null)
      aCell.setCellValue (aValue);
    return aCell;
  }

  /**
   * @param aValue
   *        The value to be set. May be <code>null</code>.
   * @return A new cell in the current row of the current sheet with the passed
   *         value
   */
  @Nonnull
  public Cell addCell (@Nullable final Date aValue)
  {
    final Cell aCell = addCell ();
    aCell.setCellType (CellType.NUMERIC);
    if (aValue != null)
      aCell.setCellValue (aValue);
    return aCell;
  }

  /**
   * @param aValue
   *        The value to be set.
   * @return A new cell in the current row of the current sheet with the passed
   *         value
   */
  @Nonnull
  public Cell addCell (@Nullable final LocalDate aValue)
  {
    if (aValue == null)
      return addCell ();
    return addCell (PDTFactory.createZonedDateTime (aValue));
  }

  /**
   * @param aValue
   *        The value to be set. May be <code>null</code>.
   * @return A new cell in the current row of the current sheet with the passed
   *         value
   */
  @Nonnull
  public Cell addCell (@Nullable final LocalDateTime aValue)
  {
    if (aValue == null)
      return addCell ();
    return addCell (PDTFactory.createZonedDateTime (aValue));
  }

  /**
   * @param aValue
   *        The value to be set. May be <code>null</code>.
   * @return A new cell in the current row of the current sheet with the passed
   *         value
   */
  @Nonnull
  public Cell addCell (@Nullable final ZonedDateTime aValue)
  {
    if (aValue == null)
      return addCell ();
    return addCell (GregorianCalendar.from (aValue));
  }

  /**
   * @param aValue
   *        The value to be set. May be <code>null</code>.
   * @return A new cell in the current row of the current sheet with the passed
   *         value
   */
  @Nonnull
  public Cell addCell (@Nullable final BigInteger aValue)
  {
    if (aValue == null)
      return addCell ();

    if (aValue.compareTo (BIGINT_MIN_LONG) >= 0 && aValue.compareTo (BIGINT_MAX_LONG) <= 0)
      return addCell (aValue.longValue ());

    // Too large - add as string
    return addCell (aValue.toString ());
  }

  /**
   * @param dValue
   *        The value to be set.
   * @return A new cell in the current row of the current sheet with the passed
   *         value
   */
  @Nonnull
  public Cell addCell (final double dValue)
  {
    final Cell aCell = addCell ();
    aCell.setCellType (CellType.NUMERIC);
    aCell.setCellValue (dValue);
    return aCell;
  }

  /**
   * @param aValue
   *        The value to be set. May be <code>null</code>.
   * @return A new cell in the current row of the current sheet with the passed
   *         value
   */
  @Nonnull
  public Cell addCell (@Nullable final BigDecimal aValue)
  {
    if (aValue == null)
      return addCell ();

    try
    {
      return addCell (aValue.doubleValue ());
    }
    catch (final NumberFormatException ex)
    {
      // Add as string if too large for a double
      return addCell (aValue.toString ());
    }
  }

  /**
   * @param aValue
   *        The value to be set. May be <code>null</code>.
   * @return A new cell in the current row of the current sheet with the passed
   *         value
   */
  @Nonnull
  public Cell addCell (@Nullable final RichTextString aValue)
  {
    final Cell aCell = addCell ();
    aCell.setCellType (CellType.STRING);
    if (aValue != null)
      aCell.setCellValue (aValue);
    return aCell;
  }

  /**
   * @param sValue
   *        The value to be set. May be <code>null</code>.
   * @return A new cell in the current row of the current sheet with the passed
   *         value
   */
  @Nonnull
  public Cell addCell (@Nullable final String sValue)
  {
    final Cell aCell = addCell ();
    aCell.setCellType (CellType.STRING);
    if (sValue != null)
      aCell.setCellValue (sValue);
    return aCell;
  }

  /**
   * @param sFormula
   *        The formula to be set. May be <code>null</code> to set no formula.
   * @return A new cell in the current row of the current sheet with the passed
   *         formula
   */
  @Nonnull
  public Cell addCellFormula (@Nullable final String sFormula)
  {
    final Cell aCell = addCell ();
    aCell.setCellType (CellType.FORMULA);
    aCell.setCellFormula (sFormula);
    return aCell;
  }

  /**
   * Add a merge region in the current row. Note: only the content of the first
   * cell is used as the content of the merged cell!
   *
   * @param nFirstCol
   *        First column to be merged (inclusive). 0-based
   * @param nLastCol
   *        Last column to be merged (inclusive). 0-based, must be larger than
   *        {@code nFirstCol}
   * @return index of this region
   */
  public int addMergeRegionInCurrentRow (@Nonnegative final int nFirstCol, @Nonnegative final int nLastCol)
  {
    final int nCurrentRowIndex = getRowIndex ();
    return addMergeRegion (nCurrentRowIndex, nCurrentRowIndex, nFirstCol, nLastCol);
  }

  /**
   * Adds a merged region of cells (hence those cells form one)
   *
   * @param nFirstRow
   *        Index of first row
   * @param nLastRow
   *        Index of last row (inclusive), must be equal to or larger than
   *        {@code nFirstRow}
   * @param nFirstCol
   *        Index of first column
   * @param nLastCol
   *        Index of last column (inclusive), must be equal to or larger than
   *        {@code nFirstCol}
   * @return index of this region
   */
  public int addMergeRegion (@Nonnegative final int nFirstRow,
                             @Nonnegative final int nLastRow,
                             @Nonnegative final int nFirstCol,
                             @Nonnegative final int nLastCol)
  {
    return m_aLastSheet.addMergedRegion (new CellRangeAddress (nFirstRow, nLastRow, nFirstCol, nLastCol));
  }

  /**
   * Set the cell style of the last added cell
   *
   * @param aExcelStyle
   *        The style to be set.
   */
  public void addCellStyle (@Nonnull final ExcelStyle aExcelStyle)
  {
    ValueEnforcer.notNull (aExcelStyle, "ExcelStyle");
    if (m_aLastCell == null)
      throw new IllegalStateException ("No cell present for current row!");

    CellStyle aCellStyle = m_aStyleCache.getCellStyle (aExcelStyle);
    if (aCellStyle == null)
    {
      aCellStyle = m_aWB.createCellStyle ();
      aExcelStyle.fillCellStyle (m_aWB, aCellStyle, m_aCreationHelper);
      m_aStyleCache.addCellStyle (aExcelStyle, aCellStyle);
      m_nCreatedCellStyles++;
    }
    m_aLastCell.setCellStyle (aCellStyle);
  }

  /**
   * @return The number of unique styles in the current workbook. Always &ge; 0.
   * @since 5.0.0
   */
  @Nonnegative
  public int getCreatedCellStyleCount ()
  {
    return m_nCreatedCellStyles;
  }

  /**
   * @return The number of cells in the current row in the current sheet,
   *         0-based
   */
  @Nonnegative
  public int getCellCountInRow ()
  {
    return m_nLastRowCellIndex;
  }

  /**
   * @return The maximum number of cells in a single row in the current sheet,
   *         0-based.
   */
  @Nonnegative
  public int getMaximumCellCountInRowInSheet ()
  {
    return m_nMaxCellIndex;
  }

  /**
   * Auto size all columns to be matching width in the current sheet
   */
  public void autoSizeAllColumns ()
  {
    // auto-adjust all columns (except description and image description)
    for (short nCol = 0; nCol < m_nMaxCellIndex; ++nCol)
      try
      {
        m_aLastSheet.autoSizeColumn (nCol);
      }
      catch (final IllegalArgumentException ex)
      {
        // Happens if a column is too large
        s_aLogger.warn ("Failed to resize column " + nCol + ": column too wide!");
      }
  }

  /**
   * Add an auto filter on the first row on all columns in the current sheet.
   */
  public void autoFilterAllColumns ()
  {
    autoFilterAllColumns (0);
  }

  /**
   * @param nRowIndex
   *        The 0-based index of the row, where to set the filter. Add an auto
   *        filter on all columns in the current sheet.
   */
  public void autoFilterAllColumns (@Nonnegative final int nRowIndex)
  {
    // Set auto filter on all columns
    // Use the specified row (param1, param2)
    // From first column to last column (param3, param4)
    m_aLastSheet.setAutoFilter (new CellRangeAddress (nRowIndex, nRowIndex, 0, m_nMaxCellIndex - 1));
  }

  /**
   * Write the current workbook to a file
   *
   * @param aFile
   *        The file to write to. May not be <code>null</code>.
   * @return {@link ESuccess}
   */
  @Nonnull
  public ESuccess writeTo (@Nonnull final File aFile)
  {
    return writeTo (FileHelper.getOutputStream (aFile));
  }

  /**
   * Write the current workbook to an output stream.
   *
   * @param aOS
   *        The output stream to write to. May not be <code>null</code>. Is
   *        automatically closed independent of the success state.
   * @return {@link ESuccess}
   */
  @Nonnull
  public ESuccess writeTo (@Nonnull @WillClose final OutputStream aOS)
  {
    try
    {
      ValueEnforcer.notNull (aOS, "OutputStream");

      if (m_nCreatedCellStyles > 0 && s_aLogger.isDebugEnabled ())
        s_aLogger.debug ("Writing Excel workbook with " + m_nCreatedCellStyles + " different cell styles");

      m_aWB.write (aOS);
      return ESuccess.SUCCESS;
    }
    catch (final IOException ex)
    {
      if (!StreamHelper.isKnownEOFException (ex))
        s_aLogger.error ("Failed to write Excel workbook to output stream " + aOS, ex);
      return ESuccess.FAILURE;
    }
    finally
    {
      StreamHelper.close (aOS);
    }
  }
}
