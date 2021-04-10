/**
 * Copyright (C) 2014-2021 Philip Helger (www.helger.com)
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
import java.io.UncheckedIOException;
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
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.WorkbookUtil;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.helger.commons.ValueEnforcer;
import com.helger.commons.datetime.PDTFactory;
import com.helger.commons.io.EAppend;
import com.helger.commons.io.file.FileHelper;
import com.helger.commons.io.resource.IWritableResource;
import com.helger.commons.io.stream.NonBlockingByteArrayOutputStream;
import com.helger.commons.io.stream.StreamHelper;
import com.helger.commons.state.ESuccess;
import com.helger.poi.excel.style.ExcelStyle;
import com.helger.poi.excel.style.ExcelStyleCache;

/**
 * A utility class for creating very simple Excel workbooks.
 *
 * @author Philip Helger
 */
public final class WorkbookCreationHelper implements AutoCloseable
{
  private static final Logger LOGGER = LoggerFactory.getLogger (WorkbookCreationHelper.class);

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

  public void close ()
  {
    try
    {
      m_aWB.close ();
    }
    catch (final IOException ex)
    {
      throw new UncheckedIOException (ex);
    }
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
    m_aLastSheet = sName == null ? m_aWB.createSheet () : m_aWB.createSheet (WorkbookUtil.createSafeSheetName (sName));
    m_nLastSheetRowIndex = 0;
    m_aLastRow = null;
    m_nLastRowCellIndex = 0;
    m_aLastCell = null;
    m_nMaxCellIndex = 0;
    return m_aLastSheet;
  }

  private void _ensureSheet ()
  {
    if (m_aLastSheet == null)
      throw new IllegalStateException ("A sheet needs to be present to perform this! Call createNewSheet");
  }

  /**
   * @return A new row in the current sheet.
   */
  @Nonnull
  public Row addRow ()
  {
    _ensureSheet ();
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

  private void _ensureRow ()
  {
    if (m_aLastRow == null)
      throw new IllegalStateException ("A row needs to be present to perform this! Call addRow");
  }

  /**
   * @return A new cell in the current row of the current sheet
   */
  @Nonnull
  public Cell addCell ()
  {
    _ensureRow ();
    m_aLastCell = m_aLastRow.createCell (m_nLastRowCellIndex++);
    m_aLastCell.setBlank ();

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
    aCell.setCellValue (bValue);
    return aCell;
  }

  /**
   * Added a new cell as date/time.
   * <p>
   * Important: don't forget to call {@link #addCellStyle(ExcelStyle)} with
   * something like <code>new ExcelStyle ().setDataFormat ("dd.mm.yyyy");</code>
   * after a date/time cell!
   * <p>
   * Important: Excel cannot correctly handle dates/times before
   * {@link CExcel#EXCEL_MINIMUM_DATE 1900-01-01}
   *
   * @param aValue
   *        The value to be set. May be <code>null</code>.
   * @return A new cell in the current row of the current sheet with the passed
   *         value
   */
  @Nonnull
  public Cell addCell (@Nullable final Calendar aValue)
  {
    final Cell aCell = addCell ();
    if (aValue != null)
      aCell.setCellValue (aValue);
    return aCell;
  }

  /**
   * Added a new cell as date/time.
   * <p>
   * Important: don't forget to call {@link #addCellStyle(ExcelStyle)} with
   * something like <code>new ExcelStyle ().setDataFormat ("dd.mm.yyyy");</code>
   * after a date/time cell!
   * <p>
   * Important: Excel cannot correctly handle dates/times before
   * {@link CExcel#EXCEL_MINIMUM_DATE 1900-01-01}
   *
   * @param aValue
   *        The value to be set. May be <code>null</code>.
   * @return A new cell in the current row of the current sheet with the passed
   *         value
   */
  @Nonnull
  public Cell addCell (@Nullable final Date aValue)
  {
    final Cell aCell = addCell ();
    if (aValue != null)
      aCell.setCellValue (aValue);
    return aCell;
  }

  /**
   * Added a new cell as date/time.
   * <p>
   * Important: don't forget to call {@link #addCellStyle(ExcelStyle)} with
   * something like <code>new ExcelStyle ().setDataFormat ("dd.mm.yyyy");</code>
   * after a date/time cell!
   * <p>
   * Important: Excel cannot correctly handle dates/times before
   * {@link CExcel#EXCEL_MINIMUM_DATE 1900-01-01}
   *
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
   * Added a new cell as date/time.
   * <p>
   * Important: don't forget to call {@link #addCellStyle(ExcelStyle)} with
   * something like <code>new ExcelStyle ().setDataFormat ("dd.mm.yyyy");</code>
   * after a date/time cell!
   * <p>
   * Important: Excel cannot correctly handle dates/times before
   * {@link CExcel#EXCEL_MINIMUM_DATE 1900-01-01}
   *
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
   * Added a new cell as date/time.
   * <p>
   * Important: don't forget to call {@link #addCellStyle(ExcelStyle)} with
   * something like <code>new ExcelStyle ().setDataFormat ("dd.mm.yyyy");</code>
   * after a date/time cell!
   * <p>
   * Important: Excel cannot correctly handle dates/times before
   * {@link CExcel#EXCEL_MINIMUM_DATE 1900-01-01}
   *
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

    if (CExcel.canBeNumericValue (aValue))
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
    _ensureSheet ();
    return m_aLastSheet.addMergedRegion (new CellRangeAddress (nFirstRow, nLastRow, nFirstCol, nLastCol));
  }

  private void _ensureCell ()
  {
    if (m_aLastCell == null)
      throw new IllegalStateException ("A cell needs to be present to perform this! Call addCell");
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
    _ensureCell ();

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
    _ensureSheet ();

    // auto-adjust all columns (except description and image description)
    for (short nCol = 0; nCol < m_nMaxCellIndex; ++nCol)
      try
      {
        m_aLastSheet.autoSizeColumn (nCol);
      }
      catch (final IllegalArgumentException ex)
      {
        // Happens if a column is too large
        LOGGER.warn ("Failed to resize column " + nCol + ": column too wide!");
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
   * Add an auto filter on all columns in the current sheet.
   *
   * @param nRowIndex
   *        The 0-based index of the row, where to set the filter.
   */
  public void autoFilterAllColumns (@Nonnegative final int nRowIndex)
  {
    _ensureSheet ();

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
   * Write the current workbook to a writable resource.
   *
   * @param aRes
   *        The resource to write to. May not be <code>null</code>.
   * @return {@link ESuccess}
   */
  @Nonnull
  public ESuccess writeTo (@Nonnull final IWritableResource aRes)
  {
    return writeTo (aRes.getOutputStream (EAppend.TRUNCATE));
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

      if (m_nCreatedCellStyles > 0 && LOGGER.isDebugEnabled ())
        LOGGER.debug ("Writing Excel workbook with " + m_nCreatedCellStyles + " different cell styles");

      m_aWB.write (aOS);
      return ESuccess.SUCCESS;
    }
    catch (final IOException ex)
    {
      if (!StreamHelper.isKnownEOFException (ex))
        LOGGER.error ("Failed to write Excel workbook to output stream " + aOS, ex);
      return ESuccess.FAILURE;
    }
    finally
    {
      StreamHelper.close (aOS);
    }
  }

  /**
   * Helper method to get the whole workbook as a single byte array.
   *
   * @return <code>null</code> if writing failed. See log files for details.
   */
  @Nullable
  public byte [] getAsByteArray ()
  {
    try (final NonBlockingByteArrayOutputStream aBAOS = new NonBlockingByteArrayOutputStream ())
    {
      if (writeTo (aBAOS).isFailure ())
        return null;
      return aBAOS.getBufferOrCopy ();
    }
  }
}
