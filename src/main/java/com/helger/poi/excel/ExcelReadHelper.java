/*
 * Copyright (C) 2014-2023 Philip Helger (www.helger.com)
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

import java.io.IOException;
import java.io.InputStream;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.util.Date;

import javax.annotation.Nonnull;
import javax.annotation.Nullable;
import javax.annotation.concurrent.Immutable;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.NotOLE2FileException;
import org.apache.poi.poifs.filesystem.OfficeXmlFileException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.helger.commons.datetime.PDTFactory;
import com.helger.commons.io.IHasInputStream;
import com.helger.commons.io.stream.StreamHelper;
import com.helger.commons.string.StringHelper;

import edu.umd.cs.findbugs.annotations.SuppressFBWarnings;

/**
 * Misc Excel read helper methods.
 *
 * @author Philip Helger
 */
@Immutable
@SuppressFBWarnings ("JCIP_FIELD_ISNT_FINAL_IN_IMMUTABLE_CLASS")
public final class ExcelReadHelper
{
  private static final Logger LOGGER = LoggerFactory.getLogger (ExcelReadHelper.class);

  private ExcelReadHelper ()
  {}

  /**
   * Try to read an Excel {@link Workbook} from the passed
   * {@link IHasInputStream}. First XLS is tried, than XLSX, as XLS files can be
   * identified more easily.
   *
   * @param aIIS
   *        The input stream provider to read from.
   * @return <code>null</code> if the content of the InputStream could not be
   *         interpreted as Excel file
   */
  @Nullable
  public static Workbook readWorkbookFromInputStream (@Nonnull final IHasInputStream aIIS)
  {
    InputStream aIS = null;
    try
    {
      // Try to read as XLS
      aIS = aIIS.getInputStream ();
      if (aIS == null)
      {
        // Failed to open input stream -> no need to continue
        return null;
      }
      return new HSSFWorkbook (aIS);
    }
    catch (final OfficeXmlFileException ex)
    {
      // No XLS -> try XSLS
      StreamHelper.close (aIS);
      try
      {
        // Re-retrieve the input stream, to ensure we read from the beginning!
        aIS = aIIS.getInputStream ();
        return new XSSFWorkbook (aIS);
      }
      catch (final IOException ex2)
      {
        LOGGER.error ("Error trying to read XLSX file from " + aIIS, ex);
      }
    }
    catch (final NotOLE2FileException ex)
    {
      LOGGER.error ("Error trying to read non-Excel file from " + aIIS + ": " + ex.getMessage ());
    }
    catch (final IOException ex)
    {
      LOGGER.error ("Error trying to read XLS file from " + aIIS, ex);
    }
    finally
    {
      // Ensure the InputStream is closed. The data structures are in memory!
      StreamHelper.close (aIS);
    }
    return null;
  }

  @Nonnull
  private static Number _getAsNumberObject (final double dValue)
  {
    if (dValue == (int) dValue)
    {
      // It's not a real double value, it's an int value
      return Integer.valueOf ((int) dValue);
    }
    if (dValue == (long) dValue)
    {
      // It's not a real double value, it's a long value
      return Long.valueOf ((long) dValue);
    }
    // It's a real floating point number
    return Double.valueOf (dValue);
  }

  /**
   * Return the best matching Java object underlying the passed cell.<br>
   * Note: Date values cannot be determined automatically!
   *
   * @param aCell
   *        The cell to be queried. May be <code>null</code>.
   * @return <code>null</code> if the cell is <code>null</code> or if it is of
   *         type blank.
   */
  @Nullable
  public static Object getCellValueObject (@Nullable final Cell aCell)
  {
    if (aCell == null)
      return null;

    final CellType eCellType = aCell.getCellType ();
    switch (eCellType)
    {
      case NUMERIC:
        return _getAsNumberObject (aCell.getNumericCellValue ());
      case STRING:
        return aCell.getStringCellValue ();
      case BOOLEAN:
        return Boolean.valueOf (aCell.getBooleanCellValue ());
      case FORMULA:
        final CellType eFormulaResultType = aCell.getCachedFormulaResultType ();
        switch (eFormulaResultType)
        {
          case NUMERIC:
            return _getAsNumberObject (aCell.getNumericCellValue ());
          case STRING:
            return aCell.getStringCellValue ();
          case BOOLEAN:
            return Boolean.valueOf (aCell.getBooleanCellValue ());
          default:
            throw new IllegalArgumentException ("The cell formula type " + eFormulaResultType + " is unsupported!");
        }
      case BLANK:
        return null;
      default:
        throw new IllegalArgumentException ("The cell type " + eCellType + " is unsupported!");
    }
  }

  @Nullable
  public static String getCellValueString (@Nullable final Cell aCell)
  {
    final Object aObject = getCellValueObject (aCell);
    return aObject == null ? null : aObject.toString ();
  }

  @Nullable
  public static String getCellValueNormalizedString (@Nullable final Cell aCell)
  {
    final String sValue = getCellValueString (aCell);
    if (sValue == null)
      return null;

    // Remove all control characters
    final char [] aChars = sValue.toCharArray ();
    final StringBuilder aSB = new StringBuilder (aChars.length);
    for (final char c : aChars)
      if (Character.getType (c) != Character.CONTROL)
        aSB.append (c);

    // And trim away all unnecessary spaces
    return StringHelper.replaceAllRepeatedly (aSB.toString ().trim (), "  ", " ");
  }

  @Nullable
  @SuppressFBWarnings ("NP_BOOLEAN_RETURN_NULL")
  public static Boolean getCellValueBoolean (@Nullable final Cell aCell)
  {
    final Object aValue = getCellValueObject (aCell);
    if (aValue != null && !(aValue instanceof Boolean))
    {
      LOGGER.warn ("Failed to get cell value as boolean: " + aValue.getClass ());
      return null;
    }
    return (Boolean) aValue;
  }

  @Nullable
  public static Number getCellValueNumber (@Nullable final Cell aCell)
  {
    final Object aValue = getCellValueObject (aCell);
    if (aValue != null && !(aValue instanceof Number))
    {
      LOGGER.warn ("Failed to get cell value as number: " + aValue.getClass ());
      return null;
    }
    return (Number) aValue;
  }

  @Nullable
  public static Date getCellValueJavaDate (@Nullable final Cell aCell)
  {
    if (aCell != null)
      try
      {
        return aCell.getDateCellValue ();
      }
      catch (final RuntimeException ex)
      {
        // fall through
        LOGGER.warn ("Failed to get cell value as date: " + ex.getMessage ());
      }
    return null;
  }

  @Nullable
  public static LocalDateTime getCellValueLocalDateTime (@Nullable final Cell aCell)
  {
    final Date aDate = getCellValueJavaDate (aCell);
    return aDate == null ? null : PDTFactory.createLocalDateTime (aDate);
  }

  @Nullable
  public static LocalDate getCellValueLocalDate (@Nullable final Cell aCell)
  {
    final Date aDate = getCellValueJavaDate (aCell);
    return aDate == null ? null : PDTFactory.createLocalDate (aDate);
  }

  @Nullable
  public static LocalTime getCellValueLocalTime (@Nullable final Cell aCell)
  {
    final Date aDate = getCellValueJavaDate (aCell);
    return aDate == null ? null : PDTFactory.createLocalTime (aDate);
  }

  @Nullable
  public static RichTextString getCellValueRichText (@Nullable final Cell aCell)
  {
    return aCell == null ? null : aCell.getRichStringCellValue ();
  }

  @Nullable
  public static String getCellFormula (@Nullable final Cell aCell)
  {
    if (aCell != null)
      try
      {
        return aCell.getCellFormula ();
      }
      catch (final RuntimeException ex)
      {
        // fall through
        LOGGER.warn ("Failed to get cell formula: " + ex.getMessage ());
      }
    return null;
  }

  @Nullable
  public static Hyperlink getHyperlink (@Nullable final Cell aCell)
  {
    return aCell == null ? null : aCell.getHyperlink ();
  }

  public static boolean canBeReadAsNumericCell (@Nullable final Cell aCell)
  {
    if (aCell == null)
      return false;
    final CellType eType = aCell.getCellType ();
    return eType == CellType.BLANK || eType == CellType.NUMERIC || eType == CellType.FORMULA;
  }
}
