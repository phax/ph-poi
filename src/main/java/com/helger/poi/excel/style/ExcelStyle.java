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
package com.helger.poi.excel.style;

import javax.annotation.Nonnull;
import javax.annotation.Nullable;
import javax.annotation.concurrent.NotThreadSafe;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Workbook;

import com.helger.commons.ICloneable;
import com.helger.commons.ValueEnforcer;
import com.helger.commons.equals.EqualsUtils;
import com.helger.commons.hash.HashCodeGenerator;
import com.helger.commons.string.ToStringGenerator;

/**
 * Represents a single excel style with enums instead of numeric values.
 *
 * @author Philip Helger
 */
@NotThreadSafe
public final class ExcelStyle implements ICloneable <ExcelStyle>
{
  /** By default text wrapping is disabled */
  public static final boolean DEFAULT_WRAP_TEXT = false;

  private EExcelAlignment m_eAlign;
  private EExcelVerticalAlignment m_eVAlign;
  private boolean m_bWrapText = DEFAULT_WRAP_TEXT;
  private String m_sDataFormat;
  private IndexedColors m_eFillBackgroundColor;
  private IndexedColors m_eFillForegroundColor;
  private EExcelPattern m_eFillPattern;
  private EExcelBorder m_eBorderTop;
  private EExcelBorder m_eBorderRight;
  private EExcelBorder m_eBorderBottom;
  private EExcelBorder m_eBorderLeft;
  private short m_nFontIndex = -1;

  public ExcelStyle ()
  {}

  public ExcelStyle (@Nonnull final ExcelStyle rhs)
  {
    m_eAlign = rhs.m_eAlign;
    m_eVAlign = rhs.m_eVAlign;
    m_bWrapText = rhs.m_bWrapText;
    m_sDataFormat = rhs.m_sDataFormat;
    m_eFillBackgroundColor = rhs.m_eFillBackgroundColor;
    m_eFillForegroundColor = rhs.m_eFillForegroundColor;
    m_eFillPattern = rhs.m_eFillPattern;
    m_eBorderTop = rhs.m_eBorderTop;
    m_eBorderRight = rhs.m_eBorderRight;
    m_eBorderBottom = rhs.m_eBorderBottom;
    m_eBorderLeft = rhs.m_eBorderLeft;
    m_nFontIndex = rhs.m_nFontIndex;
  }

  @Nullable
  public EExcelAlignment getAlign ()
  {
    return m_eAlign;
  }

  @Nonnull
  public ExcelStyle setAlign (@Nullable final EExcelAlignment eAlign)
  {
    m_eAlign = eAlign;
    return this;
  }

  @Nullable
  public EExcelVerticalAlignment getVerticalAlign ()
  {
    return m_eVAlign;
  }

  @Nonnull
  public ExcelStyle setVerticalAlign (@Nullable final EExcelVerticalAlignment eVAlign)
  {
    m_eVAlign = eVAlign;
    return this;
  }

  public boolean isWrapText ()
  {
    return m_bWrapText;
  }

  @Nonnull
  public ExcelStyle setWrapText (final boolean bWrapText)
  {
    m_bWrapText = bWrapText;
    return this;
  }

  @Nullable
  public String getDataFormat ()
  {
    return m_sDataFormat;
  }

  @Nonnull
  public ExcelStyle setDataFormat (@Nullable final String sDataFormat)
  {
    m_sDataFormat = sDataFormat;
    return this;
  }

  @Nullable
  public IndexedColors getFillBackgroundColor ()
  {
    return m_eFillBackgroundColor;
  }

  @Nonnull
  public ExcelStyle setFillBackgroundColor (@Nullable final IndexedColors eColor)
  {
    m_eFillBackgroundColor = eColor;
    return this;
  }

  @Nullable
  public IndexedColors getFillForegroundColor ()
  {
    return m_eFillForegroundColor;
  }

  @Nonnull
  public ExcelStyle setFillForegroundColor (@Nullable final IndexedColors eColor)
  {
    m_eFillForegroundColor = eColor;
    return this;
  }

  @Nullable
  public EExcelPattern getFillPattern ()
  {
    return m_eFillPattern;
  }

  @Nonnull
  public ExcelStyle setFillPattern (@Nullable final EExcelPattern ePattern)
  {
    m_eFillPattern = ePattern;
    return this;
  }

  @Nullable
  public EExcelBorder getBorderTop ()
  {
    return m_eBorderTop;
  }

  @Nonnull
  public ExcelStyle setBorderTop (@Nullable final EExcelBorder eBorder)
  {
    m_eBorderTop = eBorder;
    return this;
  }

  @Nullable
  public EExcelBorder getBorderRight ()
  {
    return m_eBorderRight;
  }

  @Nonnull
  public ExcelStyle setBorderRight (@Nullable final EExcelBorder eBorder)
  {
    m_eBorderRight = eBorder;
    return this;
  }

  @Nullable
  public EExcelBorder getBorderBottom ()
  {
    return m_eBorderBottom;
  }

  @Nonnull
  public ExcelStyle setBorderBottom (@Nullable final EExcelBorder eBorder)
  {
    m_eBorderBottom = eBorder;
    return this;
  }

  @Nullable
  public EExcelBorder getBorderLeft ()
  {
    return m_eBorderLeft;
  }

  @Nonnull
  public ExcelStyle setBorderLeft (@Nullable final EExcelBorder eBorder)
  {
    m_eBorderLeft = eBorder;
    return this;
  }

  @Nonnull
  public ExcelStyle setBorder (@Nullable final EExcelBorder eBorder)
  {
    return setBorderTop (eBorder).setBorderRight (eBorder).setBorderBottom (eBorder).setBorderLeft (eBorder);
  }

  public short getFontIndex ()
  {
    return m_nFontIndex;
  }

  /**
   * Set the index of the font to use. The font must have been previously
   * created via Workbook.createFont()!
   *
   * @param nFontIndex
   *        The font index to use. Values &lt; 0 indicate no font to use
   * @return this
   */
  @Nonnull
  public ExcelStyle setFontIndex (final short nFontIndex)
  {
    m_nFontIndex = nFontIndex;
    return this;
  }

  /**
   * Set the index of the font to use. The font must have been previously
   * created via Workbook.createFont()!
   *
   * @param aFont
   *        The font to use. May not be <code>null</code>.
   * @return this
   */
  @Nonnull
  public ExcelStyle setFont (@Nonnull final Font aFont)
  {
    ValueEnforcer.notNull (aFont, "Font");
    return setFontIndex (aFont.getIndex ());
  }

  @Nonnull
  public ExcelStyle getClone ()
  {
    return new ExcelStyle (this);
  }

  public void fillCellStyle (@Nonnull final Workbook aWB,
                             @Nonnull final CellStyle aCS,
                             @Nonnull final CreationHelper aCreationHelper)
  {
    if (m_eAlign != null)
      aCS.setAlignment (m_eAlign.getValue ());
    if (m_eVAlign != null)
      aCS.setVerticalAlignment (m_eVAlign.getValue ());
    aCS.setWrapText (m_bWrapText);
    if (m_sDataFormat != null)
      aCS.setDataFormat (aCreationHelper.createDataFormat ().getFormat (m_sDataFormat));
    if (m_eFillBackgroundColor != null)
      aCS.setFillBackgroundColor (m_eFillBackgroundColor.getIndex ());
    if (m_eFillForegroundColor != null)
      aCS.setFillForegroundColor (m_eFillForegroundColor.getIndex ());
    if (m_eFillPattern != null)
      aCS.setFillPattern (m_eFillPattern.getValue ());
    if (m_eBorderTop != null)
      aCS.setBorderTop (m_eBorderTop.getValue ());
    if (m_eBorderRight != null)
      aCS.setBorderRight (m_eBorderRight.getValue ());
    if (m_eBorderBottom != null)
      aCS.setBorderBottom (m_eBorderBottom.getValue ());
    if (m_eBorderLeft != null)
      aCS.setBorderLeft (m_eBorderLeft.getValue ());
    if (m_nFontIndex >= 0)
      aCS.setFont (aWB.getFontAt (m_nFontIndex));
  }

  @Override
  public boolean equals (final Object o)
  {
    if (o == this)
      return true;
    if (!(o instanceof ExcelStyle))
      return false;
    final ExcelStyle rhs = (ExcelStyle) o;
    return EqualsUtils.equals (m_eAlign, rhs.m_eAlign) &&
           EqualsUtils.equals (m_eVAlign, rhs.m_eVAlign) &&
           m_bWrapText == rhs.m_bWrapText &&
           EqualsUtils.equals (m_sDataFormat, rhs.m_sDataFormat) &&
           EqualsUtils.equals (m_eFillBackgroundColor, rhs.m_eFillBackgroundColor) &&
           EqualsUtils.equals (m_eFillForegroundColor, rhs.m_eFillForegroundColor) &&
           EqualsUtils.equals (m_eFillPattern, rhs.m_eFillPattern) &&
           EqualsUtils.equals (m_eBorderTop, rhs.m_eBorderTop) &&
           EqualsUtils.equals (m_eBorderRight, rhs.m_eBorderRight) &&
           EqualsUtils.equals (m_eBorderBottom, rhs.m_eBorderBottom) &&
           EqualsUtils.equals (m_eBorderLeft, rhs.m_eBorderLeft) &&
           m_nFontIndex == rhs.m_nFontIndex;
  }

  @Override
  public int hashCode ()
  {
    return new HashCodeGenerator (this).append (m_eAlign)
                                       .append (m_eVAlign)
                                       .append (m_bWrapText)
                                       .append (m_sDataFormat)
                                       .append (m_eFillBackgroundColor)
                                       .append (m_eFillForegroundColor)
                                       .append (m_eFillPattern)
                                       .append (m_eBorderTop)
                                       .append (m_eBorderRight)
                                       .append (m_eBorderBottom)
                                       .append (m_eBorderLeft)
                                       .append (m_nFontIndex)
                                       .getHashCode ();
  }

  @Override
  public String toString ()
  {
    return new ToStringGenerator (this).appendIfNotNull ("align", m_eAlign)
                                       .appendIfNotNull ("verticalAlign", m_eVAlign)
                                       .append ("wrapText", m_bWrapText)
                                       .appendIfNotNull ("dataFormat", m_sDataFormat)
                                       .appendIfNotNull ("fillBackgroundColor", m_eFillBackgroundColor)
                                       .appendIfNotNull ("fillForegroundColor", m_eFillForegroundColor)
                                       .appendIfNotNull ("fillPattern", m_eFillPattern)
                                       .appendIfNotNull ("borderTop", m_eBorderTop)
                                       .appendIfNotNull ("borderRight", m_eBorderRight)
                                       .appendIfNotNull ("borderBottom", m_eBorderBottom)
                                       .appendIfNotNull ("borderLeft", m_eBorderLeft)
                                       .append ("fontIndex", m_nFontIndex)
                                       .toString ();
  }
}
