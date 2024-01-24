/*
 * Copyright (C) 2014-2024 Philip Helger (www.helger.com)
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

import java.io.Serializable;

import javax.annotation.Nonnull;
import javax.annotation.Nullable;
import javax.annotation.concurrent.NotThreadSafe;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;

import com.helger.commons.ValueEnforcer;
import com.helger.commons.equals.EqualsHelper;
import com.helger.commons.hashcode.HashCodeGenerator;
import com.helger.commons.lang.ICloneable;
import com.helger.commons.string.ToStringGenerator;

/**
 * Represents a single excel style.
 *
 * @author Philip Helger
 */
@NotThreadSafe
public class ExcelStyle implements ICloneable <ExcelStyle>, Serializable
{
  /** By default text wrapping is disabled */
  public static final boolean DEFAULT_WRAP_TEXT = false;

  private HorizontalAlignment m_eAlign;
  private VerticalAlignment m_eVAlign;
  private boolean m_bWrapText = DEFAULT_WRAP_TEXT;
  private String m_sDataFormat;
  private IndexedColors m_eFillBackgroundColor;
  private IndexedColors m_eFillForegroundColor;
  private FillPatternType m_eFillPattern;
  private BorderStyle m_eBorderTop;
  private BorderStyle m_eBorderRight;
  private BorderStyle m_eBorderBottom;
  private BorderStyle m_eBorderLeft;
  private int m_nFontIndex = -1;

  public ExcelStyle ()
  {}

  public ExcelStyle (@Nonnull final ExcelStyle aOther)
  {
    ValueEnforcer.notNull (aOther, "Other");
    m_eAlign = aOther.m_eAlign;
    m_eVAlign = aOther.m_eVAlign;
    m_bWrapText = aOther.m_bWrapText;
    m_sDataFormat = aOther.m_sDataFormat;
    m_eFillBackgroundColor = aOther.m_eFillBackgroundColor;
    m_eFillForegroundColor = aOther.m_eFillForegroundColor;
    m_eFillPattern = aOther.m_eFillPattern;
    m_eBorderTop = aOther.m_eBorderTop;
    m_eBorderRight = aOther.m_eBorderRight;
    m_eBorderBottom = aOther.m_eBorderBottom;
    m_eBorderLeft = aOther.m_eBorderLeft;
    m_nFontIndex = aOther.m_nFontIndex;
  }

  @Nullable
  public HorizontalAlignment getAlign ()
  {
    return m_eAlign;
  }

  @Nonnull
  public ExcelStyle setAlign (@Nullable final HorizontalAlignment eAlign)
  {
    m_eAlign = eAlign;
    return this;
  }

  @Nullable
  public VerticalAlignment getVerticalAlign ()
  {
    return m_eVAlign;
  }

  @Nonnull
  public ExcelStyle setVerticalAlign (@Nullable final VerticalAlignment eVAlign)
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
  public FillPatternType getFillPattern ()
  {
    return m_eFillPattern;
  }

  @Nonnull
  public ExcelStyle setFillPattern (@Nullable final FillPatternType ePattern)
  {
    m_eFillPattern = ePattern;
    return this;
  }

  @Nullable
  public BorderStyle getBorderTop ()
  {
    return m_eBorderTop;
  }

  @Nonnull
  public ExcelStyle setBorderTop (@Nullable final BorderStyle eBorder)
  {
    m_eBorderTop = eBorder;
    return this;
  }

  @Nullable
  public BorderStyle getBorderRight ()
  {
    return m_eBorderRight;
  }

  @Nonnull
  public ExcelStyle setBorderRight (@Nullable final BorderStyle eBorder)
  {
    m_eBorderRight = eBorder;
    return this;
  }

  @Nullable
  public BorderStyle getBorderBottom ()
  {
    return m_eBorderBottom;
  }

  @Nonnull
  public ExcelStyle setBorderBottom (@Nullable final BorderStyle eBorder)
  {
    m_eBorderBottom = eBorder;
    return this;
  }

  @Nullable
  public BorderStyle getBorderLeft ()
  {
    return m_eBorderLeft;
  }

  @Nonnull
  public ExcelStyle setBorderLeft (@Nullable final BorderStyle eBorder)
  {
    m_eBorderLeft = eBorder;
    return this;
  }

  @Nonnull
  public ExcelStyle setBorder (@Nullable final BorderStyle eBorder)
  {
    return setBorderTop (eBorder).setBorderRight (eBorder).setBorderBottom (eBorder).setBorderLeft (eBorder);
  }

  public int getFontIndex ()
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
  public ExcelStyle setFontIndex (final int nFontIndex)
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

  public void fillCellStyle (@Nonnull final Workbook aWB, @Nonnull final CellStyle aCS, @Nonnull final CreationHelper aCreationHelper)
  {
    if (m_eAlign != null)
      aCS.setAlignment (m_eAlign);
    if (m_eVAlign != null)
      aCS.setVerticalAlignment (m_eVAlign);
    aCS.setWrapText (m_bWrapText);
    if (m_sDataFormat != null)
      aCS.setDataFormat (aCreationHelper.createDataFormat ().getFormat (m_sDataFormat));
    if (m_eFillBackgroundColor != null)
      aCS.setFillBackgroundColor (m_eFillBackgroundColor.getIndex ());
    if (m_eFillForegroundColor != null)
      aCS.setFillForegroundColor (m_eFillForegroundColor.getIndex ());
    if (m_eFillPattern != null)
      aCS.setFillPattern (m_eFillPattern);
    if (m_eBorderTop != null)
      aCS.setBorderTop (m_eBorderTop);
    if (m_eBorderRight != null)
      aCS.setBorderRight (m_eBorderRight);
    if (m_eBorderBottom != null)
      aCS.setBorderBottom (m_eBorderBottom);
    if (m_eBorderLeft != null)
      aCS.setBorderLeft (m_eBorderLeft);
    if (m_nFontIndex >= 0)
      aCS.setFont (aWB.getFontAt (m_nFontIndex));
  }

  @Override
  public boolean equals (final Object o)
  {
    if (o == this)
      return true;
    if (o == null || !getClass ().equals (o.getClass ()))
      return false;
    final ExcelStyle rhs = (ExcelStyle) o;
    return EqualsHelper.equals (m_eAlign, rhs.m_eAlign) &&
           EqualsHelper.equals (m_eVAlign, rhs.m_eVAlign) &&
           m_bWrapText == rhs.m_bWrapText &&
           EqualsHelper.equals (m_sDataFormat, rhs.m_sDataFormat) &&
           EqualsHelper.equals (m_eFillBackgroundColor, rhs.m_eFillBackgroundColor) &&
           EqualsHelper.equals (m_eFillForegroundColor, rhs.m_eFillForegroundColor) &&
           EqualsHelper.equals (m_eFillPattern, rhs.m_eFillPattern) &&
           EqualsHelper.equals (m_eBorderTop, rhs.m_eBorderTop) &&
           EqualsHelper.equals (m_eBorderRight, rhs.m_eBorderRight) &&
           EqualsHelper.equals (m_eBorderBottom, rhs.m_eBorderBottom) &&
           EqualsHelper.equals (m_eBorderLeft, rhs.m_eBorderLeft) &&
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
                                       .getToString ();
  }
}
