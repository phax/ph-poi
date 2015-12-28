/**
 * Copyright (C) 2014-2015 Philip Helger (www.helger.com)
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

import org.apache.poi.ss.usermodel.CellStyle;

/**
 * Excel border styles.
 *
 * @author Philip Helger
 */
public enum EExcelBorder
{
  BORDER_NONE (CellStyle.BORDER_NONE),
  BORDER_THIN (CellStyle.BORDER_THIN),
  BORDER_MEDIUM (CellStyle.BORDER_MEDIUM),
  BORDER_DASHED (CellStyle.BORDER_DASHED),
  BORDER_HAIR (CellStyle.BORDER_HAIR),
  BORDER_THICK (CellStyle.BORDER_THICK),
  BORDER_DOUBLE (CellStyle.BORDER_DOUBLE),
  BORDER_DOTTED (CellStyle.BORDER_DOTTED),
  BORDER_MEDIUM_DASHED (CellStyle.BORDER_MEDIUM_DASHED),
  BORDER_DASH_DOT (CellStyle.BORDER_DASH_DOT),
  BORDER_MEDIUM_DASH_DOT (CellStyle.BORDER_MEDIUM_DASH_DOT),
  BORDER_DASH_DOT_DOT (CellStyle.BORDER_DASH_DOT_DOT),
  BORDER_MEDIUM_DASH_DOT_DOT (CellStyle.BORDER_MEDIUM_DASH_DOT_DOT),
  BORDER_SLANTED_DASH_DOT (CellStyle.BORDER_SLANTED_DASH_DOT);

  private final short m_nValue;

  private EExcelBorder (final short nValue)
  {
    m_nValue = nValue;
  }

  public short getValue ()
  {
    return m_nValue;
  }
}
