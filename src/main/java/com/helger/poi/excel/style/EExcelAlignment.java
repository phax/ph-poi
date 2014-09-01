/**
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

import org.apache.poi.ss.usermodel.CellStyle;

/**
 * Excel horizontal alignment enum.
 * 
 * @author Philip Helger
 */
public enum EExcelAlignment
{
  ALIGN_GENERAL (CellStyle.ALIGN_GENERAL),
  ALIGN_LEFT (CellStyle.ALIGN_LEFT),
  ALIGN_CENTER (CellStyle.ALIGN_CENTER),
  ALIGN_RIGHT (CellStyle.ALIGN_RIGHT),
  ALIGN_FILL (CellStyle.ALIGN_FILL),
  ALIGN_JUSTIFY (CellStyle.ALIGN_JUSTIFY),
  ALIGN_CENTER_SELECTION (CellStyle.ALIGN_CENTER_SELECTION);

  private final short m_nValue;

  private EExcelAlignment (final short nValue)
  {
    m_nValue = nValue;
  }

  public short getValue ()
  {
    return m_nValue;
  }
}
