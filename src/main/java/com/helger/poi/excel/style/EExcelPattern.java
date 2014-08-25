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

import org.apache.poi.ss.usermodel.CellStyle;

/**
 * Excel pattern types
 * 
 * @author Philip Helger
 */
public enum EExcelPattern
{
  NO_FILL (CellStyle.NO_FILL),
  SOLID_FOREGROUND (CellStyle.SOLID_FOREGROUND),
  FINE_DOTS (CellStyle.FINE_DOTS),
  ALT_BARS (CellStyle.ALT_BARS),
  SPARSE_DOTS (CellStyle.SPARSE_DOTS),
  THICK_HORZ_BANDS (CellStyle.THICK_HORZ_BANDS),
  THICK_VERT_BANDS (CellStyle.THICK_VERT_BANDS),
  THICK_BACKWARD_DIAG (CellStyle.THICK_BACKWARD_DIAG),
  THICK_FORWARD_DIAG (CellStyle.THICK_FORWARD_DIAG),
  BIG_SPOTS (CellStyle.BIG_SPOTS),
  BRICKS (CellStyle.BRICKS),
  THIN_HORZ_BANDS (CellStyle.THIN_HORZ_BANDS),
  THIN_VERT_BANDS (CellStyle.THIN_VERT_BANDS),
  THIN_BACKWARD_DIAG (CellStyle.THIN_BACKWARD_DIAG),
  THIN_FORWARD_DIAG (CellStyle.THIN_FORWARD_DIAG),
  SQUARES (CellStyle.SQUARES),
  DIAMONDS (CellStyle.DIAMONDS),
  LESS_DOTS (CellStyle.LESS_DOTS),
  LEAST_DOTS (CellStyle.LEAST_DOTS);

  private final short m_nValue;

  private EExcelPattern (final short nValue)
  {
    m_nValue = nValue;
  }

  public short getValue ()
  {
    return m_nValue;
  }
}
