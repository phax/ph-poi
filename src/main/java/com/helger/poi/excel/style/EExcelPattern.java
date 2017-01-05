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
package com.helger.poi.excel.style;

import javax.annotation.Nonnull;

import org.apache.poi.ss.usermodel.FillPatternType;

/**
 * Excel pattern types
 *
 * @author Philip Helger
 * @deprecated Use {@link FillPatternType} instead
 */
@Deprecated
public enum EExcelPattern
{
  NO_FILL (FillPatternType.NO_FILL),
  SOLID_FOREGROUND (FillPatternType.SOLID_FOREGROUND),
  FINE_DOTS (FillPatternType.FINE_DOTS),
  ALT_BARS (FillPatternType.ALT_BARS),
  SPARSE_DOTS (FillPatternType.SPARSE_DOTS),
  THICK_HORZ_BANDS (FillPatternType.THICK_HORZ_BANDS),
  THICK_VERT_BANDS (FillPatternType.THICK_VERT_BANDS),
  THICK_BACKWARD_DIAG (FillPatternType.THICK_BACKWARD_DIAG),
  THICK_FORWARD_DIAG (FillPatternType.THICK_FORWARD_DIAG),
  BIG_SPOTS (FillPatternType.BIG_SPOTS),
  BRICKS (FillPatternType.BRICKS),
  THIN_HORZ_BANDS (FillPatternType.THIN_HORZ_BANDS),
  THIN_VERT_BANDS (FillPatternType.THIN_VERT_BANDS),
  THIN_BACKWARD_DIAG (FillPatternType.THIN_BACKWARD_DIAG),
  THIN_FORWARD_DIAG (FillPatternType.THIN_FORWARD_DIAG),
  SQUARES (FillPatternType.SQUARES),
  DIAMONDS (FillPatternType.DIAMONDS),
  LESS_DOTS (FillPatternType.LESS_DOTS),
  LEAST_DOTS (FillPatternType.LEAST_DOTS);

  private final FillPatternType m_eValue;

  private EExcelPattern (@Nonnull final FillPatternType eValue)
  {
    m_eValue = eValue;
  }

  @Nonnull
  public FillPatternType getValue ()
  {
    return m_eValue;
  }
}
