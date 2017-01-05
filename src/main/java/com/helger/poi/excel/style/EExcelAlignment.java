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

import org.apache.poi.ss.usermodel.HorizontalAlignment;

/**
 * Excel horizontal alignment enum.
 *
 * @author Philip Helger
 * @deprecated Use {@link HorizontalAlignment} instead
 */
@Deprecated
public enum EExcelAlignment
{
  GENERAL (HorizontalAlignment.GENERAL),
  LEFT (HorizontalAlignment.LEFT),
  CENTER (HorizontalAlignment.CENTER),
  RIGHT (HorizontalAlignment.RIGHT),
  FILL (HorizontalAlignment.FILL),
  JUSTIFY (HorizontalAlignment.JUSTIFY),
  CENTER_SELECTION (HorizontalAlignment.CENTER_SELECTION),
  DISTRIBUTED (HorizontalAlignment.DISTRIBUTED);

  private final HorizontalAlignment m_eValue;

  private EExcelAlignment (@Nonnull final HorizontalAlignment eValue)
  {
    m_eValue = eValue;
  }

  @Nonnull
  public HorizontalAlignment getValue ()
  {
    return m_eValue;
  }
}
