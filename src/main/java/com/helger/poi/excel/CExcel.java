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

import java.math.BigInteger;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.Month;

import javax.annotation.Nonnull;

import com.helger.commons.CGlobal;
import com.helger.commons.datetime.PDTFactory;

/**
 * Excel constants.
 *
 * @author Philip Helger
 */
public final class CExcel
{
  /** The minimum date Excel can handle as such */
  public static final LocalDate EXCEL_MINIMUM_DATE = PDTFactory.createLocalDate (1900, Month.JANUARY, 1);

  /** The minimum datetime Excel can handle as such */
  public static final LocalDateTime EXCEL_MINIMUM_DATE_TIME = EXCEL_MINIMUM_DATE.atStartOfDay ();

  /** Minimum number Excel can represent as a number */
  public static final BigInteger EXCEL_MINIMUM_NUMBER = CGlobal.BIGINT_MIN_LONG;

  /** Maximum number Excel can represent as a number */
  public static final BigInteger EXCEL_MAXIMUM_NUMBER = CGlobal.BIGINT_MAX_LONG;

  private CExcel ()
  {}

  public static boolean canBeNumericValue (@Nonnull final BigInteger aValue)
  {
    return aValue != null && aValue.compareTo (CExcel.EXCEL_MINIMUM_NUMBER) >= 0 && aValue.compareTo (CExcel.EXCEL_MAXIMUM_NUMBER) <= 0;
  }
}
