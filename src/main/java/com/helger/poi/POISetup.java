/*
 * Copyright (C) 2014-2022 Philip Helger (www.helger.com)
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
package com.helger.poi;

import javax.annotation.Nonnull;
import javax.annotation.concurrent.ThreadSafe;

import com.helger.commons.system.SystemProperties;

/**
 * This class can be used to initialize POI to work best with the "ph" stack.
 *
 * @author Philip Helger
 */
@ThreadSafe
public final class POISetup
{
  private static void _setValue (@Nonnull final String sKey, final int nValue)
  {
    if (!SystemProperties.containsPropertyName (sKey))
      SystemProperties.setPropertyValue (sKey, nValue);
  }

  static
  {
    // Workaround some annoying POI 3.10+ log messages
    _setValue ("HSSFWorkbook.SheetInitialCapacity", 1);
    _setValue ("HSSFSheet.RowInitialCapacity", 20);
    _setValue ("HSSFRow.ColInitialCapacity", 5);
  }

  private POISetup ()
  {}

  public static void initOnDemand ()
  {
    // empty
  }
}
