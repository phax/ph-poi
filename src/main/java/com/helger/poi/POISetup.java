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
package com.helger.poi;

import java.util.concurrent.atomic.AtomicBoolean;

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
  public static final String SYS_PROP_POI_LOGGER = "org.apache.poi.util.POILogger";
  private static final AtomicBoolean s_aInited = new AtomicBoolean (false);

  static
  {
    // Workaround some annoying POI 3.10+ log messages
    SystemProperties.setPropertyValue ("HSSFWorkbook.SheetInitialCapacity", 1);
    SystemProperties.setPropertyValue ("HSSFSheet.RowInitialCapacity", 20);
    SystemProperties.setPropertyValue ("HSSFRow.ColInitialCapacity", 5);
  }

  private POISetup ()
  {}

  public static void enableCustomLogger (final boolean bEnable)
  {
    if (bEnable)
      SystemProperties.setPropertyValue (SYS_PROP_POI_LOGGER, POISLF4JLogger.class.getName ());
    else
      SystemProperties.removePropertyValue (SYS_PROP_POI_LOGGER);
  }

  public static boolean isInited ()
  {
    return s_aInited.get ();
  }

  public static void initOnDemand ()
  {
    if (s_aInited.compareAndSet (false, true))
      enableCustomLogger (true);
  }
}
