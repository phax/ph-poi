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
package com.helger.poi;

import javax.annotation.Nonnull;
import javax.annotation.Nullable;

import org.apache.poi.util.SystemOutLogger;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * A special {@link org.apache.poi.util.POILogger} implementation that uses
 * SLF4J. Set the system property "org.apache.poi.util.POILogger" to this class,
 * to use it (see {@link com.helger.poi.POISetup} for the code). It is derived
 * from {@link SystemOutLogger}, because the super class
 * {@link org.apache.poi.util.POILogger} only has a package-private ctor.
 *
 * @author Philip Helger
 */
public class POISLF4JLogger extends SystemOutLogger
{
  private static final String BLA = "An exception occurred";

  private Logger m_aLogger;

  public POISLF4JLogger ()
  {}

  @Override
  public void initialize (final String sCat)
  {
    m_aLogger = LoggerFactory.getLogger (sCat);
  }

  /**
   * Log a message
   *
   * @param nLevel
   *        One of DEBUG, INFO, WARN, ERROR, FATAL
   * @param aMsg
   *        The object to log. This is converted to a string.
   * @param aThrowable
   *        An exception to be logged
   */
  @Override
  protected void _log (final int nLevel, @Nonnull final Object aMsg, @Nullable final Throwable aThrowable)
  {
    // >= 7
    if (nLevel >= ERROR)
    {
      if (m_aLogger.isErrorEnabled ())
      {
        if (aMsg != null)
          m_aLogger.error ("{}", aMsg, aThrowable);
        else
          m_aLogger.error (BLA, aThrowable);
      }
    }
    else
      // >= 5
      if (nLevel >= WARN)
      {
        if (m_aLogger.isWarnEnabled ())
        {
          if (aMsg != null)
            m_aLogger.warn ("{}", aMsg, aThrowable);
          else
            m_aLogger.warn (BLA, aThrowable);
        }
      }
      else
        // >= 3
        if (nLevel >= INFO)
        {
          if (m_aLogger.isInfoEnabled ())
          {
            if (aMsg != null)
              m_aLogger.info ("{}", aMsg, aThrowable);
            else
              m_aLogger.info (BLA, aThrowable);
          }
        }
        else
          // >= 1
          if (nLevel >= DEBUG)
          {
            if (m_aLogger.isDebugEnabled ())
            {
              if (aMsg != null)
                m_aLogger.debug ("{}", aMsg, aThrowable);
              else
                m_aLogger.debug (BLA, aThrowable);
            }
          }
          else
          {
            // < 1
            if (m_aLogger.isTraceEnabled ())
            {
              if (aMsg != null)
                m_aLogger.trace ("{}", aMsg, aThrowable);
              else
                m_aLogger.trace (BLA, aThrowable);
            }
          }
  }

  /**
   * Check if a logger is enabled to log at the specified level
   *
   * @param nLevel
   *        One of DEBUG, INFO, WARN, ERROR, FATAL
   * @return <code>true</code> if the logger can handle the specified error
   *         level
   */
  @Override
  public boolean check (final int nLevel)
  {
    if (nLevel == FATAL || nLevel == ERROR)
      return m_aLogger.isErrorEnabled ();
    if (nLevel == WARN)
      return m_aLogger.isWarnEnabled ();
    if (nLevel == INFO)
      return m_aLogger.isInfoEnabled ();
    if (nLevel == DEBUG)
      return m_aLogger.isDebugEnabled ();
    return m_aLogger.isTraceEnabled ();
  }
}
