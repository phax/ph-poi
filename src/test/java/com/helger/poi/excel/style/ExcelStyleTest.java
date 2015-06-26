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

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertNull;
import static org.junit.Assert.assertSame;
import static org.junit.Assert.assertTrue;

import org.apache.poi.ss.usermodel.IndexedColors;
import org.junit.Test;

import com.helger.commons.mock.CommonsTestHelper;

/**
 * Test class for class {@link ExcelStyle}.
 *
 * @author Philip Helger
 */
public final class ExcelStyleTest
{
  @Test
  public void testBasic ()
  {
    final ExcelStyle e = new ExcelStyle ();
    CommonsTestHelper.testDefaultImplementationWithEqualContentObject (e, new ExcelStyle ());
    e.setBorder (EExcelBorder.BORDER_DASH_DOT);
    CommonsTestHelper.testDefaultImplementationWithDifferentContentObject (e, new ExcelStyle ());
    CommonsTestHelper.testDefaultImplementationWithEqualContentObject (e,
                                                                       new ExcelStyle ().setBorder (EExcelBorder.BORDER_DASH_DOT));
  }

  @Test
  public void testAlign ()
  {
    final ExcelStyle e = new ExcelStyle ();
    assertNull (e.getAlign ());
    CommonsTestHelper.testGetClone (e);
    for (final EExcelAlignment eAlign : EExcelAlignment.values ())
    {
      assertSame (e, e.setAlign (eAlign));
      assertEquals (eAlign, e.getAlign ());
      CommonsTestHelper.testGetClone (e);
    }
  }

  @Test
  public void testVerticalAlign ()
  {
    final ExcelStyle e = new ExcelStyle ();
    assertNull (e.getVerticalAlign ());
    CommonsTestHelper.testGetClone (e);
    for (final EExcelVerticalAlignment eAlign : EExcelVerticalAlignment.values ())
    {
      assertSame (e, e.setVerticalAlign (eAlign));
      assertEquals (eAlign, e.getVerticalAlign ());
      CommonsTestHelper.testGetClone (e);
    }
  }

  @Test
  public void testWrapText ()
  {
    final ExcelStyle e = new ExcelStyle ();
    assertTrue (e.isWrapText () == ExcelStyle.DEFAULT_WRAP_TEXT);
    CommonsTestHelper.testGetClone (e);
    assertSame (e, e.setWrapText (true));
    assertTrue (e.isWrapText ());
    CommonsTestHelper.testGetClone (e);
    assertSame (e, e.setWrapText (false));
    assertFalse (e.isWrapText ());
    CommonsTestHelper.testGetClone (e);
  }

  @Test
  public void testDataFormat ()
  {
    final ExcelStyle e = new ExcelStyle ();
    assertNull (e.getDataFormat ());
    CommonsTestHelper.testGetClone (e);
    assertSame (e, e.setDataFormat ("abc"));
    assertEquals ("abc", e.getDataFormat ());
    CommonsTestHelper.testGetClone (e);
    assertSame (e, e.setDataFormat (null));
    assertNull (e.getDataFormat ());
    CommonsTestHelper.testGetClone (e);
  }

  @Test
  public void testFillBackgroundColor ()
  {
    final ExcelStyle e = new ExcelStyle ();
    assertNull (e.getFillBackgroundColor ());
    CommonsTestHelper.testGetClone (e);
    for (final IndexedColors eColor : IndexedColors.values ())
    {
      assertSame (e, e.setFillBackgroundColor (eColor));
      assertEquals (eColor, e.getFillBackgroundColor ());
      CommonsTestHelper.testGetClone (e);
    }
  }

  @Test
  public void testFillForegroundColor ()
  {
    final ExcelStyle e = new ExcelStyle ();
    assertNull (e.getFillForegroundColor ());
    CommonsTestHelper.testGetClone (e);
    for (final IndexedColors eColor : IndexedColors.values ())
    {
      assertSame (e, e.setFillForegroundColor (eColor));
      assertEquals (eColor, e.getFillForegroundColor ());
      CommonsTestHelper.testGetClone (e);
    }
  }

  @Test
  public void testFillPattern ()
  {
    final ExcelStyle e = new ExcelStyle ();
    assertNull (e.getFillPattern ());
    CommonsTestHelper.testGetClone (e);
    for (final EExcelPattern ePattern : EExcelPattern.values ())
    {
      assertSame (e, e.setFillPattern (ePattern));
      assertEquals (ePattern, e.getFillPattern ());
      CommonsTestHelper.testGetClone (e);
    }
  }

  @Test
  public void testBorder ()
  {
    final ExcelStyle e = new ExcelStyle ();
    assertNull (e.getBorderTop ());
    assertNull (e.getBorderRight ());
    assertNull (e.getBorderBottom ());
    assertNull (e.getBorderLeft ());
    CommonsTestHelper.testGetClone (e);
    for (final EExcelBorder eBorder : EExcelBorder.values ())
    {
      assertSame (e, e.setBorder (eBorder));
      assertEquals (eBorder, e.getBorderTop ());
      assertEquals (eBorder, e.getBorderRight ());
      assertEquals (eBorder, e.getBorderBottom ());
      assertEquals (eBorder, e.getBorderLeft ());
      CommonsTestHelper.testGetClone (e);
    }
  }
}
