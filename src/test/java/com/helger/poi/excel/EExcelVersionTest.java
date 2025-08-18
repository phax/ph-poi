/*
 * Copyright (C) 2014-2025 Philip Helger (www.helger.com)
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

import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.assertNull;
import static org.junit.Assert.assertTrue;

import java.nio.charset.StandardCharsets;

import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Test;

import com.helger.base.io.nonblocking.NonBlockingByteArrayInputStream;
import com.helger.io.resource.ClassPathResource;
import com.helger.io.resource.IReadableResource;

/**
 * Test class for class {@link EExcelVersion}.
 *
 * @author Philip Helger
 */
public final class EExcelVersionTest
{
  @Test
  public void testSimple ()
  {
    for (final EExcelVersion eVersion : EExcelVersion.values ())
    {
      assertNotNull (eVersion.createWorkbook ());
      assertNotNull (eVersion.createRichText ("Hi"));
      assertNotNull (eVersion.getFileExtension ());
      assertTrue (eVersion.getFileExtension ().startsWith ("."));
      assertNotNull (eVersion.getMimeType ());
    }
  }

  @Test
  public void testReadWorkbook ()
  {
    final IReadableResource aXLSX = new ClassPathResource ("excel/test1.xlsx");
    assertTrue (aXLSX.exists ());
    Workbook aWB = EExcelVersion.XLSX.readWorkbook (aXLSX.getInputStream ());
    assertNotNull (aWB);
    aWB = EExcelVersion.XLS.readWorkbook (aXLSX.getInputStream ());
    assertNull (aWB);
    aWB = EExcelVersion.XLS.readWorkbook (new NonBlockingByteArrayInputStream ("abc".getBytes (StandardCharsets.ISO_8859_1)));
    assertNull (aWB);

    final IReadableResource aXLS = new ClassPathResource ("excel/test1.xls");
    assertTrue (aXLS.exists ());
    aWB = EExcelVersion.XLSX.readWorkbook (aXLS.getInputStream ());
    assertNull (aWB);
    aWB = EExcelVersion.XLS.readWorkbook (aXLS.getInputStream ());
    assertNotNull (aWB);
    aWB = EExcelVersion.XLSX.readWorkbook (new NonBlockingByteArrayInputStream ("abc".getBytes (StandardCharsets.ISO_8859_1)));
    assertNull (aWB);
  }
}
