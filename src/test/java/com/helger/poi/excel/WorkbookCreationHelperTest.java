/**
 * Copyright (C) 2014-2018 Philip Helger (www.helger.com)
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

import java.io.File;

import org.junit.Test;

/**
 * Test class for class {@link WorkbookCreationHelper}.
 *
 * @author Philip Helger
 */
public final class WorkbookCreationHelperTest
{
  @Test
  public void testAddCellFormula ()
  {
    final WorkbookCreationHelper aWBCH = new WorkbookCreationHelper (EExcelVersion.XLSX);
    aWBCH.createNewSheet ("Test sheet1");
    aWBCH.addRow ();
    // Since 3.14 invalid formulas can be set without Exception
    aWBCH.addCellFormula ("ABC(A1)");

    // Test merging
    {
      aWBCH.addRow ();
      aWBCH.addCell ("Col1");
      aWBCH.addCell ("Col2");
      aWBCH.addCell ("Col3");
      aWBCH.addCell ("Col4");
      aWBCH.addCell ("Col5");

      // Keep "Col1"
      aWBCH.addMergeRegionInCurrentRow (0, 1);

      // Keep "Col3"
      aWBCH.addMergeRegionInCurrentRow (2, 4);
    }

    // Write to dummy file
    aWBCH.writeTo (new File ("mock.xlsx"));
  }
}
