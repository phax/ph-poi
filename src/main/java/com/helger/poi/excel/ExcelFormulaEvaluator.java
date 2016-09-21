/**
 * Copyright (C) 2014-2016 Philip Helger (www.helger.com)
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

import javax.annotation.Nonnull;
import javax.annotation.Nullable;

import org.apache.poi.hssf.usermodel.HSSFFormulaEvaluator;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.formula.IStabilityClassifier;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.helger.commons.string.ToStringGenerator;

@SuppressWarnings ("deprecation")
public class ExcelFormulaEvaluator
{
  private final FormulaEvaluator m_aEvaluator;

  public ExcelFormulaEvaluator (@Nonnull final Workbook aWB)
  {
    m_aEvaluator = aWB.getCreationHelper ().createFormulaEvaluator ();
  }

  public ExcelFormulaEvaluator (@Nonnull final Workbook aWB, @Nullable final IStabilityClassifier aStability)
  {
    m_aEvaluator = aWB instanceof HSSFWorkbook ? new HSSFFormulaEvaluator ((HSSFWorkbook) aWB, aStability)
                                               : XSSFFormulaEvaluator.create ((XSSFWorkbook) aWB, aStability, null);
  }

  /**
   * If cell contains a formula, the formula is evaluated and returned, else the
   * CellValue simply copies the appropriate cell value from the cell and also
   * its cell type. This method should be preferred over evaluateInCell() when
   * the call should not modify the contents of the original cell.
   *
   * @param aCell
   *        The cell to evaluate
   * @return The evaluation result
   */
  public CellValue evaluate (@Nonnull final Cell aCell)
  {
    return m_aEvaluator.evaluate (aCell);
  }

  /**
   * If cell contains formula, it evaluates the formula, and saves the result of
   * the formula. The cell remains as a formula cell. Else if cell does not
   * contain formula, this method leaves the cell unchanged. Note that the type
   * of the formula result is returned, so you know what kind of value is also
   * stored with the formula.
   *
   * <pre>
   * int evaluatedCellType = evaluator.evaluateFormulaCell (cell);
   * </pre>
   *
   * Be aware that your cell will hold both the formula, and the result. If you
   * want the cell replaced with the result of the formula, use
   * {@link #evaluateInCell(Cell)}
   *
   * @param aCell
   *        The cell to evaluate
   * @return The type of the formula result (the cell's type remains as
   *         Cell.CELL_TYPE_FORMULA however)
   */
  public int evaluateFormulaCell (@Nonnull final Cell aCell)
  {
    return m_aEvaluator.evaluateFormulaCell (aCell);
  }

  /**
   * If cell contains formula, it evaluates the formula, and puts the formula
   * result back into the cell, in place of the old formula. Else if cell does
   * not contain formula, this method leaves the cell unchanged. Note that the
   * same instance of Cell is returned to allow chained calls like:
   *
   * <pre>
   * int evaluatedCellType = evaluator.evaluateInCell (cell).getCellType ();
   * </pre>
   *
   * Be aware that your cell value will be changed to hold the result of the
   * formula. If you simply want the formula value computed for you, use
   * {@link #evaluateFormulaCell(Cell)}
   *
   * @param aCell
   *        Cell to evaluate
   * @return The cell in which it was evaluated
   */
  @Nonnull
  public Cell evaluateInCell (@Nonnull final Cell aCell)
  {
    return m_aEvaluator.evaluateInCell (aCell);
  }

  @Override
  public String toString ()
  {
    return new ToStringGenerator (this).append ("evaluator", m_aEvaluator).toString ();
  }
}
