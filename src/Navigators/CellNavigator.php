<?php
/*
* The MIT License (MIT)
*
* Copyright (c) 2014 Artur Augustyniak
*
* Permission is hereby granted, free of charge, to any person obtaining a copy
* of this software and associated documentation files (the "Software"), to deal
* in the Software without restriction, including without limitation the rights
* to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
* copies of the Software, and to permit persons to whom the Software is
* furnished to do so, subject to the following conditions:
*
* The above copyright notice and this permission notice shall be included in
* all copies or substantial portions of the Software.
*
* THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
* IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
* FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
* AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
* LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
* OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
*/
namespace Aaugustyniak\PhpExcelHandler\Navigators;


/**
 * Navigation abstraction for translate integer, zero based indices to
 * Excel cell addresses.
 *
 * @author Artur Augustyniak <artur@aaugustyniak.pl>
 * @package Aaugustyniak\PhpExcelHandler\Navigators
 */
class CellNavigator
{

    /**
     * Excel limits
     * @see http://office.microsoft.com/en-us/excel-help/excel-specifications-and-limits-HP010073849.aspx
     */
    const MAX_EXCEL_2007_COLUMNS = 16384;
    const MAX_EXCEL_2007_ROWS = 1048576;

    const ALPHABET = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";


    /**
     * Translate indices with start at top left corner
     * to excel cell address
     *
     * @param $columnIdx
     * @param $rowIdx
     * @return string
     */
    public function getAddressFor($columnIdx, $rowIdx)
    {
        return $this->getColumnAddressFor($columnIdx) . $this->getRowAddressFor($rowIdx);
    }

    /**
     * Translate row indices with start at top
     * to excel row address
     *
     * @param $rowIdx
     * @return string
     * @throws ExcelLimitException
     */
    public function getRowAddressFor($rowIdx)
    {
        $this->validateRowIdx($rowIdx);
        return (string)$rowIdx + 1;
    }

    /**
     * Translate column indices with start at left side
     * to excel column address
     *
     * @param $columnIdx
     * @return string
     * @throws ExcelLimitException
     */
    public function getColumnAddressFor($columnIdx)
    {
        $this->validateColumnIdx($columnIdx);
        $columnNumericAddress = $columnIdx + 1;
        $alphabet = self::ALPHABET;
        $alphabetLength = strlen($alphabet);
        $result = "";

        if ($columnNumericAddress == 27) {
            return 'AA';
        }
        if ($columnNumericAddress > 27) {
            $columnNumericAddress -= 1;
        }

        while ($columnNumericAddress > $alphabetLength) {
            $remainder = $columnNumericAddress % $alphabetLength;
            $result = $alphabet[$remainder] . $result;
            $columnNumericAddress = floor($columnNumericAddress / $alphabetLength);
        }

        $partial = (string)$alphabet[(int)$columnNumericAddress - 1];
        return $partial . $result;
    }

    /**
     * Validate excel row limit
     * @see http://office.microsoft.com/en-us/excel-help/excel-specifications-and-limits-HP010073849.aspx
     *
     * @param $rowIdx
     * @throws ExcelLimitException
     */
    private function validateRowIdx($rowIdx)
    {
        if ($rowIdx < 0 || $rowIdx > self::MAX_EXCEL_2007_ROWS - 1) {
            $msg = sprintf('Row index [%s] out of bonds.', $rowIdx);
            throw new ExcelLimitException($msg);
        }
    }

    /**
     * Validate excel columns limit
     * @see http://office.microsoft.com/en-us/excel-help/excel-specifications-and-limits-HP010073849.aspx
     *
     * @param $colIdx
     * @throws ExcelLimitException
     */
    private function validateColumnIdx($colIdx)
    {
        if ($colIdx < 0 || $colIdx > self::MAX_EXCEL_2007_COLUMNS - 1) {
            $msg = sprintf('Column index [%s] out of bonds.', $colIdx);
            throw new ExcelLimitException();
        }
    }

}
