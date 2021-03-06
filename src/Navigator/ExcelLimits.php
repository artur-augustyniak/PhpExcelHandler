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

namespace Aaugustyniak\PhpExcelHandler\Navigator;

use \Exception as Exception;

/**
 * ExcelLimits indicating coordinate out of excel bounds
 *
 * @author Artur Augustyniak <artur@aaugustyniak.pl>
 * @package Aaugustyniak\PhpExcelHandler\Navigator
 */
class ExcelLimits extends Exception
{

    /**
     * Excel limits
     * @see http://office.microsoft.com/en-us/excel-help/excel-specifications-and-limits-HP010073849.aspx
     */
    const MAX_EXCEL_2007_COLUMN = 16383; //16384 -1
    const MAX_EXCEL_2007_ROW = 1048575; //1048576 -1

    /**
     * Constructor with default values
     * @param string $message
     * @param int $code
     * @param Exception $previous
     */
    public function __construct($message = "Given index out of bonds", $code = 0, Exception $previous = null)
    {
        parent::__construct($message, $code, $previous);
    }


}