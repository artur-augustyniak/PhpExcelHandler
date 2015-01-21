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

namespace Aaugustyniak\PhpExcelHandler\Excel\ActionCommand;

use \PHPExcel as PHPExcel;

/**
 * @author Artur Augustyniak <artur@aaugustyniak.pl>
 */
interface ReadDataCommand
{
    /**
     *
     * @return array
     */
    public function readFrom(PHPExcel $pe);

    /**
     * @return mixed
     */
    public function fetchData();


}


//class ArrayXlsxEditor extends XlsxEditor {
//
//    public function writeFromDataSource() {
//        $i = 1;
//        foreach ($this->dataSource as $k => $v) {
//            $keyCellCoord = 'A' . $i;
//            $varCellCoord = 'B' . $i;
//            $this->activeSheet->setCellValue($keyCellCoord, $k);
//            $this->activeSheet->setCellValue($varCellCoord, $v);
//            $i++;
//        }
//    }
//
//}