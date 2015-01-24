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

namespace Aaugustyniak\PhpExcelHandler\Excel\ActionCommand\Basic;

use Aaugustyniak\PhpExcelHandler\Excel\ActionCommand\ModifyDataCommand;
use Aaugustyniak\PhpExcelHandler\Navigator\CellNavigator;
use Aaugustyniak\PhpExcelHandler\Navigator\WriteAnchorGuesser;
use \PHPExcel as PHPExcel;


/**
 * @author Artur Augustyniak <artur@aaugustyniak.pl>
 */
class WriteTabularCommand implements ModifyDataCommand
{

    /**
     * @var array
     */
    protected $payload;

    /**
     * @var WriteAnchorGuesser
     */
    protected $guesser;

    /**
     * @var PHPExcel
     */
    protected $excel;

    /**
     * @var CellNavigator
     */
    protected $navigator;

    function __construct(WriteAnchorGuesser $guesser)
    {
        $this->guesser = $guesser;
        $this->navigator = $guesser->getNavigator();
        $this->payload = $guesser->getPayload();
    }


    /**
     * @return void
     */
    public function modify(PHPExcel $pe)
    {
        $this->excel = $pe;
        $baseColIdx = $this->guesser->getColumnFromPayload();
        $rowIdx = $this->guesser->getRowFromPayload();
        foreach ($this->payload as $row) {
            if (is_array($row)) {
                $colIdx = $baseColIdx;
                foreach ($row as $cellValue) {
                    $this->writeCell($rowIdx, $colIdx, $cellValue);
                    $colIdx++;
                }
            } else {
                $this->writeCell($rowIdx, 0, $row);
            }
            $rowIdx++;
        }
    }


    protected function writeCell($row, $col, $val)
    {
        $cellAddress = $this->navigator->getAddressFor($col, $row);
        $this->excel->getActiveSheet()->setCellValue($cellAddress, $val);

    }


}