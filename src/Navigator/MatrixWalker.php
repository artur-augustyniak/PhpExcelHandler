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

use \PHPExcel as PHPExcel;

/**
 * @author Artur Augustyniak <artur@aaugustyniak.pl>
 */
abstract class MatrixWalker
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

    /**
     * @param WriteAnchorGuesser $guesser
     */
    function __construct(WriteAnchorGuesser $guesser)
    {
        $this->guesser = $guesser;
        $this->navigator = $guesser->getNavigator();
        $this->payload = $guesser->getPayload();
    }


    /**
     * @param PHPExcel $pe
     */
    public function walk(PHPExcel $pe)
    {
        $this->excel = $pe;
        $baseColIdx = $this->guesser->getColumnFromPayload();
        $rowIdx = $this->guesser->getRowFromPayload();
        foreach ($this->payload as $row) {
            if (is_array($row)) {
                $colIdx = $baseColIdx;
                foreach ($row as $cellValue) {
                    $this->actOnCell($rowIdx, $colIdx, $cellValue);
                    $colIdx++;
                }
            } else {
                $this->actOnCell($rowIdx, 0, $row);
            }
            $rowIdx++;
        }
    }

    /**
     * User defined work on cell
     *
     * @param $row
     * @param $col
     * @param $val
     * @return mixed
     */
    abstract public function actOnCell($row, $col, $val);

}