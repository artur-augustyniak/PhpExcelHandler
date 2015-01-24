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

use Aaugustyniak\PhpExcelHandler\Excel\ActionCommand\ReadDataCommand;
use Aaugustyniak\PhpExcelHandler\Navigator\MatrixWalker;
use Aaugustyniak\PhpExcelHandler\Navigator\WriteAnchorGuesser;
use \PHPExcel as PHPExcel;


/**
 * @author Artur Augustyniak <artur@aaugustyniak.pl>
 */
class ReadTabularCommand extends MatrixWalker implements ReadDataCommand
{

    private $data = array();

    /**
     * @param WriteAnchorGuesser $guesser
     */
    function __construct(WriteAnchorGuesser $guesser)
    {
        parent::__construct($guesser);
    }

    /**
     * @param $row
     * @param $col
     * @param $val
     * @throws \PHPExcel_Exception
     */
    public function actOnCell($row, $col, $val)
    {
        $cellAddress = $this->navigator->getAddressFor($col, $row);
        $cellValue = $this->excel->getActiveSheet()->getCell($cellAddress)->getValue();
        $this->data[$row][$col] = $cellValue;
    }

    /**
     *
     * @return array
     */
    public function readFrom(PHPExcel $pe)
    {
        $this->walk($pe);
    }

    /**
     * @return mixed
     */
    public function fetchData()
    {
        return $this->data;
    }
}