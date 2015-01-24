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

use Aaugustyniak\PhpExcelHandler\Excel\ActionCommand\FormatDataCommand;
use Aaugustyniak\PhpExcelHandler\Excel\Color;
use Aaugustyniak\PhpExcelHandler\Navigator\MatrixWalker;
use Aaugustyniak\PhpExcelHandler\Navigator\WriteAnchorGuesser;
use \PHPExcel as PHPExcel;

/**
 * @author Artur Augustyniak <artur@aaugustyniak.pl>
 */
class FormatTabularCommand extends MatrixWalker implements FormatDataCommand
{

    /**
     * @param WriteAnchorGuesser $guesser
     */
    function __construct(WriteAnchorGuesser $guesser)
    {
        parent::__construct($guesser);
    }


    public function format(PHPExcel $pe)
    {
        $this->walk($pe);
    }

    public function actOnCell($row, $col, $val)
    {
        $styleArray = array(
            'fill' => array(
                'type' => \PHPExcel_Style_Fill::FILL_SOLID,
                'color' => array('rgb' => $this->getColorFor($row))
            ),
            'borders' => array(
                'top' => array(
                    'style' => \PHPExcel_Style_Border::BORDER_THIN
                ),
                'bottom' => array(
                    'style' => \PHPExcel_Style_Border::BORDER_THIN
                ),
                'left' => array(
                    'style' => \PHPExcel_Style_Border::BORDER_THIN
                ),
                'right' => array(
                    'style' => \PHPExcel_Style_Border::BORDER_THIN
                ),
            )
        );
        $address = $this->navigator->getAddressFor($col, $row);
        $this->excel->getActiveSheet()->getStyle($address)->applyFromArray($styleArray);

    }

    /**
     * Color for odd-even rows and header
     * @param $row
     * @return mixed
     */
    public function getColorFor($row)
    {
        if (0 != $row && 0 != $row % 2) {
            return Color::rgb(Color::LightBlue);
        } elseif (0 == $row) {
            return Color::rgb(Color::LightGreen);
        } else {
            return Color::rgb(Color::LightGray);
        }
    }
}