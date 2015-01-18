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
namespace Aaugustyniak\PhpExcelHandler\Tests\Navigator;

use Aaugustyniak\PhpExcelHandler\Navigator\ExcelLimits;
use \PHPUnit_Framework_TestCase as TestCase;
use Aaugustyniak\PhpExcelHandler\Navigator\CellNavigator;

/**
 * @author Artur Augustyniak <artur@aaugustyniak.pl>
 */
class CellNavigatorTest extends TestCase
{

    private $navigator;

    protected function setUp()
    {
        $this->navigator = new CellNavigator();
    }

    public function testIndicesToExcelRowsAddressesSequence()
    {
        $naturalRange = \range(0, 11);
        $actualValues = array();
        foreach ($naturalRange as $y) {
            $actualValues[] = $this->navigator->getRowAddressFor($y);
        }
        $expectedValues = \range(1, 12);
        $this->assertEquals($actualValues, $expectedValues);
    }

    /**
     * @expectedException \Aaugustyniak\PhpExcelHandler\Navigator\ExcelLimits
     */
    public function testNegativeRow()
    {
        $this->navigator->getRowAddressFor(-1);
    }

    /**
     * @expectedException \Aaugustyniak\PhpExcelHandler\Navigator\ExcelLimits
     */
    public function testExcelRowOverflow()
    {
        $this->navigator->getRowAddressFor(ExcelLimits::MAX_EXCEL_2007_ROW+1);
    }

    public function testIndicesToExcelColumnsAddressesSequence()
    {
        $naturalRange = \range(0, ExcelLimits::MAX_EXCEL_2007_COLUMN);
        $actualValues = array();

        foreach ($naturalRange as $y) {
            $actualValues[] = $this->navigator->getColumnAddressFor($y);
        }
        $expectedValues = $this->generateExcelColumnAddresses(0, ExcelLimits::MAX_EXCEL_2007_COLUMN);
        $this->assertEquals($actualValues, $expectedValues);
    }

    /**
     * @expectedException \Aaugustyniak\PhpExcelHandler\Navigator\ExcelLimits
     */
    public function testNegativeColumn()
    {
        $this->navigator->getColumnAddressFor(-1);
    }

    /**
     * @expectedException \Aaugustyniak\PhpExcelHandler\Navigator\ExcelLimits
     */
    public function testExcelColumnOverflow()
    {
        $this->navigator->getColumnAddressFor(ExcelLimits::MAX_EXCEL_2007_COLUMN+1);
    }

    public function testIndicesToExcelCoordinatesSequence()
    {
        $expectedSequence = array();
        $actualSequence = array();
        for ($column = 0; $column <= ExcelLimits::MAX_EXCEL_2007_COLUMN; $column++) {
            $expectedSequence[] = $this->numberToColumnName($column) . 1;
            $actualSequence[] = $this->navigator->getAddressFor($column, 0);
        }
        $this->assertEquals($expectedSequence, $actualSequence);
    }

    /**
     * @expectedException \Aaugustyniak\PhpExcelHandler\Navigator\ExcelLimits
     */
    public function testExcelColumnOverflowInAddress()
    {
        $this->navigator->getAddressFor(ExcelLimits::MAX_EXCEL_2007_COLUMN+1, 0);
    }

    /**
     * @expectedException \Aaugustyniak\PhpExcelHandler\Navigator\ExcelLimits
     */
    public function testExcelRowOverflowInAddress()
    {
        $this->navigator->getAddressFor(0, ExcelLimits::MAX_EXCEL_2007_ROW+1);
    }

    /**
     * Generates sequence of proper excel column addresses
     *
     * @param $from
     * @param $to
     * @return array
     */
    private function generateExcelColumnAddresses($from, $to)
    {
        $range = array();
        for ($index = $from; $index <= $to; $index++) {
            $range[] = $this->numberToColumnName($index);
        }
        return $range;
    }

    /**
     * Convert column index to excel name
     * Proper solution for testing eventually more efficient implementations
     *
     * @param $number
     * @return string
     */
    private function numberToColumnName($number)
    {

        $number += 1;
        $alphabet = CellNavigator::ALPHABET;
        $abcLength = strlen($alphabet);
        $result = "";

        if ($number == 27) {
            return 'AA';
        }
        if ($number > 27) {
            $number -= 1;
        }

        while ($number > $abcLength) {
            $remainder = $number % $abcLength;

            $result = $alphabet[$remainder] . $result;
            $number = floor($number / $abcLength);
        }
        return $alphabet[(int)$number - 1] . $result;
    }

}
