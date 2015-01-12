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
namespace Aaugustyniak\PhpExcelHandler\Tests\Navigators;

use \PHPUnit_Framework_TestCase as TestCase;
use Aaugustyniak\PhpExcelHandler\Navigators\CellNavigator;

/**
 * @author Artur Augustyniak <artur@aaugustyniak.pl>
 * @package Aaugustyniak\PhpExcelHandler\Tests\Navigators
 */
class CellNavigatorTest extends TestCase
{

    private $navigator;

    protected function setUp()
    {
        $this->navigator = new CellNavigator();
    }

    public function testCartesianToExcelRowsIDXSequence()
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
     * @expectedException Aaugustyniak\PhpExcelHandler\Navigators\ExcelLimitException
     */
    public function testNegativeCartesianY()
    {
        $this->navigator->getRowAddressFor(-1);
    }

    /**
     * @expectedException Aaugustyniak\PhpExcelHandler\Navigators\ExcelLimitException
     */
    public function testExcelOverflowCartesianY()
    {
        $this->navigator->getRowAddressFor(CellNavigator::MAX_EXCEL_2007_ROWS);
    }

    public function testCartesianToExcelColumnsIDXSequence()
    {
        $naturalRange = \range(0, CellNavigator::MAX_EXCEL_2007_COLUMNS - 1);
        $actualValues = array();

        foreach ($naturalRange as $y) {
            $actualValues[] = $this->navigator->getColumnAddressFor($y);
        }
        $expectedValues = $this->generateExcelColumnsIDXs(0, CellNavigator::MAX_EXCEL_2007_COLUMNS - 1);
        $this->assertEquals($actualValues, $expectedValues);
    }

    /**
     * @expectedException Aaugustyniak\PhpExcelHandler\Navigators\ExcelLimitException
     */
    public function testNegativeCartesianX()
    {
        $this->navigator->getColumnAddressFor(-1);
    }

    /**
     * @expectedException Aaugustyniak\PhpExcelHandler\Navigators\ExcelLimitException
     */
    public function testExcelOverflowCartesianX()
    {
        $this->navigator->getColumnAddressFor(CellNavigator::MAX_EXCEL_2007_COLUMNS);
    }

    public function testCartesianToExcelCoordinatesSequence()
    {
        $expectedSequence = array();
        $actualSequence = array();
        for ($column = 0; $column < CellNavigator::MAX_EXCEL_2007_COLUMNS; $column++) {
            $expectedSequence[] = $this->numberToColumnName($column) . 1;
            $actualSequence[] = $this->navigator->getAddressFor($column, 0);
        }
        $this->assertEquals($expectedSequence, $actualSequence);
    }

    /**
     * @expectedException Aaugustyniak\PhpExcelHandler\Navigators\ExcelLimitException
     */
    public function testExcelOverflowCartesianXInAddr()
    {
        $this->navigator->getAddressFor(CellNavigator::MAX_EXCEL_2007_COLUMNS, 0);
    }

    /**
     * @expectedException Aaugustyniak\PhpExcelHandler\Navigators\ExcelLimitException
     */
    public function testExcelOverflowCartesianYInAddr()
    {
        $this->navigator->getAddressFor(0, CellNavigator::MAX_EXCEL_2007_ROWS);
    }

    /**
     * Generates sequence of proper excel column addresses
     *
     * @param $from
     * @param $to
     * @return array
     */
    private function generateExcelColumnsIDXs($from, $to)
    {
        $range = array();
        for ($index = $from; $index <= $to; $index++) {
            $range[] = $this->numberToColumnName($index);
        }
        return $range;
    }

    /**
     * Convert column index to excel name
     *
     * @param $number
     * @return string
     */
    function numberToColumnName($number)
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
