<?php

namespace Navigators;

/**
 * Description of CoordinateProviderTest
 *
 * @author aaugustyniak
 */
class CellNavigatorTest extends \PHPUnit_Framework_TestCase
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
            $actualValues[] = $this->navigator->getRow($y);
        }
        $expectedValues = \range(1, 12);
        $this->assertEquals($actualValues, $expectedValues);
    }

    /**
     * @expectedException        \Exception
     * @expectedExceptionMessage Row out of bounds.
     */
    public function testNegativeCartesianY()
    {
        $this->navigator->getRow(-1);
    }

    /**
     * @expectedException        \Exception
     * @expectedExceptionMessage Row out of bounds.
     */
    public function testExcelOverflowCartesianY()
    {
        $this->navigator->getRow(CellNavigator::MAX_EXCEL_2007_ROWS);
    }

    public function testCartesianToExcelColumnsIDXSequence()
    {
        $naturalRange = \range(0, CellNavigator::MAX_EXCEL_2007_COLUMNS - 1);
        $actualValues = array();
        foreach ($naturalRange as $y) {
            $actualValues[] = $this->navigator->getCol($y);
        }
        $expectedValues = $this->generateExcelColumnsIDXs(0, CellNavigator::MAX_EXCEL_2007_COLUMNS - 1);
        $this->assertEquals($actualValues, $expectedValues);
    }

    /**
     * @expectedException        \Exception
     * @expectedExceptionMessage Column out of bounds.
     */
    public function testNegativeCartesianX()
    {
        $this->navigator->getCol(-1);
    }

    /**
     * @expectedException        \Exception
     * @expectedExceptionMessage Column out of bounds.
     */
    public function testExcelOverflowCartesianX()
    {
        $this->navigator->getCol(CellNavigator::MAX_EXCEL_2007_COLUMNS);
    }

    public function testCartesianToExcelCoordinatesSequence()
    {
        $expectedSequence = array();
        $actualSequence = array();
        for ($column = 0; $column < CellNavigator::MAX_EXCEL_2007_COLUMNS; $column++) {
            $expectedSequence[] = $this->numberToColumnName($column) . 1;
            $actualSequence[] = $this->navigator->getAddrFor($column, 0);
        }
        $this->assertEquals($expectedSequence, $actualSequence);
    }

    /**
     * @expectedException        \Exception
     * @expectedExceptionMessage Column out of bounds.
     */
    public function testExcelOverflowCartesianXInAddr()
    {
        $this->navigator->getAddrFor(CellNavigator::MAX_EXCEL_2007_COLUMNS, 0);
    }

    /**
     * @expectedException        \Exception
     * @expectedExceptionMessage Row out of bounds.
     */
    public function testExcelOverflowCartesianYInAddr()
    {
        $this->navigator->getAddrFor(0, CellNavigator::MAX_EXCEL_2007_ROWS);
    }

    private function generateExcelColumnsIDXs($from, $to)
    {
        $range = array();
        for ($index = $from; $index <= $to; $index++) {
            $range[] = $this->numberToColumnName($index);
        }
        return $range;
    }

    function numberToColumnName($number)
    {
        $number+=1;
        $alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        $abcLength = strlen($alphabet);
        $result = "";

        if ($number == 27) {
            return 'AA';
        }
        if ($number > 27) {
            $number-=1;
        }

        while ($number > $abcLength) {
            $remainder = $number % $abcLength;

            $result = $alphabet[$remainder] . $result;
            $number = floor($number / $abcLength);
        }
        return $alphabet[$number - 1] . $result;
    }

}
