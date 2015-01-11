<?php

namespace Aaugustyniak\PhpExcelHandler\Navigators;

/**
 * Description of CellNavigator
 *
 * @author aaugustyniak
 */
class CellNavigator
{

    /**
     * @see http://office.microsoft.com/en-us/excel-help/excel-specifications-and-limits-HP010073849.aspx
     */
    const MAX_EXCEL_2007_COLUMNS = 16384;
    const MAX_EXCEL_2007_ROWS = 1048576;

    public function getAddrFor($column, $row)
    {
        return $this->getCol($column) . $this->getRow($row);
    }

    public function getRow($y)
    {
        $this->validateRowYCoord($y);
        return $y + 1;
    }

    public function getCol($x)
    {
        $this->validateColumnXCoord($x);
        $x+=1;
        $abc = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        $abc_len = strlen($abc);
        $result = "";

        if ($x == 27) {
            return 'AA';
        }
        if ($x > 27) {
            $x-=1;
        }

        while ($x > $abc_len) {
            $remainder = $x % $abc_len;
            $result = $abc[$remainder] . $result;
            $x = floor($x / $abc_len);
        }
        return $abc[$x - 1] . $result;
    }

    private function validateRowYCoord($y)
    {
        if ($y < 0 || $y > self::MAX_EXCEL_2007_ROWS - 1) {
            throw new \Exception('Row out of bounds.');
        }
    }

    private function validateColumnXCoord($x)
    {
        if ($x < 0 || $x > self::MAX_EXCEL_2007_COLUMNS - 1) {
            throw new \Exception('Column out of bounds.');
        }
    }

}
