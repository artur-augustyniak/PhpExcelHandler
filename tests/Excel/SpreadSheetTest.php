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
namespace Aaugustyniak\PhpExcelHandler\Tests\Excel;

use \PHPUnit_Framework_TestCase as TestCase;
use Aaugustyniak\PhpExcelHandler\Excel\PHPExcelElementFactory;
use Aaugustyniak\PhpExcelHandler\Excel\SpreadSheet;


class TestPhpExcelFactory implements PHPExcelElementFactory
{

    /**
     * @return \PHPExcel
     */
    public function newPHPExcelObject()
    {
        return new \PHPExcel();
    }

    /**
     * @return \PHPExcel_Worksheet
     */
    public function newPHPExcelWorkSheetObject()
    {
        return new \PHPExcel_Worksheet();
    }

    static public function getFactory()
    {
        return new self;
    }


}


/**
 * @author Artur Augustyniak <artur@aaugustyniak.pl>
 */
class SpreadSheetTest extends TestCase
{

    const DEF_WORKSHEET_NAME = 'Worksheet';

    /**
     * @var SpreadSheet
     */
    private $spreadSheet;

    protected function setUp()
    {
        $phpExcelFactory = TestPhpExcelFactory::getFactory();
        $this->spreadSheet = new SpreadSheet();
        $this->spreadSheet->setElementFactory($phpExcelFactory);
    }

    public function testCreatedSpreadSheetContainsSheetsWithTitles()
    {
        $titles = $this->spreadSheet->getSheetTiles();
        $currentActiveSheet = $this->spreadSheet->getActiveSheet();
        $this->assertTrue(is_array($titles), "Result is not an array!");
        $this->assertEquals(self::DEF_WORKSHEET_NAME, $titles[0]);
        $this->assertEquals(self::DEF_WORKSHEET_NAME, $currentActiveSheet);
    }

    public function testSpreadSheetContainsSheetsWithTitlesAfterSheetAddition()
    {
        $newWorksheetTitle = self::DEF_WORKSHEET_NAME . '_NEW';
        $this->spreadSheet->addSheet($newWorksheetTitle);
        $titles = $this->spreadSheet->getSheetTiles();
        $currentActiveSheet = $this->spreadSheet->getActiveSheet();
        $this->assertTrue(is_array($titles), "Result is not an array!");
        $this->assertEquals(self::DEF_WORKSHEET_NAME, $titles[0]);
        $this->assertEquals($newWorksheetTitle, $titles[1]);
        $this->assertEquals(self::DEF_WORKSHEET_NAME, $currentActiveSheet);
    }


    /**
     * @expectedException \Aaugustyniak\PhpExcelHandler\Excel\NoSuchElement
     */
    public function testCanNotSetActiveSheetToNonExistingOne()
    {
        $newWorksheet = self::DEF_WORKSHEET_NAME . '_NONEXISTENT';
        $this->spreadSheet->setActiveSheet($newWorksheet);

    }


    public function testActiveSheetChange()
    {
        $newWorksheetTitle = self::DEF_WORKSHEET_NAME . '_NEW_ACTIVE';
        $this->spreadSheet->addSheet($newWorksheetTitle);
        $titles = $this->spreadSheet->getSheetTiles();
        $this->spreadSheet->setActiveSheet($newWorksheetTitle);
        $currentActiveSheet = $this->spreadSheet->getActiveSheet();
        $this->assertTrue(is_array($titles), "Result is not an array!");
        $this->assertEquals(self::DEF_WORKSHEET_NAME, $titles[0]);
        $this->assertEquals($newWorksheetTitle, $titles[1]);
        $this->assertEquals($newWorksheetTitle, $currentActiveSheet);
    }


}
