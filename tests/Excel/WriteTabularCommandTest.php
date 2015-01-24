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


use Aaugustyniak\PhpExcelHandler\Excel\ActionCommand\Basic\WriteTabularCommand;
use Aaugustyniak\PhpExcelHandler\Excel\ElementFactory\DefaultPhpExcelFactory;
use Aaugustyniak\PhpExcelHandler\Excel\SpreadSheet;
use Aaugustyniak\PhpExcelHandler\Navigator\WriteAnchorGuesser;
use \PHPUnit_Framework_TestCase as TestCase;

/**
 * @author Artur Augustyniak <artur@aaugustyniak.pl>
 */
class WriteTabularCommandTest extends TestCase
{
    /**
     * @var SpreadSheet
     */
    private $spreadSheet;

    private $anchorGuesser;

    /**
     * @var string
     */
    private $systemTmp;


    protected function setUp()
    {
        $this->systemTmp = \sys_get_temp_dir();
        $phpExcelFactory = new DefaultPhpExcelFactory();
        $this->spreadSheet = new SpreadSheet($phpExcelFactory);
        $data = array(
            array("Column1", "Column2", "Column3"),
            array(1, 2, 3),
            array(4, 5, 6),
            array(7, 8, 9),
        );
        $this->anchorGuesser = new WriteAnchorGuesser($data);
        $this->anchorGuesser->forceFixIndexing();
    }

    public function testMatrixDump()
    {

        $writer = new WriteTabularCommand($this->anchorGuesser);
        $this->spreadSheet->modify($writer);
        $this->spreadSheet->saveHtml(__DIR__ . '/test.html');
    }


}