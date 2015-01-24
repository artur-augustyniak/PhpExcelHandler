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


use Aaugustyniak\PhpExcelHandler\Excel\ElementFactory\DefaultPhpExcelFactory;
use Aaugustyniak\PhpExcelHandler\Excel\SpreadSheet;
use Aaugustyniak\PhpExcelHandler\Navigator\CellNavigator;
use Aaugustyniak\PhpExcelHandler\Navigator\WriteAnchorGuesser;
use \PHPUnit_Framework_TestCase as TestCase;

/**
 * @author Artur Augustyniak <artur@aaugustyniak.pl>
 */
class WriteAnchorGuesserTest extends TestCase
{
    /**
     * @var SpreadSheet
     */
    private $spreadSheet;


    /**
     * @var CellNavigator
     */
    private $testNavigator;

    /**
     * @var WriteAnchorGuesser
     */
    private $guesser;


    protected function setUp()
    {
        $phpExcelFactory = new DefaultPhpExcelFactory();
        $this->spreadSheet = new SpreadSheet($phpExcelFactory);
        $this->testNavigator = new CellNavigator();
        $this->guesser = new WriteAnchorGuesser();
    }

    public function testRecognizeWriteAnchorEmptyArray()
    {
        $data = array();
        $this->guesser->setPayload($data);
        $expectedAnchor = $this->testNavigator->getAddressFor(0, 0);
        $actualAnchor = $this->guesser->getWriteAnchor();
        $this->assertEquals($expectedAnchor, $actualAnchor);
    }


    public function testRecognizeWriteAnchor1DArray()
    {
        $col = 0;
        for ($i = 0; $i < 10; $i++) {
            $data = array($col => 1, 2, 34);
            $this->guesser->setPayload($data);
            $expectedAnchor = $this->testNavigator->getAddressFor($col++, 0);
            $actualAnchor = $this->guesser->getWriteAnchor();
            $this->assertEquals($expectedAnchor, $actualAnchor);
        }
    }


    public function testRecognizeWriteAnchor1DHashArray()
    {
        $data = array("some-key" => 1, 2, 34);
        $this->guesser->setPayload($data);
        $expectedAnchor = $this->testNavigator->getAddressFor(0, 0);
        $actualAnchor = $this->guesser->getWriteAnchor();
        $this->assertEquals($expectedAnchor, $actualAnchor);

    }


    /**
     * diagonal
     */
    public function testRecognizeWriteAnchor2DArray()
    {
        $col = 0;
        $row = 0;
        for ($i = 0; $i < 10; $i++) {
            $data = array(
                $col => array($row => 1, 2, 34),
                1,
                array(1, 2, 3)
            );
            $this->guesser->setPayload($data);
            $expectedAnchor = $this->testNavigator->getAddressFor($col++, $row++);
            $actualAnchor = $this->guesser->getWriteAnchor();
            $this->assertEquals($expectedAnchor, $actualAnchor);
        }
    }

    /**
     * vertical
     */
    public function testRecognizeWriteAnchor2DHashArray()
    {
        $row = 0;
        for ($i = 0; $i < 10; $i++) {
            $data = array(
                "some-key" => array($row => 1, 2, 34),
                1,
                array(1, 2, 3)
            );
            $this->guesser->setPayload($data);
            $expectedAnchor = $this->testNavigator->getAddressFor(0, $row++);
            $actualAnchor = $this->guesser->getWriteAnchor();
            $this->assertEquals($expectedAnchor, $actualAnchor);
        }
    }


    /**
     * horizontal
     */
    public function testRecognizeWriteAnchor2DHashRowArray()
    {
        $col = 0;
        for ($i = 0; $i < 10; $i++) {
            $data = array(
                $col => array("some-key" => 1, 2, 34),
                1,
                array(1, 2, 3)
            );
            $this->guesser->setPayload($data);
            $expectedAnchor = $this->testNavigator->getAddressFor($col++, 0);
            $actualAnchor = $this->guesser->getWriteAnchor();
            $this->assertEquals($expectedAnchor, $actualAnchor);
        }
    }

}