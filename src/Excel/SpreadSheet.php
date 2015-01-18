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

namespace Aaugustyniak\PhpExcelHandler\Excel;

use \PHPExcel as PHPExcel;
use \Exception as Exception;

/**
 * Excel spreadsheet abstraction
 * Facade for \PHPExcel
 *
 * @author Artur Augustyniak <artur@aaugustyniak.pl>
 */
class SpreadSheet
{
    const DEF_FILE_NAME = "Workbook";

    const EXCEL_EXT = "xlsx";

    const PDF_EXT = "pdf";

    const HTML_EXT = "html";

    /**
     * @var string
     */
    private $fileName;

    /**
     * @var PHPExcelElementFactory
     */
    private $elementFactory;

    /**
     * @var PHPExcel
     */
    private $excel;

    /**
     *
     * @param string $fileName
     */
    function __construct($fileName = self::DEF_FILE_NAME)
    {
        $this->fileName = $fileName;
        $this->elementFactory = DefaultPHPExcelElementFactory::getFactory();
        $this->excel = $this->elementFactory->newPHPExcelObject();
    }

    /**
     * @param PHPExcelElementFactory $elementFactory
     */
    public function setElementFactory(PHPExcelElementFactory $elementFactory)
    {
        $this->elementFactory = $elementFactory;
    }


    private function openFile($path)
    {
        throw new NoSuchElement();
    }

    /**
     * Get titles of sheets in current document
     * @return array
     */
    public function getSheetTiles()
    {
        $titles = array();
        foreach ($this->excel->getAllSheets() as $sheet) {
            $titles[] = $sheet->getTitle();
        }
        return $titles;
    }

    /**
     * Add sheet with title
     * @param string $title
     */
    public function addSheet($title)
    {
        $newWorkSheet = $this->elementFactory->newPHPExcelWorkSheetObject();
        $newWorkSheet->setTitle($title);
        $this->excel->addSheet($newWorkSheet);
    }

    /**
     * Set active sheet using name
     * @param string $title
     * @throws NoSuchElement
     */
    public function setActiveSheet($title)
    {
        try {
            $this->excel->setActiveSheetIndexByName($title);
        } catch (Exception $e) {
            throw new NoSuchElement("No such sheet", $e->getCode(), $e);
        }
    }

    /**
     * Return current active sheet name
     * @return string
     */
    public function getActiveSheet()
    {
        return $this->excel->getActiveSheet()->getTitle();
    }


    public function modify(ModifyDataCommand $mdc)
    {

    }

    public function read(ReadDataCommand $rdc)
    {

    }


    public function format(FormatDataCommand $fdc)
    {

    }


    public function getExcelStream()
    {

    }

    public function getPdfStream()
    {

    }

    public function getHtmlStream()
    {

    }

    public function getExcelResponse()
    {

    }

    public function getPdfResponse()
    {

    }

    public function getHtmlResponse()
    {

    }

    public function saveExcel($path)
    {

    }

    public function savePdf($path)
    {

    }


    public function saveHtml($path)
    {

    }


}