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

use n3b\Bundle\Util\HttpFoundation\StreamResponse\StreamResponse;
use n3b\Bundle\Util\HttpFoundation\StreamResponse\StreamWriterWrapper;
use \PHPExcel as PHPExcel;
use \Exception as Exception;
use \PHPExcel_Writer_Abstract as AbstractWriter;
use Aaugustyniak\PhpExcelHandler\Excel\ActionCommand\FormatDataCommand;
use Aaugustyniak\PhpExcelHandler\Excel\ActionCommand\ModifyDataCommand;
use Aaugustyniak\PhpExcelHandler\Excel\ActionCommand\ReadDataCommand;
use Aaugustyniak\PhpExcelHandler\Excel\ElementFactory\PHPExcelElementFactory;

/**
 *
 * @param $path
 */

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
     * @var \PHPExcel_Writer_Excel2007
     */
    private $writer;

    /**
     * @param PHPExcelElementFactory $elementFactory
     */
    function __construct(PHPExcelElementFactory $elementFactory)
    {
        $this->fileName = self::DEF_FILE_NAME;
        $this->elementFactory = $elementFactory;
        $this->excel = $this->elementFactory->newPHPExcelObject();
        $this->writer = $this->elementFactory->newPHPExcelWriterFrom($this->excel);
    }

    /**
     * Get name of current file (without extension)
     * @return string
     */
    public function getFileName()
    {
        return $this->fileName;
    }

    /**
     * Open file for edit
     * @param $path
     * @throws NoSuchElement
     */
    public function openFile($path)
    {
        try {
            $this->excel = $this->elementFactory->newPHPExcelObjectFromFile($path);
            $this->fileName = $this->extractFileNameFromPath($path);
        } catch (Exception $e) {
            throw new NoSuchElement("No such file", $e->getCode(), $e);
        }
    }

    /**
     * Return file name without extension
     * @param $path
     * @return string
     */
    private function extractFileNameFromPath($path)
    {
        return basename($path, '.' . self::EXCEL_EXT);
    }

    /**
     * Make workbook modifications provided by user implementation of
     * ModifyDataCommand
     *
     * @param ModifyDataCommand $mdc
     */
    public function modify(ModifyDataCommand $mdc)
    {
        $mdc->modify($this->excel);
    }

    /**
     * Scrap  workbook data using user strategy implemented by
     * ReadDataCommand
     *
     * @param ReadDataCommand $rdc
     * @return mixed
     */
    public function read(ReadDataCommand $rdc)
    {
        $rdc->readFrom($this->excel);
        return $rdc->fetchData();
    }

    /**
     * Format workbook/cells with strategy  provided by user implementation of
     * FormatDataCommand
     *
     * @param FormatDataCommand $fdc
     */
    public function format(FormatDataCommand $fdc)
    {
        $fdc->format($this->excel);
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


    /**
     * Get raw excel file stream
     *
     * @return string
     * @throws Exception
     */
    public function getExcelStream()
    {
        $output = null;
        try {
            $excelWriter = $this->elementFactory->newPHPExcelWriterFrom($this->excel);
            \ob_start();
            $excelWriter->save('php://output');
            $output = \ob_get_clean();
        } catch (Exception $e) {
            throw new Exception("Cannot write to php://output", $e->getCode(), $e);
        }
        return $output;
    }

    /**
     * * Get raw pdf file stream
     *
     * @return string
     * @throws Exception
     */
    public function getPdfStream()
    {
        $output = null;
        try {
            $pdfWriter = $this->elementFactory->newPHPExcelPdfWriterFrom($this->excel);
            \ob_start();
            $pdfWriter->save('php://output');
            $output = \ob_get_clean();
        } catch (Exception $e) {
            throw new Exception("Cannot write to php://output", $e->getCode(), $e);
        }
        return $output;
    }

    /**
     * * Get raw html file stream
     *
     * @return string
     * @throws Exception
     */
    public function getHtmlStream()
    {
        $output = null;
        try {
            $pdfWriter = $this->elementFactory->newPHPExcelHtmlWriterFrom($this->excel);
            \ob_start();
            $pdfWriter->save('php://output');
            $output = \ob_get_clean();
        } catch (Exception $e) {
            throw new Exception("Cannot write to php://output", $e->getCode(), $e);
        }
        return $output;
    }

    /**
     * @return StreamResponse
     */
    public function getExcelResponse()
    {
        $writerWrapper = new StreamWriterWrapper();
        $writerWrapper->setWriter($this->elementFactory->newPHPExcelWriterFrom($this->excel), 'save');
        $response = new StreamResponse($writerWrapper);
        $response->headers->set('Content-Type', 'text/vnd.ms-excel; charset=utf-8');
        $response->headers->set(
            'Content-Disposition', sprintf('attachment;filename=%s.xlsx', $this->fileName)
        );
        // If you are using a https connection, you have to set those two headers for compatibility with IE <9
        $response->headers->set('Pragma', 'public');
        $response->headers->set('Cache-Control', 'maxage=1');
        return $response;
    }

    /**
     * @return StreamResponse
     */
    public function getPdfResponse()
    {
        $writerWrapper = new StreamWriterWrapper();
        $writerWrapper->setWriter($this->elementFactory->newPHPExcelPdfWriterFrom($this->excel), 'save');
        $response = new StreamResponse($writerWrapper);
        $response->headers->set('Content-Type', 'application/pdf');
        $response->headers->set(
            'Content-Disposition', sprintf('attachment;filename=%s.pdf', $this->fileName)
        );
        // If you are using a https connection, you have to set those two headers for compatibility with IE <9
        $response->headers->set('Pragma', 'public');
        $response->headers->set('Cache-Control', 'maxage=1');
        return $response;
    }

    public function getHtmlResponse()
    {
        $writerWrapper = new StreamWriterWrapper();
        $writerWrapper->setWriter($this->elementFactory->newPHPExcelHtmlWriterFrom($this->excel), 'save');
        $response = new StreamResponse($writerWrapper);
        $response->headers->set('Content-Type', 'text/html');
        $response->headers->set(
            'Content-Disposition', sprintf('attachment;filename=%s.html', $this->fileName)
        );
        // If you are using a https connection, you have to set those two headers for compatibility with IE <9
        $response->headers->set('Pragma', 'public');
        $response->headers->set('Cache-Control', 'maxage=1');
        return $response;
    }

    /**
     * @param $path
     * @throws Exception
     */
    public function saveExcel($path)
    {
        $writer = $this->elementFactory->newPHPExcelWriterFrom($this->excel);
        $this->saveOn($path, $writer);
    }

    /**
     * @param $path
     * @throws Exception
     */
    public function savePdf($path)
    {
        $writer = $this->elementFactory->newPHPExcelPdfWriterFrom($this->excel);
        $this->saveOn($path, $writer);
    }

    /**
     * @param $path
     * @throws Exception
     */
    public function saveHtml($path)
    {
        $writer = $this->elementFactory->newPHPExcelHtmlWriterFrom($this->excel);
        $this->saveOn($path, $writer);
    }

    /**
     *
     * @param $path
     * @param WriterWrapper $writer
     * @throws Exception
     */
    private function saveOn($path, WriterWrapper $writer)
    {
        try {
            $writer->save($path);
        } catch (Exception $e) {
            $msg = sprintf("Cannot save file %s", $path);
            throw new Exception($msg, $e->getCode(), $e);
        }
    }


}