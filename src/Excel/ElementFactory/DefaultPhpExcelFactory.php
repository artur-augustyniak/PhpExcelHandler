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

namespace Aaugustyniak\PhpExcelHandler\Excel\ElementFactory;

/**
 * @author Artur Augustyniak <artur@aaugustyniak.pl>
 */
class DefaultPhpExcelFactory implements PHPExcelElementFactory
{
    /**
     * @TODO test/prod paths
     */
    function __construct()
    {
        $rendererName = \PHPExcel_Settings::PDF_RENDERER_TCPDF;
        $rendererLibraryPath = __DIR__ . '/../../../vendor/tecnick.com/tcpdf';

        \PHPExcel_Settings::setPdfRenderer($rendererName, $rendererLibraryPath);
    }


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

    /**
     * @return \PHPExcel
     */
    public function newPHPExcelObjectFromFile($path)
    {
        $templateFileType = \PHPExcel_IOFactory::identify($path);
        $excelTemplate = \PHPExcel_IOFactory::load($path);
        $phpExcelWriter = \PHPExcel_IOFactory::createWriter($excelTemplate, $templateFileType);
        //$this->phpExcelWriter->setPreCalculateFormulas(true);
        return $phpExcelWriter->getPHPExcel();
    }

    /**
     * @return \PHPExcel_Writer_Excel2007
     */
    public function newPHPExcelWriter()
    {
        return new \PHPExcel_Writer_Excel2007();
    }

    /**
     * @param \PHPExcel $pe
     * @return \PHPExcel_Writer_PDF
     */
    public function newPHPExcelPdfWriterFrom(\PHPExcel $pe)
    {
        return new \PHPExcel_Writer_PDF($pe);
    }

    /**
     * @return \PHPExcel_Writer_HTML
     */
    public function newPHPExcelHtmlWriterFrom(\PHPExcel $pe)
    {
        return new \PHPExcel_Writer_HTML($pe);
    }

    /**
     * @return \PHPExcel_Writer_Excel2007
     */
    public function newPHPExcelWriterFrom(\PHPExcel $pe)
    {
        return new \PHPExcel_Writer_Excel2007($pe);
    }
}