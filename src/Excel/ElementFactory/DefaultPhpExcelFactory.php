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

use Aaugustyniak\PhpExcelHandler\Excel\NoSuchElement;
use Aaugustyniak\PhpExcelHandler\Excel\WriterWrapper;
use Composer\Autoload\ClassLoader;

/**
 * @author Artur Augustyniak <artur@aaugustyniak.pl>
 */
class DefaultPhpExcelFactory implements PHPExcelElementFactory
{

    function __construct()
    {
        /**
         * @var ClassLoader
         */
        $loader = $this->getLoader();
        $rendererPath = dirname($loader->findFile('TCPDF'));
        $this->setPdfRendererPath($rendererPath);
    }

    /**
     * @param $path
     */
    public function setPdfRendererPath($path)
    {
        $rendererName = \PHPExcel_Settings::PDF_RENDERER_TCPDF;
        \PHPExcel_Settings::setPdfRenderer($rendererName, $path);
    }


    private function getLoader()
    {
        foreach (spl_autoload_functions() as $func) {
            if (is_array($func) && isset($func[1]) && $func[1] == 'loadClass') {
                return $func[0];
            }
        }
        $msg = sprintf("Can't find autoloader, please use %s via istance of class %s",
            'setPdfRendererPath($path)',
            __CLASS__);
        throw new NoSuchElement($msg);
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
        $phpExcelWriter->setPreCalculateFormulas(true);
        return $phpExcelWriter->getPHPExcel();
    }


    /**
     * @param \PHPExcel $pe
     * @return WriterWrapper
     */
    public function newPHPExcelPdfWriterFrom(\PHPExcel $pe)
    {
        return new WriterWrapper(new \PHPExcel_Writer_PDF($pe));
    }

    /**
     * @return WriterWrapper
     */
    public function newPHPExcelHtmlWriterFrom(\PHPExcel $pe)
    {
        return new WriterWrapper( new  \PHPExcel_Writer_HTML($pe));
    }

    /**
     * @return WriterWrapper
     */
    public function newPHPExcelWriterFrom(\PHPExcel $pe)
    {
        return new WriterWrapper(new \PHPExcel_Writer_Excel2007($pe));
    }
}