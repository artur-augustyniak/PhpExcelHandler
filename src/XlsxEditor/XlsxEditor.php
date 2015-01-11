<?php

namespace XlsxEditor;

use n3b\Bundle\Util\HttpFoundation\StreamResponse\StreamResponse;
use n3b\Bundle\Util\HttpFoundation\StreamResponse\StreamWriterWrapper;

/**
 * @author Artur Augustyniak <artur@aaugustyniak.pl>
 */
abstract class XlsxEditor
{

    protected $outputFileName = 'output_file';
    protected $writeAreaCol = 'A';
    protected $writeAreaRow = 0;
    protected $templatePath;
    protected $activeSheet;
    protected $dataSource;
    protected $phpExcel;
    private $phpExcelWriter;

    public function __construct()
    {
        $rendererName = \PHPExcel_Settings::PDF_RENDERER_TCPDF;
        $rendererLibraryPath = realpath(dirname(__FILE__) . '/../../../../slam/tcpdf');
        \PHPExcel_Settings::setPdfRenderer($rendererName, $rendererLibraryPath);
        $this->templatePath = $this->templatePath = realpath(dirname(__FILE__) . '/../../assets/default.xlsx');
        $this->initExcelWriterWithTemplate();
    }

    public abstract function writeFromDataSource();

    /**
     * Hook for actions before rendering output file
     */
    protected function preGenerateHook()
    {
        
    }

    private function initExcelWriterWithTemplate()
    {
        $templateFileType = \PHPExcel_IOFactory::identify($this->templatePath);
        $excelTemplate = \PHPExcel_IOFactory::load($this->templatePath);
        $this->phpExcelWriter = \PHPExcel_IOFactory::createWriter($excelTemplate, $templateFileType);
        $this->phpExcelWriter->setPreCalculateFormulas(true);
        $this->phpExcel = $this->phpExcelWriter->getPHPExcel();
        $this->activeSheet = $this->phpExcel->getActiveSheet();
    }

    public function setSheetTemplatePath($path)
    {
        $this->templatePath = $path;
        $this->initExcelWriterWithTemplate();
    }

    public function setDataSource($dataSource)
    {
        $this->dataSource = $dataSource;
    }

    public function setOutputFileName($outputFileName)
    {
        $this->outputFileName = $outputFileName;
    }

    public function setWriteAreaColumn($writeAreaCol)
    {
        $this->writeAreaCol = $writeAreaCol;
    }

    public function setWriteAreaRow($writeAreaRow)
    {
        $this->writeAreaRow = $writeAreaRow;
    }

    public function generateXlsxResponse()
    {
        $writerWrapper = new StreamWriterWrapper();
        $this->preGenerateHook();
        $writerWrapper->setWriter($this->phpExcelWriter, 'save');
        $response = new StreamResponse($writerWrapper);
        $response->headers->set('Content-Type', 'text/vnd.ms-excel; charset=utf-8');
        $response->headers->set(
            'Content-Disposition', sprintf('attachment;filename=%s.xlsx', $this->outputFileName)
        );
        // If you are using a https connection, you have to set those two headers for compatibility with IE <9
        $response->headers->set('Pragma', 'public');
        $response->headers->set('Cache-Control', 'maxage=1');
        return $response;
    }

    public function generatePdfResponse()
    {
        $writerWrapper = new StreamWriterWrapper();
        $this->preGenerateHook();
        $pdfWriter = new \PHPExcel_Writer_PDF($this->phpExcel);
        $writerWrapper->setWriter($pdfWriter, 'save');
        $response = new StreamResponse($writerWrapper);
        $response->headers->set('Content-Type', 'application/pdf');
        $response->headers->set(
            'Content-Disposition', sprintf('attachment;filename=%s.pdf', $this->outputFileName)
        );
        // If you are using a https connection, you have to set those two headers for compatibility with IE <9
        $response->headers->set('Pragma', 'public');
        $response->headers->set('Cache-Control', 'maxage=1');
        return $response;
    }

    public function saveTo($dir)
    {
        $filePath = $this->makeAndValidatePathFor($dir);
        $this->phpExcelWriter->save($filePath);
        return $filePath;
    }

    private function makeAndValidatePathFor($dir)
    {
        $fullFilePath = $dir . DIRECTORY_SEPARATOR . $this->outputFileName . '.xlsx';
        if (\is_writable($dir) && !\file_exists($fullFilePath)) {
            return $fullFilePath;
        } else {
            throw new \Exception("Dir not writable or file exists.");
        }
    }

    protected function getShiftedRow($idx)
    {
        return $idx + $this->writeAreaRow;
    }

}
