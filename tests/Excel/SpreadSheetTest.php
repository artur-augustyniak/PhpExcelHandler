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


use PHPExcel;
use \PHPUnit_Framework_TestCase as TestCase;
use Aaugustyniak\PhpExcelHandler\Excel\ElementFactory\PHPExcelElementFactory;
use Aaugustyniak\PhpExcelHandler\Excel\SpreadSheet;
use Symfony\Component\Finder\Finder;


class TestPhpExcelFactory implements PHPExcelElementFactory
{
    /**
     * @TODO test/prod paths
     */
    function __construct()
    {
        $rendererName = \PHPExcel_Settings::PDF_RENDERER_TCPDF;
        $rendererLibraryPath = __DIR__ . '/../../vendor/tecnick.com/tcpdf';
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
    public function newPHPExcelWriterFrom(PHPExcel $pe)
    {
        return new \PHPExcel_Writer_Excel2007($pe);
    }
}


/**
 * @author Artur Augustyniak <artur@aaugustyniak.pl>
 */
class SpreadSheetTest extends TestCase
{

    const TEST_FILE_NAME = "default";
    const DEF_WORKSHEET_NAME = 'Worksheet';
    const DEF_FILE_NAME = "Workbook";


    /**
     * @var SpreadSheet
     */
    private $spreadSheet;


    private $systemTmp;


    protected function setUp()
    {
        $this->systemTmp = \sys_get_temp_dir();
        $phpExcelFactory = new TestPhpExcelFactory();
        $this->spreadSheet = new SpreadSheet($phpExcelFactory);
    }

    public function testCreatedSpreadSheetHasDefaultFileName()
    {
        $this->assertEquals(self::DEF_FILE_NAME, $this->spreadSheet->getFileName());
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

    /**
     * @expectedException \Aaugustyniak\PhpExcelHandler\Excel\NoSuchElement
     */
    public function testOpenNotExistingFile()
    {
        $this->spreadSheet->openFile("wrong.xlsx");
    }


    public function testOpeningFileResultsInProperInternalChanges()
    {
        $this->spreadSheet->openFile($this->getTestFilePath());
        /**
         * @see ../Assets/default.xlsx
         */
        $expectedTitles = array("TEST_1", "TEST_2", "TEST_3");
        $actualTitles = $this->spreadSheet->getSheetTiles();
        $expectedActiveSheet = "TEST_1";
        $actualActiveSheet = $this->spreadSheet->getActiveSheet();
        $this->assertEquals($expectedTitles, $actualTitles);
        $this->assertEquals($expectedActiveSheet, $actualActiveSheet);
        $this->assertEquals(self::TEST_FILE_NAME, $this->spreadSheet->getFileName());
    }

    /**
     * @return string
     */
    private function getTestFilePath()
    {
        $finder = new Finder();

        $iterator = $finder
            ->files()
            ->name(self::TEST_FILE_NAME . '.xlsx')
            ->depth(2)
            ->in('.');
        $testFilePath = "";
        foreach ($iterator as $file) {
            $testFilePath = $file->getRealpath();
            break;
        }
        return $testFilePath;
    }

    public function testExcelOutputStream()
    {
        $this->spreadSheet->openFile($this->getTestFilePath());
        $actualStream = $this->spreadSheet->getExcelStream();
        $this->assertContains('[Content_Types].xml͔]', $actualStream);
    }

    public function testPdfOutputStream()
    {
        $this->spreadSheet->openFile($this->getTestFilePath());
        $actualStream = $this->spreadSheet->getPdfStream();
        $this->assertContains('%PDF-1.7', $actualStream);
    }

    public function testHtmlOutputStream()
    {
        $this->spreadSheet->openFile($this->getTestFilePath());
        $actualStream = $this->spreadSheet->getHtmlStream();
        $this->assertContains('<meta http-equiv="Content-Type" content="text/html; charset=utf-8">',
            $actualStream);
    }

    public function testExcelResponse()
    {
        $this->spreadSheet->openFile($this->getTestFilePath());
        $actualResponse = $this->spreadSheet->getExcelResponse();

        $this->assertEquals('text/vnd.ms-excel; charset=utf-8', $actualResponse->headers->get('Content-Type'));
        $this->assertEquals('attachment;filename=' . self::TEST_FILE_NAME . '.xlsx', $actualResponse->headers->get('Content-Disposition'));
        $this->assertEquals('public', $actualResponse->headers->get('Pragma'));
        $this->assertEquals('maxage=1, private', $actualResponse->headers->get('Cache-Control'));
        \ob_start();
        $actualResponse->sendContent();
        $actualContent = \ob_get_clean();
        $this->assertContains('[Content_Types].xml͔]', $actualContent);
    }

    public function testPdfResponse()
    {
        $this->spreadSheet->openFile($this->getTestFilePath());
        $actualResponse = $this->spreadSheet->getPdfResponse();

        $this->assertEquals('application/pdf', $actualResponse->headers->get('Content-Type'));
        $this->assertEquals('attachment;filename=' . self::TEST_FILE_NAME . '.pdf', $actualResponse->headers->get('Content-Disposition'));
        $this->assertEquals('public', $actualResponse->headers->get('Pragma'));
        $this->assertEquals('maxage=1, private', $actualResponse->headers->get('Cache-Control'));
        \ob_start();
        $actualResponse->sendContent();
        $actualContent = \ob_get_clean();
        $this->assertContains('%PDF-1.7', $actualContent);
    }

    public function testHtmlResponse()
    {
        $this->spreadSheet->openFile($this->getTestFilePath());
        $actualResponse = $this->spreadSheet->getHtmlResponse();

        $this->assertEquals('text/html', $actualResponse->headers->get('Content-Type'));
        $this->assertEquals('attachment;filename=' . self::TEST_FILE_NAME . '.html', $actualResponse->headers->get('Content-Disposition'));
        $this->assertEquals('public', $actualResponse->headers->get('Pragma'));
        $this->assertEquals('maxage=1, private', $actualResponse->headers->get('Cache-Control'));
        \ob_start();
        $actualResponse->sendContent();
        $actualContent = \ob_get_clean();
        $this->assertContains('<meta http-equiv="Content-Type" content="text/html; charset=utf-8">',
            $actualContent);
    }


    public function testSaveExcel()
    {
        $this->spreadSheet->openFile($this->getTestFilePath());
        $outputFilePath = $this->systemTmp . DIRECTORY_SEPARATOR . self::DEF_FILE_NAME . '.' . SpreadSheet::EXCEL_EXT;
        $this->spreadSheet->saveExcel($outputFilePath);
        $savedFile = file_get_contents($outputFilePath);
        $this->assertContains('[Content_Types].xml͔]', $savedFile);
        unlink($outputFilePath);
    }


    public function testSavePdf()
    {
        $this->spreadSheet->openFile($this->getTestFilePath());
        $outputFilePath = $this->systemTmp . DIRECTORY_SEPARATOR . self::DEF_FILE_NAME . '.' . SpreadSheet::PDF_EXT;
        $this->spreadSheet->savePdf($outputFilePath);
        $savedFile = file_get_contents($outputFilePath);
        $this->assertContains('%PDF-1.7', $savedFile);
        unlink($outputFilePath);
    }


    public function testSaveHtml()
    {
        $this->spreadSheet->openFile($this->getTestFilePath());
        $outputFilePath = $this->systemTmp . DIRECTORY_SEPARATOR . self::DEF_FILE_NAME . '.' . SpreadSheet::HTML_EXT;
        $this->spreadSheet->saveHtml($outputFilePath);
        $savedFile = file_get_contents($outputFilePath);
        $this->assertContains('<meta http-equiv="Content-Type" content="text/html; charset=utf-8">',
            $savedFile);
        unlink($outputFilePath);
    }


    public function testModify()
    {
        $commandMock = $this->getMock('Aaugustyniak\PhpExcelHandler\Excel\ActionCommand\ModifyDataCommand');
        $commandMock->expects($this->once())->method('modify');
        $this->spreadSheet->modify($commandMock);

    }

    public function testFormat()
    {
        $commandMock = $this->getMock('Aaugustyniak\PhpExcelHandler\Excel\ActionCommand\FormatDataCommand');
        $commandMock->expects($this->once())->method('format');
        $this->spreadSheet->format($commandMock);
    }


    public function testReadData()
    {
        $commandMock = $this->getMock('Aaugustyniak\PhpExcelHandler\Excel\ActionCommand\ReadDataCommand');
        $commandMock->expects($this->once())->method('readFrom');
        $commandMock->expects($this->once())->method('fetchData');
        $this->spreadSheet->read($commandMock);
    }


}
