# API for simplified PhpExcel navigation and generating reports.

# PhpExcelHandler

Right now this stuff works only under *nix systems.

If you want i.e. fill excel template with data and send it to output or save locally it is for you.
Navigator classes take PhpExcel cell selection into cartesian like coordinates. 

## Installation 

If you don’t have Composer yet, you should [get it](http://getcomposer.org) now.

1. Add the package to your `composer.json`:

        "require": {
          ...
          "aaugustyniak/phpexcelhandler": "1.0.0",
          ...
        }

2. Install:

        $ php composer.phar install

3. And use:

```php
  class WriteTabularCommandTest extends TestCase
  {
      /**
       * @var SpreadSheet
       */
      private $spreadSheet;
  
      /**
       * @var WriteAnchorGuesser
       */
      private $anchorGuesser;
  
      /**
       * @var array
       */
      private $data;
  
  
      protected function setUp()
      {
          $phpExcelFactory = new DefaultPhpExcelFactory();
          $this->spreadSheet = new SpreadSheet($phpExcelFactory);
          $this->data = array(
              array("Column1", "Column2", "Column3"),
              array(1001, 2001, 3001),
              array(4001, 5001, 6001),
              array(7001, 8001, 9001),
          );
          $this->anchorGuesser = new WriteAnchorGuesser($this->data);
          $this->anchorGuesser->forceFixIndexing();
      }
  
  
      public function testReadModifiedData()
      {
          $writer = new WriteTabularCommand($this->anchorGuesser);
          $this->spreadSheet->modify($writer);
  
          $reader = new ReadTabularCommand($this->anchorGuesser);
          $readData = $this->spreadSheet->read($reader);
          $this->assertEquals($this->data, $readData);
      }

