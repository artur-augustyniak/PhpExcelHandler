[![Build Status](https://travis-ci.org/artur-augustyniak/PhpExcelHandler.svg?branch=master)](https://travis-ci.org/artur-augustyniak/PhpExcelHandler)
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
require_once "vendor/autoload.php";

$data = array(
            array("Column1", "Column2", "Column3"),
            array(1001, 2001, 3001),
            array(4001, 5001, 6001),
            array(7001, 8001, 9001),
        );


$phpExcelFactory = new DefaultPhpExcelFactory();
$spreadSheet = new SpreadSheet($phpExcelFactory);

$anchorGuesser = new WriteAnchorGuesser($data);

$writer = new WriteTabularCommand($anchorGuesser);
$spreadSheet->modify($writer);
$outputHtml = $spreadSheet->getHtmlStream();
echo $outputHtml;


