# API for simplified PhpExcel navigation and generating reports.

# PhpExcelHandler

Right now this stuff works only under *nix systems.
It's simple wrapper API for nohup exec call.

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
    $editor = new ArrayXlsxEditor();
    $editor->setSheetTemplatePath('some_file.xlsx');
    $editor->setOutputFileName('example_output');
    $editor->setDataSource($_SERVER);
    $editor->writeFromDataSource();
    $response = $editor->generateXlsxResponse();
    //or
    //$response = $editor->generatePdfResponse();
    $response->send();

