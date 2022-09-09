<?php

include __DIR__ . "/vendor/autoload.php";

use App\ExcelToCsv;

$excelfile = new ExcelToCsv("public/import/excelFiles");
// $excelfile->convert();
