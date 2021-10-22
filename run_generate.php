<?php
require_once 'phpdocx/Classes/Phpdocx/Create/CreateDocx.php';
require_once 'phpdocx/Classes/Phpdocx/Create/CreateDocxFromTemplate.php';

require_once 'PHPExcel/PHPExcel/Classes/PHPExcel.php';
require_once 'PHPExcel/PHPExcel/Classes/PHPExcel/IOFactory.php';

// param parse
$args = $_SERVER ['argv'];
$rate = $argv[1];

//$objPHPExcel = PHPExcel_IOFactory::load("invoice_receipt.xlsx");
$objPHPExcel = PHPExcel_IOFactory::load("invoice_receipt.xlsm");

$lastDayLastMonth = strtotime('last day of previous month');
$toDay = strtotime('today');
$salary = 10000000;
if (empty($rate)){
    $rate = 202.38;
}

$address = 'Your_address';
$name = 'Vu Van Ly';
$nameNotSpace = str_replace(' ', '_', $name);
$phoneNumber = 'Your_phone';

$dataArray = array(
    'Z4' => "NO.r-" . date('ymd', $toDay) ."-0001",
    'B9' => '★　　VND'. number_format($salary) . '-
('. number_format($salary/$rate) .' 円)',
    'B17' => '※' . date('Y年m月d日', $lastDayLastMonth) . 'の為替レート（' . $rate . ' VND：1円)',
    'B19' => date('Y年m月d日', $toDay) . '　上記の金額、正に領収いたしました。',
    'D25' => date('Y年m月', $lastDayLastMonth) . '分の報酬',
    'J25' => $salary,
    'T39' => "Address: $address

$name

TEL：" . $phoneNumber . "　　　　印"
);

$objPHPExcel->setActiveSheetIndex(0);

foreach ($dataArray as $cell => $value){
    $objPHPExcel->getActiveSheet()->setCellValue($cell, $value);
}
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
//$excelFileName = $nameNotSpace . '_' . date('Y_m', $lastDayLastMonth) .  '_receipt.xlsx';
$excelFileName = $nameNotSpace . '_' . date('Y_m', $lastDayLastMonth) .  '_receipt.xlsm';
$objWriter->save($excelFileName);

$docx = new Phpdocx\Create\CreateDocxFromTemplate('invoice_receipt.docx');

$variables = array(
    'NAME' => strtoupper($name),
    'ADDRESS' => $address,
    'PHONE_NUMBER' => $phoneNumber,
    'INVOICE_NUMBER' => date('ymd', $lastDayLastMonth) . '-0001',
    'INVOICE_DATE' => date('Y/m/d', $lastDayLastMonth),
    'RECEIPT_DATE' => date('Y/m/d', $toDay),
    'SALARY_VND' => number_format($salary),
    'SALARY_JP' => number_format($salary/$rate),
    'RATE_DATE' => date('Y年m月d日', $lastDayLastMonth),
    'RATE' => $rate,
    'SALARY_TIME' => date('Y年m月', $lastDayLastMonth),
);

$docx->replaceVariableByText($variables);
$wordFileName = $nameNotSpace . '_' . date('Y_m', $lastDayLastMonth) .  '_invoice';
$docx->createDocx($wordFileName);

//exec('"C:\Program Files\Microsoft Office 15\root\office15\winword.exe"  ' . $wordFileName  . '.docx /mFilePrintDefault /mFileExit /q /n && ' . '"C:\Program Files\Microsoft Office 15\root\office15\excel.exe"  ' . $excelFileName);
exec('"C:\Users\vanly\AppData\Local\Kingsoft\WPS Office\10.2.0.7516\office6\wps.exe" -p  ' . $wordFileName  . '.docx');
exec('"C:\Users\vanly\AppData\Local\Kingsoft\WPS Office\10.2.0.7516\office6\wps.exe" -p  ' . $wordFileName  . '.docx');

//exec("taskkill /f /im winword.exe");
//exec('"C:\Users\vanly\AppData\Local\Kingsoft\WPS Office\10.2.0.7516\office6\et.exe"  -p "' . __DIR__ . DIRECTORY_SEPARATOR .  $excelFileName . '"');
//exec('"C:\Users\vanly\AppData\Local\Kingsoft\WPS Office\10.2.0.7516\office6\et.exe"  -p "' . __DIR__ . DIRECTORY_SEPARATOR .  $excelFileName . '"');
exec('"C:\Program Files\Microsoft Office 15\root\office15\excel.exe"  ' . $excelFileName );
exec('"C:\Program Files\Microsoft Office 15\root\office15\excel.exe"  ' . $excelFileName );


?>