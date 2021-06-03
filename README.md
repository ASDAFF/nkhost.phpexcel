# nkhost.phpexcel

**Описание**

Библиотека PHPExcel с поддержкой mbstring.func_overload=0 и mbstring.func_overload=2. Данное решение решает проблему подключения PHPExcel к Bitrix "Управление сайтом" Без необходимости изменять func_overload на сервере. Просто берите, устанавливайте и пользуйтесь.

После установки подключите глобальную переменную $PHPEXCELPATH в своем скрипте.
В данной переменной хранится путь к библиотекам PHPExcel в том виде, как они хранятся в оригинальной библиотеке Сайт PHPExcel
Пример подключения и использования:
```php
if (CModule::IncludeModule("nkhost.phpexcel")){
     global $PHPEXCELPATH;      

      // Ваш код далее
      require_once ($PHPEXCELPATH . '/PHPExcel/IOFactory.php');  
      $xls = PHPExcel_IOFactory::load("/tmp/file.xlsx");
      // Устанавливаем индекс активного листа
      $xls->setActiveSheetIndex(0);
      // Получаем активный лист
      $sheet = $xls->getActiveSheet();

      for ($i = 1; $i <= $sheet->getHighestRow(); $i++) {
      $nColumn = PHPExcel_Cell::columnIndexFromString($sheet->getHighestColumn());
      for ($j = 0; $j < $nColumn; $j++) {
         $arProducts[$i][$j] = $sheet->getCellByColumnAndRow($j, $i)->getValue();
      }
      }
      
      var_dump($arProducts);
}
```
Или создание файла

```php
<?php
require($_SERVER["DOCUMENT_ROOT"] . "/bitrix/modules/main/include/prolog_before.php");

if (CModule::IncludeModule("nkhost.phpexcel")){
     global $PHPEXCELPATH;      
 // Подключаем класс для работы с excel
require_once($PHPEXCELPATH.'/PHPExcel.php');
// Подключаем класс для вывода данных в формате excel
require_once($PHPEXCELPATH.'/PHPExcel/Writer/Excel5.php');
 
// Создаем объект класса PHPExcel
$xls = new PHPExcel();
// Устанавливаем индекс активного листа
$xls->setActiveSheetIndex(0);
// Получаем активный лист
$sheet = $xls->getActiveSheet();
// Подписываем лист
$sheet->setTitle('Таблица умножения');
 
// Вставляем текст в ячейку A1
$sheet->setCellValue("A1", 'Таблица умножения mbstring.func_overload='. ini_get('mbstring.func_overload'));
$sheet->getStyle('A1')->getFill()->setFillType(
    PHPExcel_Style_Fill::FILL_SOLID);
$sheet->getStyle('A1')->getFill()->getStartColor()->setRGB('EEEEEE');
 
// Объединяем ячейки
$sheet->mergeCells('A1:H1');
 
// Выравнивание текста
$sheet->getStyle('A1')->getAlignment()->setHorizontal(
    PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
 
for ($i = 2; $i < 10; $i++) {
    for ($j = 2; $j < 10; $j++) {
        // Выводим таблицу умножения
        $sheet->setCellValueByColumnAndRow(
                                          $i - 2,
                                          $j,
                                          $i . "x" .$j . "=" . ($i*$j));
        // Применяем выравнивание
        $sheet->getStyleByColumnAndRow($i - 2, $j)->getAlignment()->
                setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
    }
}
     // Выводим HTTP-заголовки
 header ( "Expires: Mon, 1 Apr 1974 05:00:00 GMT" );
 header ( "Last-Modified: " . gmdate("D,d M YH:i:s") . " GMT" );
 header ( "Cache-Control: no-cache, must-revalidate" );
 header ( "Pragma: no-cache" );
 header ( "Content-type: application/vnd.ms-excel" );
 header ( "Content-Disposition: attachment; filename=matrix.xls" );
 
// Выводим содержимое файла
 $objWriter = new PHPExcel_Writer_Excel5($xls);
 $objWriter->save('php://output');
}
```