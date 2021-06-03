<?
global $PHPEXCELPATH;
if (ini_get('mbstring.func_overload') & 2) {
	$PHPEXCELPATH = __DIR__ . "/lib/PHPExcel/Classes_overload2";
} else {
	$PHPEXCELPATH = __DIR__ . "/lib/PHPExcel/Classes_overload0";
}

?>
