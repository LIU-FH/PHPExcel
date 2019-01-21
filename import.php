<?php
//本地文件地址
$file = "assets/导出excel.xls";

require_once 'Classes/PHPExcel.php';
// 判断文件是什么格式
$type = pathinfo($file);
$type = strtolower($type["extension"]);
if ($type=='xlsx') {
    $type='Excel2007';
}elseif($type=='xls') {
    $type = 'Excel5';
}
// 判断使用哪种格式
$objReader = \PHPExcel_IOFactory::createReader($type);
$objPHPExcel = $objReader->load($file);
$sheet = $objPHPExcel->getSheet(0);
// 取得总行数
$highestRow = $sheet->getHighestRow();
// 取得总列数
$highestColumn = $sheet->getHighestColumn();
//总列数转换成数字
$numHighestColum = \PHPExcel_Cell::columnIndexFromString("$highestColumn");
//循环读取excel文件,读取一条,插入一条
$data=array();
//从第一行开始读取数据
for($j=1;$j<=$highestRow;$j++){
    //从A列读取数据
    for($k=0;$k<=$numHighestColum-1;$k++){
        //数字列转换成字母
        $columnIndex = \PHPExcel_Cell::stringFromColumnIndex($k);
        // 读取单元格
        $data[$j][]=$objPHPExcel->getActiveSheet()->getCell("$columnIndex$j")->getFormattedValue();
    }
}
print_r($data);
exit();