<?php
/**
 * 测试所用数据
 */
$expTableData = array(
    array(
        'id'=>1,
        'name'=>"李青",
        'nickname'=>"盲僧",
        'intro'=>"李青是一个近战战士型英雄，拥有很高的机动性和爆发力，单挑和小规模团战能力很强，同时李青也是非常优秀的打野英雄，非常擅长野区的遭遇战和Gank，是非常致命的英雄人物。",
    ),
    array(
        'id'=>2,
        'name'=>"迅捷斥候",
        'nickname'=>"提莫",
        'intro'=>"提莫是一个偏向可爱风的英雄，但本身实力并不弱，虽然射程短，但是输出能力非常高，不要被娇小的外表欺骗了，玩好提莫上单一样可以带领团队走向胜利。",
    ),
);
$expCellName = array(
    array( "id",'序号'),
    array( "name",'名字'),
    array( "nickname",'昵称'),
    array( "intro",'简介'),
);

/**
 * $expCellName 表头
 * $expTableData 数据
 */
//实例PHPExcel
require_once  'Classes/PHPExcel.php';
$objPHPExcel = new \PHPExcel();
//所有表头字号
$cellName = array('A','B','C','D','E','F','G','H','I','J',
    'K','L','M','N','O','P','Q','R','S','T','U',
    'V','W','X','Y','Z','AA','AB','AC','AD','AE',
    'AF','AG','AH','AI','AJ','AK','AL','AM','AN',
    'AO','AP','AQ','AR','AS','AT','AU','AV','AW','AX','AY','AZ');
//计算数据行数和列数
$cellNum = count($expCellName);
$dataNum = count($expTableData);
//添加表头
for($i=0;$i<$cellNum;$i++){
    //设置列宽度
//    $objPHPExcel->getActiveSheet()->getColumnDimension()->setWidth(20);
    //设置表头内容
    $objPHPExcel->setActiveSheetIndex(0)->setCellValue($cellName[$i].'1', $expCellName[$i][1]);
}
//添加数据
for($i=0;$i<$dataNum;$i++){
    for($j=0;$j<$cellNum;$j++){

        //设置列高度
//        $objPHPExcel->getActiveSheet()->getRowDimension(2)->setRowHeight(100);

        //字体垂直居中
//        $objPHPExcel->getActiveSheet()->getStyle('A1')->getAlignment()->setVertical(\PHPExcel_Style_Alignment::VERTICAL_CENTER);

        //合并行
//        $objPHPExcel->getActiveSheet()->mergeCells('A1:B1');

        //合并列
//        $objPHPExcel->getActiveSheet()->mergeCells('A1:A2');

        //设置字体  createText()添加默认字体的内容，createTextRun，添加自定义字体的内容，$objRichText为拼接的内容
//        $objRichText = new \PHPExcel_RichText();
//        $objRichText->createText($cellName[$j].($i+2));
//        $objPayable = $objRichText->createTextRun("自定义");
//        $objPayable->getFont()->setBold(true);
//        $objPayable->getFont()->setItalic(true);
//        $objPayable->getFont()->setColor(new PHPExcel_Style_Color('FF008000'));
//        $objPayable->getFont()->setUnderline(true);
//        $objRichText->createText($expTableData[$i][$expCellName[$j][0]]);
//        $objPHPExcel->getActiveSheet(0)->setCellValue($cellName[$j].($i+2), $objRichText);

//        //插入超链接
//        $objPHPExcel->getActiveSheet()->getCell($cellName[$j].($i+2))->getHyperlink()->setUrl('https://github.com/PHPOffice/PHPExcel');
//        $objPHPExcel->getActiveSheet()->getCell($cellName[$j].($i+2))->getHyperlink()->setTooltip('https://github.com/PHPOffice/PHPExcel');

        //插入图片
//        $objDrawing = new PHPExcel_Worksheet_Drawing();
//        $objDrawing->setPath('assets/1.jpg');
//        $objDrawing->setHeight(30);
//        $objDrawing->setWidth(30);
//        $objDrawing->setCoordinates('A3');
//        $objDrawing->setOffsetX(12);
//        $objDrawing->setOffsetY(12);
//        $objDrawing->setWorksheet($objPHPExcel->getActiveSheet());

        $objPHPExcel->getActiveSheet(0)->setCellValue($cellName[$j].($i+2), $expTableData[$i][$expCellName[$j][0]]);
    }
}
//求平均数：=AVERAGE(A2:A3)；求和：=SUM(A2:A3)；最大值：=MAX(A2:A3)；最小值：=MIN(A2:A4)；计数：=COUNT(A2:A4)；其他函数参考excel客户端
//$objPHPExcel->getActiveSheet()->setCellValue('A4', '=AVERAGE(A2:A3)');


//Excel5：xls格式，Excel2007：xlsx格式，HTML：导出html文件，PDF：导出pdf文件，导出pdf格式需要PDF渲染器
$objWriter = \PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');

//保存excel文件到本地，用于下载
$file = "assets/导出excel.xls";
$objWriter->save($file);
echo $file;

//直接打开
//$objWriter->save('php://output');
exit;