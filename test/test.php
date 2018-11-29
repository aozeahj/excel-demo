<?php
/**
 * Created by PhpStorm.
 * User: aozeahj
 * Date: 2018/11/16
 * Time: 6:04 PM
 */

require __DIR__ . "/../vendor/autoload.php";

$powerExcel = new \ExcelDemo\NormalTable();

$powerExcel->singleSheet(['name','class'], [['zhangsan','1304'],['lisi', 222]], '测试', 'kkk', '/tmp/', true);

//$excel_data[] = array('sheet_name' => '111', 'sheet_title' => ['城市','端','uv','pv'], 'dim_cnt' => 2, 'sheet_data' =>[['北京','app',222,10000],['北京', 'app',333,40000],['北京', 'apph5',333,40000], ['上海', 'apph5',333,40000],['杭州', 'apph5',333,40000],['杭州', 'pc',333,40000]]);
//$excel_data[] = array('sheet_name' => '222', 'sheet_title' => ['城市','端','uv','pv'], 'dim_cnt' => 2, 'sheet_data' =>[['北京','app',222,10000],['北京', 'app',333,40000],['北京', 'apph5',333,40000], ['上海', 'apph5',333,40000],['杭州', 'apph5',333,40000],['杭州', 'pc',333,40000]]);


//$powerExcel->singleSheet(['城市','端','uv','pv'], 2, [['北京','app',222,10000],['北京', 'app',333,40000],['北京', 'apph5',333,40000], ['上海', 'apph5',333,40000],['杭州', 'apph5',333,40000],['杭州', 'pc',333,40000]], '测试', 'kkk', '/tmp/', true);

$powerExcel->numerouseSheet($excel_data, '多sheet测试', '/tmp/', true);

//$powerExcel->merge([],[],'cee');