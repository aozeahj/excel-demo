<?php
/**
 * Created by PhpStorm.
 * User: aozeahj
 * Date: 2018/11/29
 * Time: 5:52 PM
 * 普通的excel 表格格式，原始数据如何排列，excel数据就如何排列
 */

namespace ExcelDemo;

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

class NormalTable extends Base{

    public function singleSheet($sheet_title = array(), $sheet_data = array(), $sheetName = '', $filename = '', $save_path = '/tmp/', $is_down = true){
        $preadSheet = new Spreadsheet();

        $activeSheet = $preadSheet->getActiveSheet();//sheet 从0开始
        $activeSheet->setTitle($sheetName);


        if (!empty($sheet_title)){
            //写第一行数据
            $title_cnt = count($sheet_title);
            $range = $this->getPrangeInRow(0, $title_cnt, $this->_row_index);
            $activeSheet->mergeCells($range);
            $activeSheet->setCellValue($this->getCoordinateDuringRange($range), '数据导出：'. date('Y-m-d H:i:s'));

            $this->nextRow();
            //写第二行数据，标题
            foreach ($sheet_title as $column_index => $title){
                $activeSheet->setCellValue($this->getPcoordinate($column_index), ' ' . $title);
            }

            //换行
            $this->nextRow();
        }

        if (!empty($sheet_data)){
            foreach ($sheet_data as $index => $sheet_row){
                foreach ( $sheet_row as $column_index => $column_value){
                    $activeSheet->setCellValue($this->getPcoordinate($column_index), ' ' . $column_value);
                }

                $this->nextRow();
            }
        }

        //全局设置单元格格式
        $row_cnt= count($sheet_data)+2;
        $column_cnt = count($sheet_title);
        $need_set_style_range = $this->getPrange(0, 2, $column_cnt, $row_cnt);

        $this->setStyle($activeSheet, $need_set_style_range);

        $writer = new Xlsx($preadSheet);

        if ($is_down){
            $this->download($writer, $filename);
        }

        return $this->saveLocalExcel($writer, $filename, $save_path);
    }


    public function numerouseSheet($excel_data, $filename = '', $save_path = '/tmp/', $is_down = true){
        if (empty($excel_data)){
            return false;
        }

        $preadSheet = new Spreadsheet();
        foreach ($excel_data as $sheet_index => $sheet){

            //Spreadsheet 实例化时，自动生成 第 0 格sheet
            if ($sheet_index == $this->_default_sheet_index){
                $activeSheet = $preadSheet->getActiveSheet();
            }else{
                $activeSheet = $preadSheet->createSheet($sheet_index);
            }

            $sheet_name  = isset($sheet['sheet_name']) ? $sheet['sheet_name'] : 'sheet'. $sheet_index;
            $sheet_title = isset($sheet['sheet_title']) ? $sheet['sheet_title'] : [];
            $sheet_data  = isset($sheet['sheet_data']) ? $sheet['sheet_data'] : [];

            $activeSheet->setTitle($sheet_name);

            $this->setRowIndex(1);
            if (!empty($sheet_title)){
                //写第一行数据
                $title_cnt = count($sheet_title);
                $range = $this->getPrangeInRow(0, $title_cnt, $this->_row_index);
                $activeSheet->mergeCells($range);
                $activeSheet->setCellValue($this->getCoordinateDuringRange($range), '数据导出：'. date('Y-m-d H:i:s'));

                $this->nextRow();
                //写第二行数据，标题
                foreach ($sheet_title as $column_index => $title){
                    $activeSheet->setCellValue($this->getPcoordinate($column_index), ' ' . $title);
                }

                //换行
                $this->nextRow();
            }

            if (!empty($sheet_data)){
                foreach ($sheet_data as $index => $sheet_row){
                    foreach ( $sheet_row as $column_index => $column_value){
                        $activeSheet->setCellValue($this->getPcoordinate($column_index), ' ' . $column_value);
                    }

                    $this->nextRow();
                }
            }

            //全局设置单元格格式
            $row_cnt= count($sheet_data)+2;
            $column_cnt = count($sheet_title);
            $need_set_style_range = $this->getPrange(0, 2, $column_cnt, $row_cnt);

            $this->setStyle($activeSheet, $need_set_style_range);
        }

        $writer = new Xlsx($preadSheet);

        if ($is_down){
            $this->download($writer, $filename);
        }

        return $this->saveLocalExcel($writer, $filename, $save_path);
    }

}