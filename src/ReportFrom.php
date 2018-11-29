<?php
/**
 * Created by PhpStorm.
 * User: aozeahj
 * Date: 2018/11/29
 * Time: 2:53 PM
 *
 * 报表，即以维度和指标组合的表格，支持相同维度的行合并
 *
 */

namespace ExcelDemo;

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

class ReportFrom extends Base{

    /**
     * 用来暂存需要满足合并条件的维度值，最后使用依靠这个合并行单元格
     * 格式
     * array(
     *      array('start_row' => 1, 'end_row' =>3, 'p_value' => '北京') //第一列为城市维度，1～3行的值都是北京，最后合并
     *      array('start_row' => 2, 'end_row' =>3, 'p_value' => 'app') // 第二列为客户端维度，2～3行的值都是app，最后合并
     * )
     * @var array
     */
    private $_merge_dim_entry = array();

    public function singleSheet($sheet_title = array(), $dim_cnt, $sheet_data = array(), $sheetName = '', $filename = '', $save_path = '/tmp/', $is_down = true){
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
                $activeSheet->setCellValue($this->getPcoordinate($column_index), $title);
            }

            //换行
            $this->nextRow();
        }

        if (!empty($sheet_data)){
            $pre_sheet_row = [];
            foreach ($sheet_data as $index => $sheet_row){
                $continue_compare = true;
                foreach ( $sheet_row as $column_index => $column_value){

                    //上一行
                    $pre_column_value =  isset($pre_sheet_row[$column_index]) ? $pre_sheet_row[$column_index] : false;

                    //判断当前行维度列是否与上一行值相同，相同记录到合并中
                    if ($continue_compare && $column_index < $dim_cnt && $pre_column_value === $column_value){
                        $this->registerMergeEntry($column_index, $this->_row_index, $column_value);
                    }else{
                        $activeSheet->setCellValue($this->getPcoordinate($column_index), $column_value);

                        if ($column_index == 0){
                            $continue_compare = false;
                            $this->mergeRegisterEntry($activeSheet);//合并记录需要合并行
                        }
                    }
                }

                $pre_sheet_row = $sheet_row;
                $this->nextRow();
            }
            $this->mergeRegisterEntry($activeSheet);//合并记录需要合并行
        }

        //全局设置单元格格式
        $row_cnt= count($sheet_data)+2;
        $column_cnt = count($sheet_title);
        $need_set_style_range = $this->getPrange(0, 2, $column_cnt, $row_cnt);


        $activeSheet->getStyle($need_set_style_range)->getAlignment()->setHorizontal($this->getHorizontal());
        $activeSheet->getStyle($need_set_style_range)->getAlignment()->setVertical($this->getVertical());


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
            $dim_cnt     = isset($sheet['dim_cnt']) ? $sheet['dim_cnt'] : 1;

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
                    $activeSheet->setCellValue($this->getPcoordinate($column_index), $title);
                }

                //换行
                $this->nextRow();
            }

            if (!empty($sheet_data)){
                $pre_sheet_row = [];
                foreach ($sheet_data as $index => $sheet_row){
                    $continue_compare = true;
                    foreach ( $sheet_row as $column_index => $column_value){

                        //上一行
                        $pre_column_value =  isset($pre_sheet_row[$column_index]) ? $pre_sheet_row[$column_index] : false;

                        //判断当前行维度列是否与上一行值相同，相同记录到合并中
                        if ($continue_compare && $column_index < $dim_cnt && $pre_column_value === $column_value){
                            $this->registerMergeEntry($column_index, $this->_row_index, $column_value);
                        }else{
                            $activeSheet->setCellValue($this->getPcoordinate($column_index), $column_value);

                            if ($column_index == 0){
                                $continue_compare = false;
                                $this->mergeRegisterEntry($activeSheet);//合并记录需要合并行
                            }
                        }
                    }

                    $pre_sheet_row = $sheet_row;
                    $this->nextRow();
                }
                $this->mergeRegisterEntry($activeSheet);//合并记录需要合并行
            }

            //全局设置单元格格式
            $row_cnt= count($sheet_data)+2;
            $column_cnt = count($sheet_title);
            $need_set_style_range = $this->getPrange(0, 2, $column_cnt, $row_cnt);


            $activeSheet->getStyle($need_set_style_range)->getAlignment()->setHorizontal($this->getHorizontal());
            $activeSheet->getStyle($need_set_style_range)->getAlignment()->setVertical($this->getVertical());

        }

        $writer = new Xlsx($preadSheet);

        if ($is_down){
            $this->download($writer, $filename);
        }

        return $this->saveLocalExcel($writer, $filename, $save_path);
    }



    private function registerMergeEntry($column_index, $row_index, $p_value){
        if (isset($this->_merge_dim_entry[$column_index])){
            $this->_merge_dim_entry[$column_index]['end_row'] = $row_index;
        }else{
            $this->_merge_dim_entry[$column_index] = array(
                'start_row' => $row_index-1,
                'end_row' => $row_index,
                'value' => $p_value,
            );
        }
    }

    private function mergeRegisterEntry(Worksheet &$activeSheet){
        if (empty($this->_merge_dim_entry)){
            return ;
        }

        foreach ($this->_merge_dim_entry as $dim_index => $entry){
            $range = $this->getPrange($dim_index, $entry['start_row'], $dim_index, $entry['end_row']);
            $activeSheet->mergeCells($range);
            $activeSheet->setCellValue($this->getPcoordinate($dim_index, $entry['end_row']), $entry['value']);
        }

        $this->_merge_dim_entry = [];
    }
}