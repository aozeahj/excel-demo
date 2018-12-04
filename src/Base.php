<?php
/**
 * Created by PhpStorm.
 * User: aozeahj
 * Date: 2018/11/29
 * Time: 2:15 PM
 *
 *                     column: x
 * ----------------------->
 * |
 * |
 * |
 * |
 * |
 * |
 * |
 * |  row : y
 * V
 */

namespace ExcelDemo;

use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

class Base{

    /**
     * excel cellName array
     * excel 列索引
     * @var array
     */
    private  $_column_index_arr = array(
        'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z',
        'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO', 'AP', 'AQ', 'AR', 'AS', 'AT', 'AU', 'AV', 'AW', 'AX', 'AY', 'AZ',
        'BA', 'BB', 'BC', 'BD', 'BE', 'BF', 'BG', 'BH', 'BI', 'BJ', 'BK', 'BL', 'BM', 'BN', 'BO', 'BP', 'BQ', 'BR', 'BS', 'BT', 'BU', 'BV', 'BW', 'BX', 'BY', 'BZ',
        'CA', 'CB', 'CC', 'CD', 'CE', 'CF', 'CG', 'CH', 'CI', 'CJ', 'CK', 'CL', 'CM', 'CN', 'CO', 'CP', 'CQ', 'CR', 'CS', 'CT', 'CU', 'CV', 'CW', 'CX', 'CY', 'CZ',
        'DA', 'DB', 'DD', 'DD', 'DE', 'DF', 'DG', 'DH', 'DI', 'DJ', 'DK', 'DL', 'DM', 'DN', 'DO', 'DP', 'DQ', 'DR', 'DS', 'DT', 'DU', 'DV', 'DW', 'DX', 'DY', 'DZ',
        'EA', 'EB', 'EE', 'ED', 'EE', 'EF', 'EG', 'EH', 'EI', 'EJ', 'EK', 'EL', 'EM', 'EN', 'EO', 'EP', 'EQ', 'ER', 'ES', 'ET', 'EU', 'EV', 'EW', 'EX', 'EY', 'EZ',
        'FA', 'FB', 'FF', 'FD', 'FE', 'FF', 'FG', 'FH', 'FI', 'FJ', 'FK', 'FL', 'FM', 'FN', 'FO', 'FP', 'FQ', 'FR', 'FS', 'FT', 'FU', 'FV', 'FW', 'FX', 'FY', 'FZ',
    );

    /**
     * 当前行索引
     * @var int
     */
    protected $_row_index = 1;

    /**
     * @var string
     */
    private $_horizontal = Alignment::HORIZONTAL_CENTER;

    /**
     * @var string
     */
    private $_vertical = Alignment::VERTICAL_CENTER;

    /**
     * spreadsheet 默认自动生成第一个sheet
     * @var int
     */
    protected $_default_sheet_index = 0;

    /**
     * @param $horizontal
     */
    public function setHorizontal($horizontal){
        $this->_horizontal = $horizontal;
    }

    /**
     * @param $vertical
     */
    public function setVertical($vertical){
        $this->_vertical = $vertical;
    }

    /**
     * @param $horizontal
     */
    public function getHorizontal(){
        return $this->_horizontal;
    }

    /**
     * @return string
     */
    public function getVertical(){
        return $this->_vertical;
    }

    /**
     * 设置当前行索引
     * @param int $row_index
     */
    protected function setRowIndex(int $row_index){
        $this->_row_index = $row_index;
    }

    protected function nextRow(){
        return $this->_row_index++;
    }


    /**
     * 拼装单元格的行列索引, 当 $row_index = 0 将默认取当前当行
     * @param int $column_index
     * @param int $row_index
     * @return string
     */
    public function getPcoordinate(int $column_index, int $row_index = 0){
        $column = isset($this->_column_index_arr[$column_index]) ? $this->_column_index_arr[$column_index] : Coordinate::stringFromColumnIndex($column_index);
        $row = intval($row_index) <=0 ? $this->_row_index : intval($row_index);

        return $column . $row;
    }

    /**
     * 拼装 指定区域的单元格索引
     * @param $columnIndex1
     * @param $row1
     * @param $columnIndex2
     * @param $row2
     * @return string
     */
    public function getPrange($columnIndex1, $row1, $columnIndex2, $row2){
        return $this->getPcoordinate($columnIndex1, $row1) . ':' . $this->getPcoordinate($columnIndex2, $row2);
    }

    /**
     * 拼装 指定同一行指定区域的单元格索引
     * @param $columnIndex1
     * @param $columnIndex2
     * @param $row1
     * @return string
     */
    public function getPrangeInRow($columnIndex1, $columnIndex2,  $row1){
        return $this->getPcoordinate($columnIndex1, $row1) . ':' . $this->getPcoordinate($columnIndex2, $row1);
    }

    public function getCoordinateDuringRange($range){
        if (strpos($range, ':') === false){
            return $range;
        }else{
            return explode(':', $range)[0];
        }
    }

    /**
     * web下载excel
     * @param Xlsx $writer
     * @param $filename
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    protected function download(Xlsx $writer, $filename){
        header('pragma:public');
        header("Content-Disposition:attachment;filename=$filename.xlsx");
        $writer->save('php://output');exit();
    }

    /**
     * 下载excel到本地
     * @param Xlsx $writer
     * @param $filename
     * @param $save_path
     * @return string
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    protected function saveLocalExcel(Xlsx $writer, $filename, $save_path){
        $filename = iconv("utf-8", "gb2312", $filename);//转码
        $file_path = $save_path . '/' . $filename.'.xlsx';
        $writer->save($file_path);
        return $file_path;
    }

    /**
     * @param Worksheet $activeSheet
     * @param $range
     */
    public function setStyle(Worksheet $activeSheet, $range){
        $activeSheet->getStyle($range)->getAlignment()->setHorizontal($this->getHorizontal());
        $activeSheet->getStyle($range)->getAlignment()->setVertical($this->getVertical());
    }

}