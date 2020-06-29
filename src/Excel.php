<?php
/**
 *By Exinfinite
 *引入檔案的header前不能有輸出
 */
namespace Exinfinite;
use \PHPExcel;
use \PHPExcel_IOFactory;

class Excel {
    var $objPHPExcel = null;
    var $objActSheet = null;
    var $activeSheetIndex = 0;
    function __construct() {
        if (!isset($this->objPHPExcel)) {
            $this->objPHPExcel = new PHPExcel();
            $this->objActSheet = $this->objPHPExcel->getActiveSheet();
        }
    }
    /**
     * 將標題和fetchAll()的欄位做對應結合
     * @param  Array  $title [DB欄位對應的標題名稱]
     * @param  Array  $data_all  [二維陣列,fetchAll()出來的結果]
     */
    function col_match(Array $title, Array $data_all) {
        $data = array();
        array_push($data, array_values($title));
        foreach ($data_all as $cols) {
            $tmp_data = array();
            foreach ($title as $k => $v) {
                array_push($tmp_data, (string) strip_tags($cols[$k]));
            }
            array_push($data, $tmp_data);
        }
        return $data;
    }
    //指定特定欄位樣式
    function style(Array $style, $from_col = null, $to_col = null) {
        if (!isset($from_col) || !isset($to_col)) {
            $this->objPHPExcel->getDefaultStyle()->applyFromArray($style);
        } else {
            $this->objActSheet->getStyle("{$from_col}:{$to_col}")->applyFromArray($style);
        }
        return $this;
    }
    //指定整欄樣式
    function style_whole_col(Array $style, $col = null) {
        if (!isset($col)) {
            $this->style($style);
        } else {
            $lastrow = $this->objActSheet->getHighestRow();
            $this->objActSheet->getStyle("{$col}1:{$col}{$lastrow}")->applyFromArray($style);
        }
        return $this;
    }
    //依篩選條件指定整欄樣式
    function condition_style_whole_col(Array $style, $col = null) {
        if (!isset($col)) {
            $this->objActSheet->duplicateConditionalStyle($style);
        } else {
            $lastrow = $this->objActSheet->getHighestRow();
            $this->objActSheet->duplicateConditionalStyle($style, "{$col}1:{$col}{$lastrow}");
        }
        return $this;
    }
    /**
     * 寫入
     * @param  Array  $data [二維陣列]
     * @param  String  $start_col [起始欄位]
     */
    function write(Array $data, $start_col = "A1") {
        $this->get_active_sheet();
        $this->objActSheet->fromArray($data, null, $start_col);
        $highestCol = $this->objActSheet->getHighestDataColumn();
        foreach (range('A', $highestCol) as $column) {
            $this->objActSheet->getColumnDimension($column)->setAutoSize(true);
        }
        return $this;
    }
    /**
     * 增加工作表
     * @param String $title [工作表標題]
     * @param int $idx [工作表插入位置]
     */
    function create_sheet($title = "", $idx = null) {
        $sheet = $this->objPHPExcel->createSheet($idx);
        $sheet->setTitle($title);
        return $this;
    }
    /**
     * 設定預設顯示的工作表
     * @param int $idx [工作表id]
     */
    function set_active_sheet($idx) {
        $this->activeSheetIndex = (int) $idx;
        $this->objPHPExcel->setActiveSheetIndex($this->activeSheetIndex);
        return $this;
    }
    /**
     * 取得目前工作表
     * @param int $idx [工作表id]
     */
    function get_active_sheet() {
        $this->objActSheet = $this->objPHPExcel->getActiveSheet();
        return $this;
    }
    /**
     * 設定標題
     * @param String $title [工作表標題]
     */
    function set_title($title) {
        $this->objActSheet->setTitle($title);
        return $this;
    }
    /**
     * 輸出
     */
    function export($type = 'Excel5', $path = false) {
        trim(ob_get_clean());
        $objWriter = PHPExcel_IOFactory::createWriter($this->objPHPExcel, $type);
        if ($type == 'csv') {
            $objWriter->setDelimiter(",");
        }
        if ($path === false) {
            return $objWriter->save('php://output');
        }
        $objWriter->save($path);
    }
    /**
     * 上傳excel檔,並輸出陣列
     */
    function upload($file, $rule = array(), $highestCol = -1, $highestRow = -1) {
        try {
            if ($file["tmp_name"] != "") {
                $file_ext = explode(".", $file["name"]);
                $file_ext = strtolower($file_ext[count($file_ext) - 1]);
                $ext_permit = array("xls", "xlsx");
                if (!in_array($file_ext, $ext_permit)) { //驗證副檔名未過
                    return false;
                }
                set_time_limit(90);
                ini_set("memory_limit", "50M");
                ini_set('precision', '30');
                $now_date = date("Y-m-d H:i:s"); //匯入時間
                $objPHPExcel = PHPExcel_IOFactory::load($file["tmp_name"]); //建立excel物件
                $sheet = $objPHPExcel->getActiveSheet(); //取得工作表
                $highestCol = $sheet->getHighestDataColumn(); //最後有資料的一欄
                $highestRow = $sheet->getHighestDataRow(); //最後有資料的一列
                /**
                 * 比對xls格式
                 */
                foreach ($rule as $k => $v) {
                    if (trim($sheet->getCell($k)->getValue()) != $v) {
                        return false;
                    }
                }
                $sheetData = $sheet->rangeToArray(
                    "A1:{$highestCol}{$highestRow}"
                );
                return $sheetData;
            }
            return false;
        } catch (Exception $e) {
            echo "error：" . $e->getMessage();
        }
    }
}
?>