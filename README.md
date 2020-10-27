# excel op by phpexcel

### 安裝

先在composer.json中增加：

```json
"repositories": [
    {
        "type": "vcs",
        "url": "https://github.com/exinfinite/Excel.git"
    }
]
```
```php
composer require Exinfinite\Excel
```

### 初始化

```php
use Exinfinite\Excel;

$excel = new Excel();
```

### 使用

```php
/**
 * @param Exinfinite\Excel $excel
 * @param array $cols 每個直欄的設定及標題
 * @param array $data_array 每列資料
 * @param [type] $filename 檔名(不含副檔名)
 * @return void
 * 參數範例
 * $cols = [
        'A' => ['width' => 10, 'title' => '標題1'],
        'B' => ['width' => 25, 'title' => '標題2'],
        'C' => ['width' => 25, 'title' => '標題3'],
    ];
    $data_array = [
        ['資料1', '資料2', '資料3'],
        ['資料1', '資料2', '資料3']
    ];
 */
function excel(Exinfinite\Excel $excel, $cols = [], $data_array = [], $filename) {
    array_unshift($data_array, array_column($cols, 'title'));
    $excel->write($data_array, "A1");
    $excel->set_active_sheet(0);
    $act_sheet = $excel->objPHPExcel->getActiveSheet();
    foreach (array_combine(array_keys($cols), array_column($cols, 'width')) as $col => $w) {
        $act_sheet->getColumnDimension($col)->setWidth($w)->setAutoSize(false);
    }
    $highestCol = $act_sheet->getHighestDataColumn();
    $highestRow = $act_sheet->getHighestDataRow();
    $excel->style([
        'borders' => [
            'allborders' => [
                'style' => PHPExcel_Style_Border::BORDER_THIN,
            ],
        ],
        'alignment' => [
            'vertical' => PHPExcel_Style_Alignment::VERTICAL_CENTER,
            'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_LEFT,
            'wrap' => true,
        ],
    ], "A1", "{$highestCol}{$highestRow}");
    $excel->style([
        'fill' => [
            'type' => PHPExcel_Style_Fill::FILL_SOLID,
            'color' => ['argb' => 'D1EEEE'],
        ],
    ], "A1", "{$highestCol}1");
    header("Content-Type:application/vnd.ms-excel");
    header("Content-Disposition:attachment;filename={$filename}.xls");
    header("Cache-Control:max-age=0");
    return $excel->export();
}
```
