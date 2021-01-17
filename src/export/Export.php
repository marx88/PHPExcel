<?php

namespace mphp\excel\export;

use mphp\excel\Constants;
use mphp\excel\exceptions\ExportException;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Cell\DataType;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use PhpOffice\PhpSpreadsheet\Writer\IWriter;

/**
 * 导出array到excel.
 *
 * 使用示例：Export::run(ArrayToExcelBase::make($list));
 */
class Export
{
    /**
     * @var Config
     */
    protected $config;

    /**
     * @var IArrayToExcel
     */
    protected $arrayToExcel;

    /**
     * Export constructor.
     */
    public function __construct(IArrayToExcel $arrayToExcel)
    {
        $config = $arrayToExcel->getConfig();
        if (!$config->getTmpFilePath()) {
            throw new ExportException('缺少excel模板文件路径参数');
        }

        if (!$config->getExportFileName()) {
            $config->setExportFileName(time());
        }

        $this->arrayToExcel = $arrayToExcel;
        $this->config = $config;
    }

    /**
     * 下载导出的excel.
     */
    public static function run(IArrayToExcel $arrayToExcel)
    {
        $obj = new static($arrayToExcel);
        $obj->download();
    }

    /**
     * 下载导出.
     *
     * @return mixed
     */
    public function download()
    {
        $writer = $this->filling();
        $writer->setPreCalculateFormulas(false);
        $this->setRespHeaders($this->config->getExcelType());

        return $writer->save('php://output');
    }

    /**
     * 保存导出.
     *
     * 此函数估计没用，现在写是为了测试filling函数
     *
     * @return mixed
     */
    public function save(string $filepath)
    {
        $writer = $this->filling();
        $writer->setPreCalculateFormulas(false);

        return $writer->save($filepath);
    }

    /**
     * 填充数据.
     */
    protected function filling(): IWriter
    {
        $tmpFilePath = $this->config->getTmpFilePath();
        $excel = clone IOFactory::load($tmpFilePath);
        $sheet = $excel->getActiveSheet();
        $highestArr = $sheet->getHighestRowAndColumn();
        $rowNum = $highestArr['row'];
        $colNum = Coordinate::columnIndexFromString($highestArr['column']);
        $this->fillingDefinedCells($sheet, $rowNum, $colNum);
        $this->fillingList($sheet, $rowNum + 1, $colNum);

        return IOFactory::createWriter($excel, $this->config->getExcelType());
    }

    /**
     * 填充特殊单元格数据.
     */
    protected function fillingDefinedCells(WorkSheet $sheet, int $rowNum, int $colNum)
    {
        $prefix = $this->config->getPrefix();
        $specialCells = $this->config->getSpecialCells();
        for ($rowCur = 1; $rowCur <= $rowNum; ++$rowCur) {
            for ($colCur = 1; $colCur <= $colNum; ++$colCur) {
                $colCurStr = Coordinate::stringFromColumnIndex($colCur);
                $cell = $colCurStr.$rowCur;
                $txt = $sheet->getCell($cell)->getValue();
                $txt = is_object($txt) ? $txt->__toString() : $txt;
                $txt = trim($txt);

                // 判断单元格内的值的前缀
                if (!$txt || 0 !== strpos($txt, $prefix)) {
                    continue;
                }

                // 根据key获取值
                $key = str_replace($prefix, '', $txt);
                $val = $this->arrayToExcel->getDefinedValueByKey($key);

                // 重新设置该单元格的值
                if (in_array($cell, $specialCells, true)) {
                    $sheet->setCellValueExplicit($cell, $val, DataType::TYPE_STRING);
                } else {
                    $sheet->setCellValue($cell, $val);
                }
            }
        }
    }

    /**
     * 填充列表数据.
     */
    protected function fillingList(WorkSheet $sheet, int $startRow, int $colNum)
    {
        $specialCells = $this->config->getSpecialCells();
        $listIndex = 0;
        while (!$this->arrayToExcel->isReadListOver()) {
            $rowCur = $startRow + $listIndex;

            // 超过最大行则结束
            if ($rowCur > Constants::MAX_ROW_NUM) {
                break;
            }

            for ($colCur = 1; $colCur <= $colNum; ++$colCur) {
                $colCurStr = Coordinate::stringFromColumnIndex($colCur);
                $cell = $colCurStr.$rowCur;
                $val = $this->arrayToExcel->getListValue($colCurStr, $listIndex);

                // 设置值
                if (in_array($colCurStr, $specialCells, true)) {
                    $sheet->setCellValueExplicit($cell, $val, DataType::TYPE_STRING);
                } else {
                    $sheet->setCellValue($cell, $val);
                }

                // 拷贝样式
                $sheet->duplicateStyle($sheet->getStyle("{$colCurStr}{$startRow}"), $cell);
            }
            ++$listIndex;
        }
    }

    /**
     * 设置响应头.
     */
    protected function setRespHeaders(string $excelType)
    {
        $filename = $this->config->getExportFileName();

        // ie、360极速下载文件名乱码
        if (isset($_SERVER['HTTP_USER_AGENT']) && preg_match('/(MSIE)|(Gecko)/', $_SERVER['HTTP_USER_AGENT'])) {
            $filename = urlencode($filename);
        }

        if (Constants::EXCEL_TYPE_2003 === $excelType) {
            $filename .= '.xls';
            header('Content-Type: application/vnd.ms-excel');
        } else {
            $filename .= '.xlsx';
            header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        }
        header('Content-Disposition: attachment;filename="'.$filename.'"');
        header('Cache-Control: max-age=0');
    }
}
