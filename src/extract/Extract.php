<?php

namespace mphp\excel\extract;

use mphp\excel\Constants;
use mphp\excel\exceptions\ExtractException;
use mphp\excel\imps\ExcelReadFilter;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Reader\Csv;
use PhpOffice\PhpSpreadsheet\Reader\IReader;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

/**
 * 提取sheet数据到array.
 *
 * 使用示例：Extract::run(ExcelToArrayBase::make('filepath'));
 */
class Extract
{
    /**
     * @var IExcelToArray
     */
    protected $excelToArray;

    /**
     * @var Config
     */
    protected $config;

    /**
     * Extract constructor.
     */
    public function __construct(IExcelToArray $excelToArray)
    {
        $config = $excelToArray->getConfig();
        if (!$config || !$config->getFilepath()) {
            throw new ExtractException('缺少excel文件路径参数');
        }

        if (!is_int($config->getRowStart())) {
            $config->setRowStart(1);
        }

        if (!is_int($config->getRowEnd())) {
            $config->setRowEnd(Constants::MAX_ROW_NUM);
        }

        if (!$config->getColStart()) {
            $config->setColStart('A');
        }

        if (!$config->getColEnd()) {
            $config->setColEnd('Z');
        }

        if (!$config->getMaxRowNum()) {
            $config->setMaxRowNum(1000);
        }

        $this->excelToArray = $excelToArray;
        $this->config = $config;
    }

    /**
     * 读取sheet到array.
     */
    public static function run(IExcelToArray $excelToArray)
    {
        $obj = new static($excelToArray);
        $obj->readExcel();
    }

    /**
     * 读取excel数据.
     */
    public function readExcel()
    {
        $type = IOFactory::identify($this->config->getFilepath());
        $rowMin = $this->config->getRowStart();
        $rowMax = $this->config->getRowEnd();
        $colStart = Coordinate::columnIndexFromString($this->config->getColStart());
        $colEnd = Coordinate::columnIndexFromString($this->config->getColEnd());
        $maxRowNum = $this->config->getMaxRowNum();

        $rowStart = $rowMin;
        $rowEnd = $rowStart + $maxRowNum - 1;
        while ($rowStart <= $rowMax) {
            $reader = $this->getReader($type, $rowStart, $rowEnd);
            $excel = $reader->load($this->config->getFilepath());
            $sheet = $this->getWorkSheet($excel);
            $break = false;

            for ($rowCur = $rowStart; $rowCur <= $rowEnd; ++$rowCur) {
                $row = [];

                for ($colCur = $colStart; $colCur <= $colEnd; ++$colCur) {
                    // 计算单元格
                    $colCurStr = Coordinate::stringFromColumnIndex($colCur);
                    $cell = $colCurStr.$rowCur;

                    // 读取单元格数据
                    $row[$colCurStr] = $this->excelToArray->readCell($sheet->getCell($cell));
                }

                // 行读取完时
                if (false === $this->excelToArray->afterReadRow($row)) {
                    $break = true;

                    break;
                }
            }

            $rowStart = $rowEnd + 1;
            $rowEnd = $rowStart + $maxRowNum - 1;
            $reader = null;
            $sheet = null;
            $excel->disconnectWorksheets();
            $excel = null;

            if (true === $break) {
                break;
            }
        }

        // sheet读取完时
        $this->excelToArray->afterReadSheet();
    }

    /**
     * 获取excel的reader.
     *
     * @param string $type
     * @param int    $rowStart
     * @param int    $rowEnd
     */
    protected function getReader($type, $rowStart, $rowEnd): IReader
    {
        $reader = IOFactory::createReader($type);

        if ('CSV' === strtoupper($type) && $reader instanceof Csv) {
            call_user_func([$reader, 'setInputEncoding'], 'GBK');
        }

        $reader->setReadDataOnly(true);
        $reader->setLoadSheetsOnly($this->config->getSheetName());

        $filter = new ExcelReadFilter();
        $filter->startRow = $rowStart;
        $filter->endRow = $rowEnd;
        $reader->setReadFilter($filter);

        return $reader;
    }

    /**
     * 初始化sheet.
     *
     * @throws ExtractException
     *
     * @return WorkSheet
     */
    protected function getWorkSheet(Spreadsheet $excel)
    {
        $sheetName = $this->config->getSheetName();
        if (is_int($sheetName)) {
            // 根据索引获取sheet
            $workSheet = $excel->getSheet($sheetName);
        } elseif (is_string($sheetName)) {
            // 根据name获取sheet
            $workSheet = $excel->getSheetByName($sheetName);
        }

        if (!($workSheet instanceof Worksheet)) {
            throw new ExtractException('文件格式错误');
        }

        return $workSheet;
    }
}
