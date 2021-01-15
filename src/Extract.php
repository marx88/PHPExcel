<?php

namespace mphp\excel;

use mphp\excel\exceptions\ExtractException;
use mphp\excel\imps\ExcelReadFilter;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Shared\Date;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

/**
 * excel的sheet转array类.
 */
class Extract
{
    // ================================= 属性 =================================
    /**
     * @var Spreadsheet
     */
    protected $exl;

    /**
     * @var Worksheet
     */
    protected $sheet;

    /**
     * sheet名称.
     *
     * @var int|string
     */
    protected $sheetName = 0;

    /**
     * excel列对应的数组键名.
     *
     * 例如：'A' => 'column_name'
     *
     * @var array
     */
    protected $map = [];

    /**
     * 每行默认值
     *
     * @var array
     */
    protected $default = [];

    /**
     * 从该行开始读取 默认1.
     *
     * @var int
     */
    protected $rowStart = 1;

    /**
     * 到该行结束读取 默认最后一行.
     *
     * @var int
     */
    protected $rowEnd;

    /**
     * 从该列开始读取 默认A.
     *
     * @var string
     */
    protected $colStart = 'A';

    /**
     * 到该列结束读取 默认最后一列.
     *
     * @var string
     */
    protected $colEnd;

    /**
     * 存放返回值
     *
     * @var array
     */
    protected $data = [];

    /**
     * 存放错误信息.
     *
     * @var array
     */
    protected $error = [];

    /**
     * 文件路径.
     *
     * @var string
     */
    protected $file = '';

    /**
     * excel列与数字互转的缓存.
     *
     * @var array
     */
    protected static $columnMapIndex = [[], []];

    /**
     * Extract constructor.
     *
     * @param string     $file
     * @param int|string $name
     */
    public function __construct($file, $name = null)
    {
        // 保存file路径
        $this->file = $file;

        // 保存sheet索引
        if (!is_null($name)) {
            $this->sheetName = $name;
        }
    }

    /**
     * Extract destructor.
     */
    public function __destruct()
    {
        if ($this->exl) {
            $this->exl->disconnectWorksheets();
        }
        $this->sheet = null;
        $this->exl = null;
        $this->data = null;
        $this->error = null;
    }

    /**
     * 执行 提取sheet数据到数组.
     *
     * @throws ExtractException
     * @throws \PhpOffice\PhpSpreadsheet\Reader\Exception
     *
     * @return $this
     */
    public function run()
    {
        if (is_null($this->initExcel())) {
            return $this;
        }

        // 开始循环读数据
        // 每一行填充
        $col_start = $this->parseColumnIndex($this->colStart);
        $col_end = $this->parseColumnIndex($this->colEnd);
        for ($row_cur = $this->rowStart; $row_cur <= $this->rowEnd; ++$row_cur) {
            $row = [];

            // 该行每一列填充
            for ($col_cur_int = $col_start; $col_cur_int <= $col_end; ++$col_cur_int) {
                // 计算单元格
                $col_cur = $this->parseColumnIndex($col_cur_int);
                $cell = $col_cur.$row_cur;

                // 读取单元格数据
                $val = $this->sheet->getCell($cell)->getCalculatedValue();
                $val = is_object($val) ? $val->__toString() : $val;
                $val = trim($val);

                // 获取字段名
                $column_name = isset($this->map[$col_cur]) ? $this->map[$col_cur] : $col_cur;

                // 设置值
                $row[$column_name] = $val;
            }

            // 空行略过
            if ('' === implode('', $row)) {
                continue;
            }

            // 添加默认字段值
            $row = array_merge($this->default, $row);

            // 行数据处理 子类继承
            $this->dealRow($row, $row_cur);
            if (!isset($row) || empty($row)) {
                continue;
            }

            $this->data[] = $row;
            unset($row);
        }

        return $this;
    }

    /**
     * @param null|string $colEnd
     *
     * @return $this
     */
    public function setColEnd($colEnd)
    {
        $this->colEnd = $colEnd;

        return $this;
    }

    /**
     * @param string $colStart
     *
     * @return $this
     */
    public function setColStart($colStart)
    {
        $this->colStart = $colStart;

        return $this;
    }

    /**
     * @return $this
     */
    public function setDefault(array $default)
    {
        $this->default = $default;

        return $this;
    }

    /**
     * @return $this
     */
    public function setMap(array $map)
    {
        $this->map = $map;

        return $this;
    }

    /**
     * @param null|int $rowEnd
     *
     * @return $this
     */
    public function setRowEnd($rowEnd)
    {
        $this->rowEnd = $rowEnd;

        return $this;
    }

    /**
     * @param int $rowStart
     *
     * @return $this
     */
    public function setRowStart($rowStart)
    {
        $this->rowStart = $rowStart;

        return $this;
    }

    /**
     * @return array
     */
    public function getData()
    {
        return $this->data;
    }

    /**
     * @return array
     */
    public function getError()
    {
        return $this->error;
    }

    /**
     * 加载excel.
     *
     * @throws ExtractException
     * @throws \PhpOffice\PhpSpreadsheet\Reader\Exception
     *
     * @return null|Spreadsheet
     */
    protected function initExcel()
    {
        if (is_null($this->exl)) {
            /** @var \PhpOffice\PhpSpreadsheet\Reader\BaseReader $reader */
            $type = IOFactory::identify($this->file);
            $reader = IOFactory::createReader($type);

            if ('CSV' === strtoupper($type) && $reader instanceof \PhpOffice\PhpSpreadsheet\Reader\Csv) {
                call_user_func([$reader, 'setInputEncoding'], 'GBK');
            }

            $reader->setReadDataOnly(true);

            if ($this->rowStart && $this->rowEnd) {
                $filter = new ExcelReadFilter();
                $filter->startRow = $this->rowStart;
                $filter->endRow = $this->rowEnd;
                $reader->setReadFilter($filter);
            }

            $this->exl = $reader->load($this->file);

            if ($this->exl) {
                // 初始化sheet
                $this->initSheet();

                // 初始化默认配置
                $this->initConfig();
            }
            unset($reader);
        }

        return $this->exl;
    }

    /**
     * 获取工作表.
     *
     * @throws ExtractException
     */
    protected function initSheet()
    {
        if (is_int($this->sheetName)) {
            // 根据索引获取sheet
            $this->sheet = $this->exl->getSheet($this->sheetName);
        } elseif (is_string($this->sheetName)) {
            // 根据name获取sheet
            $this->sheet = $this->exl->getSheetByName($this->sheetName);
        }

        if (!($this->sheet instanceof Worksheet)) {
            throw new ExtractException('文件格式错误');
        }
    }

    /**
     * 设置默认配置.
     */
    protected function initConfig()
    {
        if (is_null($this->rowEnd)) {
            $this->setRowEnd($this->sheet->getHighestRow());
        }
        if (is_null($this->colEnd)) {
            $this->setColEnd($this->sheet->getHighestColumn());
        }
    }

    /**
     * 数字转列 列转数字
     * 从0开始：0 => A, 1 => B, ... , 25 => Z, 26 => AA, ... 701 => ZZ, ...
     * 从1开始：1 => A, 2 => B, ... , 26 => Z, 27 => AA, ... 702 => ZZ, ...
     *
     * @param      $val
     * @param bool $from_zero A从1开始还是从0，默认从0
     *
     * @return float|int|string
     */
    protected function parseColumnIndex($val, $from_zero = true)
    {
        $key = $val;
        $index = true === $from_zero ? 0 : 1;
        if (!isset($this->columnMapIndex[$index][$key])) {
            if (is_int($val)) {
                $rest = '';
                $val = true === $from_zero ? $val : $val - 1;
                do {
                    $mod = intval($val % 26) + ('' === $rest ? 1 : 0);
                    $val = intval($val / 26);
                    $rest = chr(64 + $mod).$rest;
                } while ($val > 0);
            } else {
                $val = str_split(strtoupper($val));
                $num = count($val) - 1;
                $rest = true === $from_zero ? -1 : 0;
                foreach ($val as $item) {
                    $rest += (ord($item) - 64) * pow(26, $num--);
                }
            }

            $this->columnMapIndex[$index][$key] = $rest;
            $this->columnMapIndex[$index][$rest] = $key;
        }

        return $this->columnMapIndex[$index][$key];
    }

    /**
     * 行处理 可用来行数据格式化、数据检测 子类继承后可重写.
     *
     * @param array $row
     * @param int   $row_num
     */
    protected function dealRow(&$row, $row_num)
    {
    }

    // ================================= 个别格式提取 =================================

    /**
     * excel时间转php时间.
     *
     * @param mixed  $date
     * @param string $format
     *
     * @return string
     */
    protected function excelDateToPhp($date, $format = 'Y-m-d')
    {
        try {
            $time_temp = date($format, Date::excelToTimestamp($date));
        } catch (\Exception $e) {
            $time_temp = '';
        }

        if (is_string($time_temp) && '' !== $time_temp) {
            $date = $time_temp;
        }

        return $date;
    }
}
