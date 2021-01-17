<?php

namespace mphp\excel\imps;

use mphp\excel\extract\Config;
use mphp\excel\extract\IExcelToArray;
use PhpOffice\PhpSpreadsheet\Cell\Cell;

class ExcelToArrayBase implements IExcelToArray
{
    /**
     * 保存sheet数据.
     *
     * @var array
     */
    protected $sheetData = [];

    /**
     * @var Config
     */
    protected $config;

    /**
     * 列名与键名的映射.
     *
     * 例：'A' => 'column_name'
     *
     * @var array
     */
    protected $map = [];

    /**
     * 创建IExcelToArray实例.
     *
     * @return static
     */
    public static function make(string $filepath)
    {
        $obj = new static();
        $obj->config = new Config();
        $obj->config->setFilepath($filepath);

        return $obj;
    }

    /**
     * 返回配置信息.
     */
    public function getConfig(): Config
    {
        return $this->config;
    }

    /**
     * 读取单元格数据.
     *
     * 读取每个单元格时调用该函数
     */
    public function readCell(Cell $cell): string
    {
        $val = $cell->getCalculatedValue();
        if (is_object($val)) {
            $val = $val->__toString();
        }

        return trim($val);
    }

    /**
     * 处理行数据.
     *
     * 每一行读取完时调用该函数
     *
     * 返回值：
     *   true：继续读下一行；
     *   false：中断读取；
     */
    public function afterReadRow(array &$row): bool
    {
        foreach ($this->map as $columnName => $keyName) {
            if (array_key_exists($columnName, $row)) {
                $row[$keyName] = $row[$columnName];
                unset($row[$columnName]);
            }
        }

        // 只保存非空行
        if ('' !== implode('', $row)) {
            array_push($this->sheetData, $row);
        }

        return true;
    }

    /**
     * 处理sheet数据.
     *
     * sheet读取完时调用该函数
     */
    public function afterReadSheet()
    {
    }

    /**
     * 设置Map.
     */
    public function setMap(array $map)
    {
        $this->map = $map;
    }

    /**
     * 获取Sheet数据.
     */
    public function getSheetData()
    {
        return $this->sheetData;
    }
}
