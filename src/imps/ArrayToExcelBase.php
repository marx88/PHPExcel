<?php

namespace mphp\excel\imps;

use mphp\excel\export\Config;
use mphp\excel\export\IArrayToExcel;

class ArrayToExcelBase implements IArrayToExcel
{
    /**
     * @var Config
     */
    protected $config;

    /**
     * excel模板文件完整物理路径.
     *
     * @var string
     */
    protected $tmpFilePath;

    /**
     * @var int
     */
    protected $curListRowIndex = 0;

    /**
     * @var array|\ArrayAccess
     */
    protected $definedData;

    /**
     * @var array|\ArrayAccess
     */
    protected $list;

    /**
     * @var int
     */
    protected $listRowNum;

    /**
     * 列名与键名的映射.
     *
     * 例：'A' => 'column_name'
     *
     * @var array
     */
    protected $map = [];

    /**
     * 创建IArrayToExcel实例.
     *
     * @param array|\ArrayAccess $list
     * @param array|\ArrayAccess $definedData
     *
     * @return static
     */
    public static function make($list, $definedData = null)
    {
        $obj = new static();
        $obj->config = new Config();
        $obj->config->setTmpFilePath($obj->tmpFilePath);
        $obj->list = $list;
        $obj->definedData = is_null($definedData) ? [] : $definedData;
        $obj->listRowNum = count($obj->list);

        return $obj;
    }

    /**
     * 获取配置.
     */
    public function getConfig(): Config
    {
        return $this->config;
    }

    /**
     * 根据键名获取特殊值
     */
    public function getDefinedValueByKey(string $key): string
    {
        return isset($this->definedData[$key]) ? $this->definedData[$key] : '';
    }

    /**
     * 根据行索引、列名获取列表值
     */
    public function getListValue(string $columnName, int $rowIndex): string
    {
        $this->curListRowIndex = $rowIndex;

        if (!isset($this->list[$rowIndex])) {
            $this->curListRowIndex = $this->listRowNum;

            return '';
        }

        if (isset($this->map[$columnName])) {
            $columnName = $this->map[$columnName];
        }

        if (!isset($this->list[$rowIndex][$columnName])) {
            return '';
        }

        return $this->list[$rowIndex][$columnName];
    }

    /**
     * 列表已读取完毕.
     */
    public function isReadListOver(): bool
    {
        return $this->curListRowIndex + 1 >= $this->listRowNum;
    }
}
