<?php

namespace mphp\excel\export;

interface IArrayToExcel
{
    /**
     * 获取配置
     */
    public function getConfig(): Config;

    /**
     * 根据键名获取特殊值
     */
    public function getDefinedValueByKey(string $key): string;

    /**
     * 根据行索引、列名获取列表值
     */
    public function getListValue(string $columnName, int $rowIndex): string;

    /**
     * 列表已读取完毕.
     */
    public function isReadListOver(): bool;
}
