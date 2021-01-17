<?php

namespace mphp\excel\extract;

use PhpOffice\PhpSpreadsheet\Cell\Cell;

interface IExcelToArray
{
    /**
     * 返回配置信息.
     */
    public function getConfig(): Config;

    /**
     * 读取单元格数据.
     *
     * 读取每个单元格时调用该函数
     */
    public function readCell(Cell $cell): string;

    /**
     * 处理行数据.
     *
     * 每一行读取完时调用该函数
     *
     * 返回值：
     *   true：继续读下一行；
     *   false：中断读取；
     */
    public function afterReadRow(array &$row): bool;

    /**
     * 处理sheet数据.
     *
     * sheet读取完时调用该函数
     */
    public function afterReadSheet();
}
