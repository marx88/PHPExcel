<?php

namespace mphp\excel\imps;

use PhpOffice\PhpSpreadsheet\Reader\IReadFilter;

/**
 * excel读取过滤器.
 */
class ExcelReadFilter implements IReadFilter
{
    public $startRow = 1;

    public $endRow;

    public function readCell($column, $row, $worksheetName = '')
    {
        //如果endRow没有设置表示读取全部
        if (!$this->endRow) {
            return true;
        }

        //只读取指定的行
        if ($row >= $this->startRow && $row <= $this->endRow) {
            return true;
        }

        return false;
    }
}
