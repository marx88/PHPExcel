<?php

namespace mphp\excel\tests\something\lib;

use mphp\excel\imps\ExcelToArrayBase;

class ExcelToArrayDemo extends ExcelToArrayBase
{
    protected $map = [
        'A' => 'name',
        'B' => 'sex',
        'C' => 'age',
    ];

    /**
     * 创建IExcelToArray实例.
     *
     * @return static
     */
    public static function make(string $filepath)
    {
        $obj = parent::make($filepath);
        $obj->config->setRowStart(2);
        $obj->config->setRowEnd(4);
        $obj->config->setColEnd('C');
        $obj->config->setMaxRowNum(2);

        return $obj;
    }

    /**
     * 用来验证测试结果.
     */
    public function getData()
    {
        return $this->sheetData;
    }
}
