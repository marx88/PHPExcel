<?php

namespace mphp\excel\tests\something\lib;

use mphp\excel\imps\ArrayToExcelBase;

class ArrayToExcelDemo extends ArrayToExcelBase
{
    protected $tmpFilePath = __DIR__.'/../xls/export_demo.xls';

    protected $map = [
        'A' => 'name',
        'B' => 'sex',
        'C' => 'age',
    ];

    /**
     * 创建IExcelToArray实例.
     *
     * @param array|\ArrayAccess $list
     * @param array|\ArrayAccess $definedData
     *
     * @return static
     */
    public static function make($list, $definedData = null)
    {
        $obj = parent::make($list, $definedData);
        $obj->config->setSpecialCells(['C']);

        return $obj;
    }
}
