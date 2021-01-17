<?php

namespace mphp\excel\tests\imps;

use mphp\excel\export\Export;
use mphp\excel\tests\something\lib\ArrayToExcelDemo;
use PHPUnit\Framework\TestCase;

/**
 * @internal
 * @coversNothing
 */
class ArrayToExcelTest extends TestCase
{
    public function testExport()
    {
        $cases = [
            [
                'list' => [
                    [
                        'name' => '小明',
                        'sex' => '男',
                        'age' => '18',
                    ],
                    [
                        'name' => '小红',
                        'sex' => '女',
                        'age' => '16',
                    ],
                    [
                        'name' => '李雷',
                        'sex' => '',
                        'age' => '',
                    ],
                ],
                'definedData' => [
                    'year' => '2021',
                ],
                'savePath' => __DIR__.'/../something/cache/export_demo_'.date('YmdHis').'.xls',
            ],
        ];

        foreach ($cases as $case) {
            $obj = ArrayToExcelDemo::make($case['list'], $case['definedData']);
            $export = new Export($obj);
            $export->save($case['savePath']);
            $this->assertFileExists($case['savePath']);
        }
    }
}
