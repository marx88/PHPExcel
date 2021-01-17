<?php

namespace mphp\excel\tests\imps;

use mphp\excel\extract\Extract;
use mphp\excel\tests\something\lib\ExcelToArrayDemo;
use PHPUnit\Framework\TestCase;

/**
 * @internal
 * @coversNothing
 */
class ExcelToArrayTest extends TestCase
{
    public function testExtract()
    {
        $cases = [
            [
                'filepath' => __DIR__.'/../something/xls/extract_demo.xls',
                'expected' => [
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
            ],
        ];

        foreach ($cases as $case) {
            $obj = ExcelToArrayDemo::make($case['filepath']);
            Extract::run($obj);
            $rest = $obj->getData();
            $this->assertSame($case['expected'], $rest);
        }
    }
}
