<?php

namespace mphp\excel;

use mphp\excel\exceptions\ExportException;
use PhpOffice\PhpSpreadsheet\Cell\DataType;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use PhpOffice\PhpSpreadsheet\Writer\IWriter;

/**
 * 根据模板导出excel.
 */
class Export
{
    const EXCEL_TYPE_2007 = 'Xlsx';

    const EXCEL_TYPE_2003 = 'Xls';

    const LIST_KEY = '__LIST__';

    const OHTER_DATA_KEY = '__OTHER_DATA__';

    /**
     * 模板
     *
     * @var Spreadsheet
     */
    protected $tpl;

    /**
     * 新的excel对象
     *
     * @var Spreadsheet
     */
    protected $clone;

    /**
     * 工作表.
     *
     * @var Worksheet
     */
    protected $sheet;

    /**
     * 模板路径.
     *
     * @var string
     */
    protected $path = '';

    /**
     * 缓存路径.
     *
     * @var string
     */
    protected $pathCache = '';

    /**
     * 模板类型.
     *
     * @var string
     */
    protected $type = '';

    /**
     * 错误信息.
     *
     * @var string
     */
    protected $error = '';

    /**
     * 字段前缀
     *
     * @var string
     */
    protected $prefix = '#';

    /**
     * 特殊类型列.
     *
     * 比如身份证需要是纯文本不然会被科学计数
     *
     * @var array
     */
    protected $specialCell = [];

    /**
     * 填充模板的数据.
     *
     * @var array
     */
    protected $data = [];

    /**
     * ExportExcelFromTemplate constructor.
     *
     * @param string $path       模板路径
     * @param string $type       模板类型 目前支持2003[Xls]和2007[Xlsx]
     * @param string $cache_path 缓存路径
     */
    public function __construct($path, $type = '', $cache_path = '')
    {
        $this->path = $path;
        $this->pathCache = pathinfo($cache_path ?: $path, PATHINFO_DIRNAME).DIRECTORY_SEPARATOR;

        $type = ucfirst($type);
        if (in_array($type, [self::EXCEL_TYPE_2003, self::EXCEL_TYPE_2007], true)) {
            $this->type = $type;
        } else {
            $this->type = self::EXCEL_TYPE_2003;
        }

        try {
            $this->tpl = IOFactory::load($path);
        } catch (\Exception $e) {
            $this->error = $e->getMessage();
        }
    }

    /**
     * 保存文件.
     *
     * @param array  $list
     * @param string $filename
     * @param array  $other_data
     *
     * @return bool|string
     */
    public function save($list, $filename = '', $other_data = [])
    {
        $temp_path = $this->pathCache.$this->parseFileName($filename);

        try {
            $this->data = [static::LIST_KEY => $list, static::OHTER_DATA_KEY => $other_data];

            $writer = $this->filling();
            $writer->setPreCalculateFormulas(false);

            $writer->save($temp_path);
        } catch (\Exception $e) {
            $this->error = $e->getMessage();

            return false;
        }

        return $temp_path;
    }

    /**
     * 导出.
     *
     * @param array  $list
     * @param string $filename
     * @param array  $other_data
     *
     * @throws ExportException
     *
     * @return mixed
     */
    public function download($list, $filename = '', $other_data = [])
    {
        try {
            $this->data = [static::LIST_KEY => $list, static::OHTER_DATA_KEY => $other_data];

            $writer = $this->filling();
            $writer->setPreCalculateFormulas(false);

            $filename = $filename ?: time();

            // ie、360极速下载文件名乱码
            if (preg_match('/(MSIE)|(Gecko)/', $_SERVER['HTTP_USER_AGENT'])) {
                $filename = urlencode($filename);
            }

            if ($this->type === static::EXCEL_TYPE_2003) {
                $filename .= '.xls';
                header('Content-Type: application/vnd.ms-excel');
            } else {
                $filename .= '.xlsx';
                header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
            }
            header('Content-Disposition: attachment;filename="'.$filename.'"');
            header('Cache-Control: max-age=0');

            return $writer->save('php://output');
        } catch (\Exception $e) {
            throw new ExportException($e);
        }
    }

    /**
     * @return string
     */
    public function getError()
    {
        return $this->error;
    }

    /**
     * @param string $prefix
     *
     * @return static
     */
    public function setPrefix($prefix)
    {
        if (is_string($prefix) && !empty($prefix)) {
            $this->prefix = $prefix;
        }

        return $this;
    }

    /**
     * @return static
     */
    public function setSpecialCell(array $specialCell)
    {
        $this->specialCell = $specialCell;

        return $this;
    }

    /**
     * 填充模板
     *
     * @throws ExportException
     *
     * @return IWriter
     */
    protected function filling()
    {
        if (is_null($this->tpl)) {
            throw new ExportException($this->getError() ?: "模板路径不存在:{$this->path}");
        }

        $this->clone = clone $this->tpl;
        $this->sheet = $this->clone->getActiveSheet();

        // 读取模板 多用于读、修改表头 例如：#year年 改成 2019年
        foreach ($this->sheet->getRowDimensions() as $y => $row) {
            foreach ($this->sheet->getColumnDimensions() as $x => $col) {
                $cell = $x.$y;
                $txt = $this->sheet->getCell($cell)->getValue();
                $txt = is_object($txt) ? $txt->__toString() : $txt;
                $txt = trim($txt);

                // 判断单元格内的值的前缀
                if (!$txt || 0 !== strpos($txt, $this->prefix)) {
                    continue;
                }

                // 判断other_data中是否有该key的值
                $key = str_replace($this->prefix, '', $txt);
                if (!isset($this->data[static::OHTER_DATA_KEY][$key])) {
                    continue;
                }

                // 重新设置该单元格的值
                $this->setValue($x, $y, $this->data[static::OHTER_DATA_KEY][$key]);
            }
        }

        // 处理列表数据
        $this->dealList();

        return IOFactory::createWriter($this->clone, $this->type);
    }

    /**
     * 处理列表数据 子类可重写 复杂业务在这里处理.
     *
     * @throws Exception
     */
    protected function dealList()
    {
        // 初始化开始行及字段映射
        $row_start = $this->sheet->getHighestRow();
        $col_end = $this->sheet->getHighestColumn($row_start);
        $map = [];
        foreach ($variable as $key => $value) {
            // code...
        }

        $row_cur = $row_start;
        foreach ($this->data[static::LIST_KEY] as $row) {
            // 初始化开始列
            $col_cur = 'A';

            foreach ($row as $key => $value) {
                // 填充数据
                $cell_cur = "{$col_cur}{$row_cur}";
                $this->setValue($col_cur, $row_cur, $value);

                // 拷贝样式
                $this->sheet->duplicateStyle($this->sheet->getStyle("{$col_cur}{$row_start}"), $cell_cur);

                // 列自增
                ++$col_cur;
            }

            // 行自增
            ++$row_cur;
        }
    }

    /**
     * 设置单元格的值，如果设置了特殊列或单元格 则使用特殊设置 如科学计数.
     *
     * @param string $col
     * @param string $row
     * @param mixed  $value
     */
    protected function setValue($col, $row, $value)
    {
        $cell = sprintf('%s%d', $col, $row);
        if (in_array($col, $this->specialCell, true) || in_array($cell, $this->specialCell, true)) {
            $this->sheet->setCellValueExplicit($cell, $value, DataType::TYPE_STRING);
        } else {
            $this->sheet->setCellValue($cell, $value);
        }
    }

    /**
     * 获取新文件名.
     *
     * @param string $name
     *
     * @return string
     */
    protected function parseFileName($name)
    {
        if (empty($name)) {
            list($name) = $this->randomName($name);
        } else {
            $name = pathinfo($name, PATHINFO_BASENAME);
        }

        return $name;
    }

    /**
     * 随机生成名称.
     *
     * @param string $filename
     *
     * @return array
     */
    protected function randomName($filename)
    {
        $data = [uniqid(), time(), date('Y-m-d'), md5(time()), md5(uniqid())];
        shuffle($data);

        $temp = md5(implode('#', $data));
        $filename = pathinfo($filename ?: $this->path, PATHINFO_BASENAME);
        $name = explode('.', $filename);

        return ["{$name[0]}_{$temp}.{$name[1]}", $filename];
    }
}
