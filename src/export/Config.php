<?php

namespace mphp\excel\export;

use mphp\excel\Constants;

class Config
{
    /**
     * excel模板文件完整物理路径.
     *
     * @var string
     */
    protected $tmpFilePath;

    /**
     * 特殊字段前缀.
     *
     * @var string
     */
    protected $prefix = '#';

    /**
     * 导出文件名.
     *
     * @var string
     */
    protected $exportFileName = '';

    /**
     * 特殊单元格.
     *
     * @var array
     */
    protected $specialCells = [];

    /**
     * Excel文件类型.
     *
     * @var string
     */
    protected $excelType = Constants::EXCEL_TYPE_2003;

    /**
     * @return $this
     */
    public function setTmpFilePath(string $tmpFilePath)
    {
        $this->tmpFilePath = $tmpFilePath;

        return $this;
    }

    /**
     * @return string
     */
    public function getTmpFilePath()
    {
        return $this->tmpFilePath;
    }

    /**
     * @return $this
     */
    public function setPrefix(string $prefix)
    {
        $this->prefix = $prefix;

        return $this;
    }

    /**
     * @return string
     */
    public function getPrefix()
    {
        return $this->prefix;
    }

    /**
     * @return $this
     */
    public function setExportFileName(string $exportFileName)
    {
        $this->exportFileName = $exportFileName;

        return $this;
    }

    /**
     * @return string
     */
    public function getExportFileName()
    {
        return $this->exportFileName;
    }

    /**
     * @return $this
     */
    public function setSpecialCells(array $specialCells)
    {
        $this->specialCells = $specialCells;

        return $this;
    }

    /**
     * @return array
     */
    public function getSpecialCells()
    {
        return $this->specialCells;
    }

    /**
     * @return $this
     */
    public function setExcelType(string $excelType)
    {
        $this->excelType = $excelType;

        return $this;
    }

    /**
     * @return string
     */
    public function getExcelType()
    {
        return $this->excelType;
    }
}
