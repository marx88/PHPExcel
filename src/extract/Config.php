<?php

namespace mphp\excel\extract;

class Config
{
    /**
     * excel文件完整物理路径.
     *
     * @var string
     */
    protected $filepath;

    /**
     * sheet名称.
     *
     * @var int|string
     */
    protected $sheetName = 0;

    /**
     * 从该行开始读取 默认1.
     *
     * @var int
     */
    protected $rowStart = 1;

    /**
     * 到该行结束读取 默认最后一行.
     *
     * @var int
     */
    protected $rowEnd;

    /**
     * 从该列开始读取 默认A.
     *
     * @var string
     */
    protected $colStart = 'A';

    /**
     * 到该列结束读取 默认最后一列.
     *
     * @var string
     */
    protected $colEnd = 'Z';

    /**
     * 最大读取行数.
     *
     * @var int
     */
    protected $maxRowNum = 1000;

    /**
     * @param string $filepath
     *
     * @return $this
     */
    public function setFilepath($filepath)
    {
        $this->filepath = $filepath;

        return $this;
    }

    /**
     * @return null|string
     */
    public function getFilepath()
    {
        return $this->filepath;
    }

    /**
     * @param null|int|string $sheetName
     *
     * @return $this
     */
    public function setSheetName($sheetName)
    {
        $this->sheetName = $sheetName;

        return $this;
    }

    /**
     * @return null|int|string
     */
    public function getSheetName()
    {
        return $this->sheetName;
    }

    /**
     * @param null|string $colEnd
     *
     * @return $this
     */
    public function setColEnd($colEnd)
    {
        $this->colEnd = $colEnd;

        return $this;
    }

    /**
     * @return null|string
     */
    public function getColEnd()
    {
        return $this->colEnd;
    }

    /**
     * @param string $colStart
     *
     * @return $this
     */
    public function setColStart($colStart)
    {
        $this->colStart = $colStart;

        return $this;
    }

    /**
     * @return string
     */
    public function getColStart()
    {
        return $this->colStart;
    }

    /**
     * @param null|int $rowEnd
     *
     * @return $this
     */
    public function setRowEnd($rowEnd)
    {
        $this->rowEnd = $rowEnd;

        return $this;
    }

    /**
     * @return null|int
     */
    public function getRowEnd()
    {
        return $this->rowEnd;
    }

    /**
     * @param int $rowStart
     *
     * @return $this
     */
    public function setRowStart($rowStart)
    {
        $this->rowStart = $rowStart;

        return $this;
    }

    /**
     * @return int
     */
    public function getRowStart()
    {
        return $this->rowStart;
    }

    /**
     * @param int $maxRowNum
     *
     * @return $this
     */
    public function setMaxRowNum($maxRowNum)
    {
        $this->maxRowNum = $maxRowNum;

        return $this;
    }

    /**
     * @return int
     */
    public function getMaxRowNum()
    {
        return $this->maxRowNum;
    }
}
