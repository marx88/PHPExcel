<?php

namespace mphp\excel;

class Parse
{
    private static $map = [];

    /**
     * 列名转数字.
     *
     * @param string $col
     *
     * @return int
     */
    public static function colToNum($col)
    {
        if (!isset(static::$map[$col])) {
            $col_list = str_split(strtoupper($col));
            $num = count($col_list) - 1;
            $rest = -1;
            foreach ($col_list as $val) {
                $rest += (ord($val) - 64) * pow(26, $num--);
            }
            static::$map[$col] = $rest;
        }

        return static::$map[$col];
    }

    /**
     * 数字转列名.
     *
     * @param int $num
     *
     * @return string
     */
    public static function numtoCol($num)
    {
        if (!isset(static::$map[$num])) {
            $rest = '';
            do {
                $mod = intval($num % 26) + ('' === $rest ? 1 : 0);
                $num = intval($num / 26);
                $rest = chr(64 + $mod).$rest;
            } while ($num > 0);
            static::$map[$num] = $rest;
        }

        return static::$map[$num];
    }

    /**
     * 列名自增.
     *
     * @param string $col
     *
     * @return string
     */
    public static function selfInc($col)
    {
        $num = static::colToNum($col);

        return static::numtoCol(++$num);
    }
}
