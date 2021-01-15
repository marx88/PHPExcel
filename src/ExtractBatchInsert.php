<?php

namespace mphp\excel;

/**
 * 批量插入.
 *
 * Class ExtractBatchInsert
 */
class ExtractBatchInsert extends Extract
{
    protected $map = [
        'B' => 'title',
        'C' => 'type_code',
        'D' => 'answer',
        'E' => 'option_a',
        'F' => 'option_b',
        'G' => 'option_c',
        'H' => 'option_d',
        'I' => 'option_e',
        'J' => 'option_f',
        'K' => 'option_g',
    ];

    /**
     * 批量插入.
     *
     * @return string
     */
    public function insertAll()
    {
        if (empty($this->data)) {
            return 0;
        }

        $name = 'table_name';
        $column = '(`job_type_id`, `title`, `option`, `type_code`, `answer`, `create_time`, `update_time`)';
        $val = implode(',', $this->data);

        return sprintf('%s %s %s %s %s %s;', 'insert', 'into', $name, $column, 'values', $val);
    }

    public static function batchImport($file, $default = [], $start_row = 2, $max_row = 10000)
    {
        set_time_limit(600);

        $fn_import = function () use (&$file, &$default, &$start_row, &$max_row) {
            $end_row = $start_row + $max_row;
            $right_num = 0;
            $error = [];

            while (true) {
                $excel = new static($file);
                $excel->setDefault($default);
                $excel->setRowStart($start_row);
                $excel->setRowEnd($end_row);
                $count = $excel->run()->insertAll();

                // 没有新增行且没有错误行 中断循环
                $error_excel = $excel->getError();
                if ($count < 1 && empty($error_excel)) {
                    break;
                }

                $right_num += $count; // 统计有效行数
                $start_row = $end_row + 1; // 开始行
                $end_row = $start_row + $max_row; // 结束行
                $error = array_merge($error, $error_excel); // 记录问题行

                $excel = null;
            }

            return [$right_num, $error];
        };

        return $fn_import;
    }

    protected function dealRow(&$row, $row_num)
    {
        // 判空
        if (
            !isset($row['title']) || '' === $row['title']
            || !isset($row['type_code'])
            || '' === $row['type_code']
            || !isset($row['answer'])
            || '' === $row['answer']
        ) {
            $this->error[] = sprintf('第%s行数据不完整', $row_num);
            $row = null;

            return null;
        }

        // 判类型
        if (false !== strpos($row['type_code'], '单')) {
            $row['type_code'] = 1;
        } elseif (false !== strpos($row['type_code'], '多')) {
            $row['type_code'] = 2;
        } elseif (false !== strpos($row['type_code'], '判')) {
            $row['type_code'] = 3;
        } else {
            $this->error[] = sprintf('第%s行题型必须为：单项选择、多项选择、判断题中的一项', $row_num);
            $row = null;

            return null;
        }

        // 判选项
        $row['option'] = [];
        foreach ($this->map as $val) {
            if (false === strpos($val, 'option_')) {
                continue;
            }

            if (!isset($row[$val]) || '' === $row[$val]) {
                continue;
            }

            $opt_name = strtoupper(str_replace('option_', '', $val));
            $row['option'][$opt_name] = $row[$val];
        }
        if (empty($row['option'])) {
            $this->error[] = sprintf('第%s行数据不完整', $row_num);
            $row = null;

            return null;
        }

        $row['option'] = json_encode($row['option'], JSON_UNESCAPED_UNICODE);

        $row = sprintf(
            "(%s, '%s', '%s', %s, '%s', now(), now())",
            $row['job_type_id'],
            $row['title'],
            $row['option'],
            $row['type_code'],
            $row['answer']
        );
    }
}
