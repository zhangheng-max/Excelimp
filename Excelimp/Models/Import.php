<?php namespace Phpcmf\Model\Excelimp;

/**
 * Excel数据导入插件 - 导入模型
 * 适用于迅睿CMS 4.7.3
 *
 * 使用系统自带 PhpSpreadsheet 解析 Excel/CSV，
 * 通过官方 save_content() 接口写入模块内容（含索引表、主表、附表、栏目数据、URL生成）。
 */

class Import extends \Phpcmf\Model {

    // 允许的文件扩展名
    protected $allow_ext = ['xlsx', 'xls', 'csv'];

    // 允许导入的模块字段类型（其他类型如联动、单选组等按文本处理）
    protected $allow_fieldtype = [
        'Text', 'Textarea', 'Textbtn', 'Ueditor', 'Editor', 'Date', 'Time',
        'Number', 'Touchspin', 'Integer', 'File', 'Files', 'Image', 'Images',
        'Select', 'Selects', 'Radio', 'Checkbox', 'Cat', 'Cats', 'Linkage', 'Linkages',
    ];

    /**
     * 读取Excel文件
     * @param string $file 文件绝对路径
     * @return array dr_return_data格式
     */
    public function read_excel($file) {

        if (!is_file($file)) {
            return dr_return_data(0, dr_lang('文件不存在'));
        }

        $ext = strtolower(pathinfo($file, PATHINFO_EXTENSION));
        if (!in_array($ext, $this->allow_ext)) {
            return dr_return_data(0, dr_lang('文件格式不支持，仅支持xlsx/xls/csv'));
        }

        // 引入系统自带PhpSpreadsheet
        $vendor = FCPATH.'App/Vendor/autoload.php';
        if (!is_file($vendor)) {
            return dr_return_data(0, dr_lang('系统缺少PhpSpreadsheet组件'));
        }
        require_once $vendor;

        try {
            if ($ext == 'csv') {
                $reader = new \PhpOffice\PhpSpreadsheet\Reader\Csv();
                // 自动识别编码（兼容GBK导出的csv）
                $enc = $this->_detect_encoding($file);
                if ($enc) {
                    $reader->setInputEncoding($enc);
                }
                // 自动识别分隔符
                $reader->setDelimiter($this->_detect_delimiter($file));
                $reader->setEnclosure('"');
                $spreadsheet = $reader->load($file);
            } else {
                $reader = \PhpOffice\PhpSpreadsheet\IOFactory::createReaderForFile($file);
                // 需要读取样式信息，用于识别“日期格式”单元格
                // （readDataOnly会丢失格式，导致日期列读成Excel序列号数字）
                $reader->setReadDataOnly(false);
                $spreadsheet = $reader->load($file);
            }
        } catch (\Throwable $e) {
            return dr_return_data(0, dr_lang('文件解析失败：%s', $e->getMessage()));
        }

        $sheet = $spreadsheet->getActiveSheet();

        // 识别日期格式的列：取每列第一个非空单元格，判断其数字格式是否为日期
        $date_cols = [];
        $highest_col = $sheet->getHighestColumn();
        $check_rows = min($sheet->getHighestRow(), 50); // 只检查前50行，足够判断列格式
        $col_letter = 'A';
        $col_index = 0;
        while (true) {
            $is_date = false;
            for ($r = 1; $r <= $check_rows; $r++) {
                $cell = $sheet->getCell($col_letter.$r);
                $v = $cell->getValue();
                if ($v !== null && $v !== '') {
                    $fmt = (string)$cell->getStyle()->getNumberFormat()->getFormatCode();
                    // 日期格式特征：含 y/m/d 等日期符号，且非纯文本格式(@)与数字格式
                    if (preg_match('/[ymd]/i', $fmt)
                        && stripos($fmt, '@') === false
                        && stripos($fmt, '[') === false
                        && stripos($fmt, '0') === false) {
                        $is_date = true;
                    }
                    break;
                }
            }
            if ($is_date) {
                $date_cols[$col_index] = true;
            }
            if ($col_letter == $highest_col) {
                break;
            }
            $col_letter++;
            $col_index++;
        }

        $data = $sheet->toArray(null, true, true, false);

        // 值级日期列检测（补充）：针对 General 格式的数字列（如Excel里直接是 46261 序列号）。
        // 取每列前10个非空数据单元格，若全部为纯数字、无前导零、且在日期序列号/时间戳范围，则视为日期列。
        $col_cnt = count($data[0]);
        for ($ci = 0; $ci < $col_cnt; $ci++) {
            if (isset($date_cols[$ci])) {
                continue; // 格式检测已命中
            }
            $is_date = true;
            $checked = 0;
            foreach ($data as $ri => $row) {
                if ($ri == 0) {
                    continue; // 跳过表头
                }
                $v = isset($row[$ci]) ? $row[$ci] : '';
                if ($v === null || $v === '') {
                    continue;
                }
                $checked++;
                $s = trim((string)$v);
                if (!is_numeric($s) || preg_match('/^0\d/', $s)) {
                    $is_date = false;
                    break;
                }
                $f = (float)$s;
                if (!(($f > 20000 && $f < 60000) || ($f > 1000000000 && $f < 2000000000))) {
                    $is_date = false;
                    break;
                }
                if ($checked >= 10) {
                    break;
                }
            }
            if ($is_date && $checked > 0) {
                $date_cols[$ci] = true;
            }
        }

        // 清理空行
        $rows = [];
        foreach ($data as $r) {
            $row = [];
            foreach ($r as $v) {
                // 公式计算结果、富文本等对象统一转字符串
                if (is_object($v)) {
                    if (method_exists($v, '__toString')) {
                        $row[] = trim((string)$v);
                    } else {
                        $row[] = '';
                    }
                } elseif (is_bool($v)) {
                    $row[] = $v ? '1' : '0';
                } elseif ($v === null) {
                    $row[] = '';
                } else {
                    $row[] = trim((string)$v);
                }
            }
            // 整行都为空则跳过
            $empty = true;
            foreach ($row as $v) {
                if ($v !== '') {
                    $empty = false;
                    break;
                }
            }
            if (!$empty) {
                $rows[] = $row;
            }
        }

        if (!$rows) {
            return dr_return_data(0, dr_lang('Excel中没有有效数据'));
        }

        // 日期格式列：把Excel日期序列号/时间戳统一转为 Y-m-d 字符串
        // （否则Text等字段会把日期读成46261这类序列号数字）
        if ($date_cols) {
            foreach ($rows as $ri => $row) {
                foreach ($date_cols as $ci => $_) {
                    if (!isset($row[$ci]) || $row[$ci] === '' || !is_numeric($row[$ci])) {
                        continue;
                    }
                    $v = (float)$row[$ci];
                    if ($v > 1000000000) {
                        // 10位时间戳
                        $rows[$ri][$ci] = date('Y-m-d', (int)$v);
                    } elseif ($v > 10000 && $v < 60000) {
                        // Excel日期序列号
                        try {
                            $rows[$ri][$ci] = \PhpOffice\PhpSpreadsheet\Shared\Date::excelToDateTimeObject($v)->format('Y-m-d');
                        } catch (\Throwable $e) {
                            // 转换失败则保留原值
                        }
                    }
                }
            }
        }

        // 统一列数
        $max = 0;
        foreach ($rows as $r) {
            $c = count($r);
            if ($max < $c) {
                $max = $c;
            }
        }
        foreach ($rows as $i => $r) {
            $rows[$i] = array_pad($r, $max, '');
        }

        return dr_return_data(1, 'ok', [
            'rows' => $rows,
            'total' => count($rows),
            'cols' => $max,
        ]);
    }

    /**
     * 检测CSV文件编码
     */
    protected function _detect_encoding($file) {

        $fp = fopen($file, 'r');
        if (!$fp) {
            return '';
        }
        $content = fread($fp, 8192);
        fclose($fp);
        if (!$content) {
            return '';
        }

        // BOM判断
        if (strpos($content, "\xEF\xBB\xBF") === 0) {
            return 'UTF-8';
        }
        if (strpos($content, "\xFF\xFE") === 0 || strpos($content, "\xFE\xFF") === 0) {
            return 'UTF-16';
        }

        // UTF-8有效
        if (function_exists('mb_check_encoding') && mb_check_encoding($content, 'UTF-8')) {
            return 'UTF-8';
        }

        // 常见中文编码
        foreach (['GB18030', 'BIG-5'] as $enc) {
            if (function_exists('mb_check_encoding') && mb_check_encoding($content, $enc)) {
                return $enc;
            }
        }

        return 'GB18030';
    }

    /**
     * 检测CSV分隔符
     */
    protected function _detect_delimiter($file) {

        $fp = fopen($file, 'r');
        if (!$fp) {
            return ',';
        }
        $line = fgets($fp);
        fclose($fp);
        if (!$line) {
            return ',';
        }

        $delims = [',', "\t", ';', '|'];
        $best = ',';
        $bestc = 0;
        foreach ($delims as $d) {
            $c = substr_count($line, $d);
            if ($c > $bestc) {
                $bestc = $c;
                $best = $d;
            }
        }

        return $best;
    }

    /**
     * 单元格值转换
     * @param mixed $value 单元格原始值
     * @param string $fieldtype 模块字段类型
     * @param string $fieldname 字段名（用于日期启发式识别）
     * @return mixed
     */
    public function convert_value($value, $fieldtype = '', $fieldname = '') {

        if ($value === null || $value === '') {
            return '';
        }

        // 日期字段：支持时间戳 / Excel序列号 / 常见日期字符串
        if ($fieldtype == 'Date') {
            if (is_numeric($value)) {
                $v = (float)$value;
                // 10位时间戳
                if ($v > 1000000000) {
                    return (int)$v;
                }
                // Excel序列号
                if ($v > 10000 && $v < 60000) {
                    try {
                        return \PhpOffice\PhpSpreadsheet\Shared\Date::excelToDateTimeObject($v)->getTimestamp();
                    } catch (\Throwable $e) {
                        return 0;
                    }
                }
            }
            $str = trim((string)$value);
            // 兼容 2026-09-18 / 2026/09/18 / 2026.09.18 等格式
            $str = str_replace(['/', '.'], '-', $str);
            $ts = strtotime($str);
            if ($ts) {
                return $ts;
            }
            return 0;
        }

        // 整型字段
        if (in_array($fieldtype, ['Number', 'Touchspin', 'Integer'])) {
            return (int)$value;
        }

        // 普通文本字段：字段名含日期特征（riqi/date/日期/time/时间）时，
        // 若值为纯数字（Excel日期序列号或时间戳），自动转为日期字符串。
        // 例：qianziriqi=46261  → 2026-08-27
        $str_value = trim((string)$value);
        if (preg_match('/riqi|date|日期|time|时间/i', $fieldname)
            && is_numeric($str_value)
            && !preg_match('/^0\d/', $str_value) // 排除 055001 这类带前导零的编号
        ) {
            $v = (float)$str_value;
            if ($v > 1000000000 && $v < 2000000000) {
                // 10位时间戳
                return date('Y-m-d', (int)$v);
            }
            if ($v > 20000 && $v < 60000) {
                // Excel日期序列号
                try {
                    return \PhpOffice\PhpSpreadsheet\Shared\Date::excelToDateTimeObject($v)->format('Y-m-d');
                } catch (\Throwable $e) {
                    // 转换失败保留原值
                }
            }
        }

        return trim((string)$value);
    }

    /**
     * 执行导入（分批）
     * @param array $param 参数
     * @return array dr_return_data格式
     */
    public function run_import($param) {

        $file = $param['file'];
        $mid = $param['mid'];
        $catid = max(0, (int)$param['catid']);
        $skip = (int)$param['skip']; // 是否跳过首行表头
        $start = max(0, (int)$param['start']);
        $limit = min(200, max(1, (int)$param['limit']));
        $mode = $param['mode'] == 'skip' ? 'skip' : 'add'; // skip=按标题去重
        $map = $param['map']; // 列索引 => 字段名
        $uid = max(0, (int)$param['uid']);
        $author = (string)$param['author'];

        // 大批量导入：解除执行时间限制、放宽内存
        @set_time_limit(0);
        @ini_set('memory_limit', '512M');

        // ========== 数据读取：优先使用预览阶段生成的序列化缓存 ==========
        // 避免每批都重新解析Excel（2万行xlsx单次解析需数秒~数十秒，是导入慢与超时的主因）
        $dat = $file.'.dat';
        if (is_file($dat)) {
            $cached = @file_get_contents($dat);
            $rows = $cached ? @unserialize($cached) : [];
            if (!is_array($rows)) {
                return dr_return_data(0, dr_lang('数据缓存损坏，请重新上传文件'));
            }
        } else {
            $rt = $this->read_excel($file);
            if (!$rt['code']) {
                return $rt;
            }
            $rows = $rt['data']['rows'];
            if ($skip && $rows) {
                array_shift($rows); // 移除表头行
            }
            // 顺带生成缓存，后续批次直接复用
            @file_put_contents($dat, serialize($rows));
        }
        $total = count($rows);
        if ($total <= 0) {
            return dr_return_data(0, dr_lang('没有可导入的数据'));
        }

        // 模块字段
        $module = \Phpcmf\Service::L('cache')->get('module-'.SITE_ID.'-'.$mid);
        if (!$module) {
            return dr_return_data(0, dr_lang('模块不存在'));
        }
        $fields = [];
        if (isset($module['field']) && $module['field']) {
            $fields = $module['field'];
        }
        // 栏目自定义字段（mstp/ziliao等，save_content会自动写入栏目数据表）
        if (isset($module['category_data_field']) && $module['category_data_field']) {
            foreach ($module['category_data_field'] as $f) {
                if (isset($f['fieldname'])) {
                    $fields[$f['fieldname']] = $f;
                }
            }
        }

        // 栏目（名称 => id）
        $cats = [];
        $cc = \Phpcmf\Service::L('category', 'module')->get_category($mid);
        if ($cc) {
            foreach ($cc as $cat) {
                if (isset($cat['name']) && isset($cat['id'])) {
                    $cats[$cat['name']] = (int)$cat['id'];
                }
            }
        }

        // 内容模型（控制器已通过_module_init初始化）
        $content = \Phpcmf\Service::C()->content_model;
        if (!$content) {
            return dr_return_data(0, dr_lang('内容模型初始化失败'));
        }

        // 标题所在列（用于去重预取）
        $title_col = null;
        foreach ($map as $col => $fname) {
            if ($fname == 'title') {
                $title_col = (int)$col;
                break;
            }
        }

        $end = min($total, $start + $limit);

        // 去重模式：一次性预取本批范围内已存在的标题（只需查一次库）
        $exists = [];
        if ($mode == 'skip' && $title_col !== null) {
            $titles = [];
            for ($i = $start; $i < $end; $i++) {
                $t = trim((string)(isset($rows[$i][$title_col]) ? $rows[$i][$title_col] : ''));
                if ($t !== '') {
                    $titles[] = $t;
                }
            }
            if ($titles) {
                $q = $this->db->table($content->mytable)->whereIn('title', $titles)->get();
                $rs = $q ? $q->getResultArray() : [];
                foreach ($rs as $r) {
                    $exists[$r['title']] = 1;
                }
            }
        }

        // 批次处理
        $success = 0;
        $fail = 0;
        $skipcount = 0;
        $msgs = [];
        $batch_titles = []; // 本批已写入的标题（去重模式防批内重复）

        // 批次事务：整批成功才提交，失败/异常自动回滚，保证可安全重试与续传
        $this->trans_start();
        try {

        for ($i = $start; $i < $end; $i++) {

            $row = $rows[$i];
            $save = [0 => [], 1 => []];
            $row_catid = $catid;
            $error = '';

            // 字段映射
            foreach ($map as $col => $fieldname) {
                $col = (int)$col;
                $fieldname = (string)$fieldname;
                if (!$fieldname || $fieldname == 'none') {
                    continue;
                }
                $value = isset($row[$col]) ? $row[$col] : '';

                // 栏目列：按名称匹配
                if ($fieldname == '__catid__') {
                    if ($value !== '') {
                        if (isset($cats[$value])) {
                            $row_catid = $cats[$value];
                        } else {
                            $error = dr_lang('栏目【%s】不存在', $value);
                        }
                    }
                    continue;
                }

                // 跳过未知字段
                if (!isset($fields[$fieldname])) {
                    continue;
                }
                $f = $fields[$fieldname];

                // 值转换
                $value = $this->convert_value($value, $f['fieldtype'], $fieldname);

                // 安全过滤
                if ($value !== '') {
                    if (in_array($f['fieldtype'], ['Ueditor', 'Editor'])) {
                        // 富文本：保留HTML，过滤危险脚本
                        $value = \Phpcmf\Service::L('Security')->xss_clean((string)$value);
                    } else {
                        // 普通文本：与CMS后台录入行为一致
                        $value = dr_safe_replace((string)$value);
                    }
                }

                // ismain=1 写入主表，ismain=0 写入附表
                $save[$f['ismain'] ? 1 : 0][$fieldname] = $value;
            }

            if ($error) {
                $fail++;
                $msgs[] = [
                    'row' => $i + ($skip ? 2 : 1),
                    'title' => isset($save[1]['title']) ? $save[1]['title'] : '',
                    'msg' => $error,
                ];
                continue;
            }

            // 标题必填
            if (!isset($save[1]['title']) || trim($save[1]['title']) === '') {
                $fail++;
                $msgs[] = [
                    'row' => $i + ($skip ? 2 : 1),
                    'title' => '',
                    'msg' => dr_lang('标题不能为空'),
                ];
                continue;
            }

            // 按标题去重（预取结果 + 批内已写入）
            if ($mode == 'skip') {
                $t = (string)$save[1]['title'];
                if (isset($batch_titles[$t]) || isset($exists[$t])) {
                    $skipcount++;
                    continue;
                }
                $batch_titles[$t] = 1;
            }

            // 系统字段
            $save[1]['uid'] = $uid;
            // 笔名：Excel已映射则保留，否则用默认值
            if (!isset($save[1]['author']) || $save[1]['author'] === '') {
                $save[1]['author'] = $author;
            }
            $save[1]['catid'] = $row_catid;
            $save[1]['status'] = 9; // 直接发布
            $save[1]['url'] = '';
            $save[1]['link_id'] = 0;
            $save[1]['hits'] = isset($save[1]['hits']) ? (int)$save[1]['hits'] : 0;
            $save[1]['displayorder'] = isset($save[1]['displayorder']) ? (int)$save[1]['displayorder'] : 0;
            $save[1]['tableid'] = 0;
            $save[1]['inputip'] = '';
            !isset($save[1]['updatetime']) && $save[1]['updatetime'] = SYS_TIME;
            !isset($save[1]['inputtime']) && $save[1]['inputtime'] = SYS_TIME;
            $save[0]['uid'] = $uid;
            $save[0]['catid'] = $row_catid;

            // 调用官方接口保存内容
            $rt = $content->save_content(0, $save);
            if ($rt['code']) {
                $success++;
            } else {
                $fail++;
                $msgs[] = [
                    'row' => $i + ($skip ? 2 : 1),
                    'title' => $save[1]['title'],
                    'msg' => $rt['msg'],
                ];
            }
        }

        } catch (\Throwable $e) {
            // 异常：回滚整批，返回可重试
            $this->trans_rollback();
            return dr_return_data(0, dr_lang('批次执行异常，该批数据已回滚，可点击继续导入重试：%s', $e->getMessage()));
        }

        // 提交事务（状态异常时trans_comment内部已自动回滚）
        if (!$this->trans_comment()) {
            return dr_return_data(0, dr_lang('批次写入失败，该批数据已回滚，可点击继续导入重试'));
        }

        return dr_return_data(1, 'ok', [
            'total' => $total,
            'success' => $success,
            'fail' => $fail,
            'skip' => $skipcount,
            'done' => $end,
            'finish' => $end >= $total ? 1 : 0,
            'msgs' => $msgs,
        ]);
    }
}
