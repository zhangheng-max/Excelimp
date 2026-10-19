<?php namespace Phpcmf\Controllers\Admin;

/**
 * Excel数据导入插件 - 后台控制器
 * 适用于迅睿CMS 4.7.3
 *
 * 流程：
 *   1. index    选择模块、栏目，上传Excel文件
 *   2. preview  字段映射 + 数据预览
 *   3. doimport 分批执行导入（ajax）
 */

class Home extends \Phpcmf\Common {

    // 允许的文件大小 20MB
    protected $max_size = 20 * 1024 * 1024;

    // 允许的扩展名
    protected $allow_ext = ['xlsx', 'xls', 'csv'];

    public function __construct() {
        parent::__construct();
        if (!$this->admin) {
            $this->_admin_msg(0, dr_lang('无权限操作'));
        }
        \Phpcmf\Service::V()->assign([
            'menu' => \Phpcmf\Service::M('auth')->_admin_menu(
                [
                    'Excel数据导入' => [APP_DIR.'/'.\Phpcmf\Service::L('Router')->class.'/index', 'fa fa-file-excel-o'],
                    'help' => [],
                ]
            ),
        ]);
    }

    // ========== 第一步：选择模块与上传 ==========
    public function index() {

        // 处理上传
        if (IS_POST) {

            // CSRF等系统校验已由框架执行

            // 模块参数
            $mid = dr_safe_filename(\Phpcmf\Service::L('input')->post('mid'));
            $catid = max(0, (int)\Phpcmf\Service::L('input')->post('catid'));
            $skip = \Phpcmf\Service::L('input')->post('skip') ? 1 : 0;

            if (!$mid) {
                $this->_admin_msg(0, dr_lang('请选择要导入的模块'));
            }
            $module = \Phpcmf\Service::L('cache')->get('module-'.SITE_ID.'-'.$mid);
            if (!$module) {
                $this->_admin_msg(0, dr_lang('模块不存在'));
            }

            // 文件检查
            if (!isset($_FILES['file']) || !is_array($_FILES['file']) || !$_FILES['file']['name']) {
                $this->_admin_msg(0, dr_lang('请选择要上传的Excel文件'));
            }
            $file = $_FILES['file'];
            if ($file['error']) {
                $this->_admin_msg(0, dr_lang('文件上传失败（错误码%s）', $file['error']));
            }
            if ($file['size'] > $this->max_size) {
                $this->_admin_msg(0, dr_lang('文件大小不能超过20MB'));
            }
            $ext = strtolower(pathinfo($file['name'], PATHINFO_EXTENSION));
            if (!in_array($ext, $this->allow_ext)) {
                $this->_admin_msg(0, dr_lang('仅支持xlsx/xls/csv格式文件'));
            }
            if (!is_uploaded_file($file['tmp_name'])) {
                $this->_admin_msg(0, dr_lang('文件上传失败'));
            }

            // 保存到临时目录
            $fname = 'excelimp_'.date('Ymd').'_'.md5(SYS_TIME.mt_rand(1000, 9999)).'.'.$ext;
            $save = WRITEPATH.'temp/'.$fname;
            if (!@move_uploaded_file($file['tmp_name'], $save)) {
                $this->_admin_msg(0, dr_lang('文件保存失败，请检查cache/temp目录权限'));
            }

            // 进入预览
            dr_redirect(\Phpcmf\Service::L('Router')->url('excelimp/home/preview', [
                'file' => $fname,
                'mid' => $mid,
                'catid' => $catid,
                'skip' => $skip,
            ]));
            exit;
        }

        // 模块列表（含栏目）
        $modules = \Phpcmf\Service::L('cache')->get('module-'.SITE_ID);
        $data = [];
        if ($modules) {
            foreach ($modules as $dir => $t) {
                if (!isset($t['dirname']) || !$t['dirname']) {
                    continue;
                }
                $m = \Phpcmf\Service::L('cache')->get('module-'.SITE_ID.'-'.$t['dirname']);
                if (!$m || !isset($m['dirname'])) {
                    continue;
                }
                $cats = [];
                $cc = \Phpcmf\Service::L('category', 'module')->get_category($t['dirname']);
                if ($cc) {
                    foreach ($cc as $cat) {
                        $cats[] = [
                            'id' => (int)$cat['id'],
                            'name' => $cat['name'],
                        ];
                    }
                }
                $data[] = [
                    'dirname' => $t['dirname'],
                    'name' => $m['name'],
                    'cats' => $cats,
                ];
            }
        }

        \Phpcmf\Service::V()->assign([
            'module' => $data,
            'form' => dr_form_hidden(),
        ]);
        \Phpcmf\Service::V()->display('index.html');
    }

    // ========== 第二步：字段映射与预览 ==========
    public function preview() {

        // 大批量文件解析与缓存：放宽内存与执行时间
        @set_time_limit(0);
        @ini_set('memory_limit', '512M');

        $fname = dr_safe_filename(\Phpcmf\Service::L('input')->get('file'));
        $mid = dr_safe_filename(\Phpcmf\Service::L('input')->get('mid'));
        $catid = max(0, (int)\Phpcmf\Service::L('input')->get('catid'));
        $skip = \Phpcmf\Service::L('input')->get('skip') ? 1 : 0;

        // 参数校验
        if (!$fname || !is_file(WRITEPATH.'temp/'.$fname)) {
            $this->_admin_msg(0, dr_lang('文件不存在或已过期，请重新上传'));
        }
        $ext = strtolower(pathinfo($fname, PATHINFO_EXTENSION));
        if (!in_array($ext, $this->allow_ext)) {
            $this->_admin_msg(0, dr_lang('文件格式错误'));
        }
        if (!$mid) {
            $this->_admin_msg(0, dr_lang('模块参数错误'));
        }
        $module = \Phpcmf\Service::L('cache')->get('module-'.SITE_ID.'-'.$mid);
        if (!$module) {
            $this->_admin_msg(0, dr_lang('模块不存在'));
        }

        // 解析Excel
        $imp = \Phpcmf\Service::M('import', 'excelimp');
        $rt = $imp->read_excel(WRITEPATH.'temp/'.$fname);
        if (!$rt['code']) {
            $this->_admin_msg(0, $rt['msg']);
        }
        $rows = $rt['data']['rows'];

        // 表头与预览数据
        if ($skip && $rows) {
            $head = $rows[0];
            array_shift($rows); // 移除表头，后续导入直接复用
            $data_rows = array_slice($rows, 0, 8); // 预览前8行
        } else {
            $head = [];
            $cols = isset($rt['data']['cols']) ? $rt['data']['cols'] : 0;
            for ($i = 0; $i < $cols; $i++) {
                $head[] = dr_lang('第%s列', $i + 1);
            }
            $data_rows = array_slice($rows, 0, 8);
        }
        $total = count($rows);

        // 数据缓存：序列化去表头后的完整数据，分批导入时直接复用，避免每批重复解析Excel
        $dat = WRITEPATH.'temp/'.$fname.'.dat';
        @file_put_contents($dat, serialize($rows));
        if (!is_file($dat)) {
            $this->_admin_msg(0, dr_lang('数据缓存写入失败，请检查cache/temp目录权限'));
        }

        // 字段列表
        $fields = [];
        if (isset($module['field']) && $module['field']) {
            foreach ($module['field'] as $f) {
                if (isset($f['fieldname'])) {
                    $fields[] = [
                        'fieldname' => $f['fieldname'],
                        'name' => $f['name'],
                        'ismain' => (int)$f['ismain'],
                        'fieldtype' => $f['fieldtype'],
                        'group' => dr_lang('模块字段'),
                    ];
                }
            }
        }
        // 栏目自定义字段
        if (isset($module['category_data_field']) && $module['category_data_field']) {
            foreach ($module['category_data_field'] as $f) {
                if (isset($f['fieldname'])) {
                    $fields[] = [
                        'fieldname' => $f['fieldname'],
                        'name' => $f['name'],
                        'ismain' => (int)$f['ismain'],
                        'fieldtype' => $f['fieldtype'],
                        'group' => dr_lang('栏目字段'),
                    ];
                }
            }
        }

        // 栏目
        $cats = [];
        $cc = \Phpcmf\Service::L('category', 'module')->get_category($mid);
        if ($cc) {
            foreach ($cc as $cat) {
                $cats[] = [
                    'id' => (int)$cat['id'],
                    'name' => $cat['name'],
                ];
            }
        }

        \Phpcmf\Service::V()->assign([
            'file' => $fname,
            'mid' => $mid,
            'catid' => $catid,
            'skip' => $skip,
            'head' => $head,
            'rows' => $data_rows,
            'total' => $total,
            'fields' => $fields,
            'cats' => $cats,
            'module_name' => $module['name'],
            'author_name' => $this->_get_author_name(),
            'form' => dr_form_hidden(),
        ]);
        \Phpcmf\Service::V()->display('preview.html');
    }

    // ========== 第三步：执行导入（ajax分批） ==========
    public function doimport() {

        if (!IS_AJAX_POST) {
            $this->_json(0, dr_lang('非法请求'));
        }

        // 大批量导入：解除执行时间限制、放宽内存（2万条分批导入）
        @set_time_limit(0);
        @ini_set('memory_limit', '512M');

        $fname = dr_safe_filename(\Phpcmf\Service::L('input')->post('file'));
        $mid = dr_safe_filename(\Phpcmf\Service::L('input')->post('mid'));
        $catid = max(0, (int)\Phpcmf\Service::L('input')->post('catid'));
        $skip = \Phpcmf\Service::L('input')->post('skip') ? 1 : 0;
        $start = max(0, (int)\Phpcmf\Service::L('input')->post('start'));
        $limit = min(200, max(1, (int)\Phpcmf\Service::L('input')->post('limit')));
        $mode = dr_safe_replace((string)\Phpcmf\Service::L('input')->post('mode'));
        $map = \Phpcmf\Service::L('input')->post('map');
        if (!is_array($map)) {
            $map = [];
        }
        foreach ($map as $i => $f) {
            $map[$i] = dr_safe_replace((string)$f);
        }

        // 参数校验
        if (!$fname || !is_file(WRITEPATH.'temp/'.$fname)) {
            $this->_json(0, dr_lang('文件不存在或已过期，请重新上传'));
        }
        if (!$mid) {
            $this->_json(0, dr_lang('模块参数错误'));
        }

        // 初始化模块（设置内容模型与模块缓存）
        $this->_module_init($mid);
        if (!$this->module) {
            $this->_json(0, dr_lang('模块初始化失败'));
        }

        $imp = \Phpcmf\Service::M('import', 'excelimp');
        // 作者：表单填写则使用表单值，否则默认当前管理员（Excel映射的笔名列优先）
        $post_author = trim((string)\Phpcmf\Service::L('input')->post('author'));
        $rt = $imp->run_import([
            'file' => WRITEPATH.'temp/'.$fname,
            'mid' => $mid,
            'catid' => $catid,
            'skip' => $skip,
            'start' => $start,
            'limit' => $limit,
            'mode' => $mode,
            'map' => $map,
            'uid' => (int)$this->admin['uid'],
            'author' => $post_author ? $post_author : $this->_get_author_name(),
        ]);
        if (!$rt['code']) {
            $this->_json(0, $rt['msg']);
        }
        $this->_json(1, dr_lang('ok'), $rt['data']);
    }

    // 当前管理员作者名（优先会员昵称）
    protected function _get_author_name() {

        $name = '';
        $uid = (int)$this->admin['uid'];
        if ($uid) {
            $member = dr_member_info($uid);
            if ($member) {
                $name = $member['name'] ? $member['name'] : $member['username'];
            }
        }
        if (!$name) {
            $name = $this->admin['username'];
        }

        return $name;
    }

    // 下载导入模板示例（按所选模块字段生成）
    public function template() {

        $mid = dr_safe_filename(\Phpcmf\Service::L('input')->get('mid'));
        if (!$mid) {
            $this->_json(0, dr_lang('模块参数错误'));
        }
        $module = \Phpcmf\Service::L('cache')->get('module-'.SITE_ID.'-'.$mid);
        if (!$module) {
            $this->_json(0, dr_lang('模块不存在'));
        }

        $vendor = FCPATH.'App/Vendor/autoload.php';
        if (!is_file($vendor)) {
            $this->_json(0, dr_lang('系统缺少PhpSpreadsheet组件'));
        }
        require_once $vendor;

        $spreadsheet = new \PhpOffice\PhpSpreadsheet\Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();

        // 表头
        $row = 1;
        $col = 1;
        $head = [];
        foreach ($module['field'] as $f) {
            if (!isset($f['fieldname']) || !in_array($f['fieldname'], ['title', 'keywords', 'description', 'author'])) {
                continue;
            }
            $sheet->setCellValueByColumnAndRow($col, $row, $f['name']);
            $col++;
        }
        $sheet->setCellValueByColumnAndRow($col, $row, dr_lang('栏目（按名称）'));
        $col++;
        $sheet->setCellValueByColumnAndRow($col, $row, dr_lang('内容'));

        // 示例行
        $row = 2;
        $sheet->setCellValueByColumnAndRow(1, $row, dr_lang('示例标题'));

        $ext = 'xlsx';
        $filename = $mid.'_导入模板.'.$ext;
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Disposition: attachment;filename="'.$filename.'"');
        header('Cache-Control: max-age=0');

        $writer = new \PhpOffice\PhpSpreadsheet\Writer\Xlsx($spreadsheet);
        $writer->save('php://output');
        exit;
    }
}
