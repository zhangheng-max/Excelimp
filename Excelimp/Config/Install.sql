-- Excel数据导入插件 安装SQL
-- 表前缀使用 {dbprefix} 占位，安装时自动替换为实际前缀
-- 本插件不新增数据表，直接使用迅睿CMS模块内容表
-- 菜单：挂载到 内容->内容管理 分组下（幂等，可重复执行）

INSERT INTO `{dbprefix}admin_menu` (`pid`,`name`,`site`,`uri`,`url`,`mark`,`hidden`,`icon`,`displayorder`)
SELECT id,'Excel数据导入','','excelimp/home/index','','app-excelimp',0,'fa fa-file-excel-o',0
FROM `{dbprefix}admin_menu`
WHERE `mark`='content-module'
  AND NOT EXISTS (SELECT 1 FROM `{dbprefix}admin_menu` WHERE `mark`='app-excelimp')
LIMIT 1;
