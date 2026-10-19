-- Excel数据导入插件 卸载SQL
-- 仅清理本插件的后台菜单，数据表与数据不做删除

DELETE FROM `{dbprefix}admin_menu` WHERE `mark`='app-excelimp';
