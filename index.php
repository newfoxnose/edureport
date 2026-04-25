<?php
/**
 * 主入口文件
 * 注意：此版本不使用数据持久化，所有数据都在内存中处理
 */

// 错误报告设置
error_reporting(E_ALL ^ E_NOTICE);

// 加载Composer自动加载
require_once __DIR__ . '/vendor/autoload.php';

// 加载辅助函数
require_once __DIR__ . '/helpers.php';
require_once __DIR__ . '/MemoryDataQuery.php'; // 内存数据查询工具类
require_once __DIR__ . '/controller.php';

// 简单的路由系统
$action = isset($_GET['action']) ? $_GET['action'] : 'home';

// 根据action执行相应函数（所有数据都在内存中处理，不需要DataStore）
switch ($action) {
    case 'home':
        home();
        break;
    case 'import_students':
        import_students();
        break;
    case 'import_teachers':
        import_teachers();
        break;
    default:
        home();
        break;
}

