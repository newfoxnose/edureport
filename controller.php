<?php
/**
 * 控制器文件 - 包含所有业务逻辑
 */

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;

/**
 * 首页
 */
function home()
{
    include __DIR__ . '/homeview.php';
}

/**
 * 导入学生
 */
function import_students()
{
    // 检查文件上传
    if ($_FILES["file"]["error"] > 0) {
        exit("文件上传错误：" . $_FILES["file"]["error"]);
    } else {
        // 检查文件类型
        $fileType = $_FILES["file"]["type"];
        $allowedTypes = array("application/vnd.ms-excel", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        
        if (in_array($fileType, $allowedTypes) || 
            pathinfo($_FILES["file"]["name"], PATHINFO_EXTENSION) == 'xls' || 
            pathinfo($_FILES["file"]["name"], PATHINFO_EXTENSION) == 'xlsx') {
            
            // 根据文件类型选择读取器
            $reader = IOFactory::createReaderForFile($_FILES["file"]["tmp_name"]);
            
            // 载入excel文件
            $spreadsheet = $reader->load($_FILES["file"]["tmp_name"]);
            $sheet = $spreadsheet->getActiveSheet();
            
            $highestRow = $sheet->getHighestRow();
            
            // 在内存中存储学生数据
            $studentsData = array();
            $successCount = 0;
            $errorCount = 0;
            $errors = array();
            
            // 从第2行开始读取数据（第1行是标题）
            for ($row = 2; $row <= $highestRow; $row++) {
                $identity_number = trim_identity_number($sheet->getCell('C' . $row)->getValue()); // 身份证号码
                // 跳过空行
                if (empty($identity_number)) {
                    continue;
                }
                
                $minority = trim_str($sheet->getCell('D' . $row)->getValue()); // 少数民族
                $phase = trim_str($sheet->getCell('E' . $row)->getValue()); // 学段
                $grade = trim_str($sheet->getCell('F' . $row)->getValue()); // 年级
                $boarder = trim_str($sheet->getCell('G' . $row)->getValue()); // 寄宿生（是/否，与少数民族列用法一致）
                $leftbehind = trim_str($sheet->getCell('H' . $row)->getValue()); // 农村留守儿童（是/否）
                
                // 从身份证号码计算年龄和性别
                $age = get_age_from_id_september($identity_number);
                $sex = get_sex_from_id_chinese($identity_number);
                
                // 准备数据（存储在内存中）
                $studentData = array(
                    'id' => $successCount + 1,
                    'identity_number' => $identity_number,
                    'minority' => $minority,
                    'phase' => $phase,
                    'grade' => $grade,
                    'boarder' => $boarder,
                    'leftbehind' => $leftbehind,
                    'age' => $age,
                    'sex' => $sex
                );
                
                // 添加到内存数组
                $studentsData[] = $studentData;
                $successCount++;
            }
            
            // 有有效行时必须先下发 Excel：若先 echo HTML，响应体已开始，attachment 头无效（PHP Desktop/内嵌浏览器将无法下载）
            if ($successCount > 0) {
                generate_students_statistics($studentsData);
            }
            
            // 无数据时展示结果页（与下载互斥，避免混排 HTML 与二进制）
            echo "<h2>导入完成</h2>";
            echo "<p>成功导入：" . $successCount . " 条记录</p>";
            echo "<p>失败：" . $errorCount . " 条记录</p>";
            // 无有效行时不会调用 generate_excel，避免用户误以为已生成文件
            if ($successCount === 0) {
                echo "<p><strong>未生成 Excel：</strong>未读取到有效数据。请确认从第 2 行起 <strong>C 列</strong>为身份证号码且非空。</p>";
            }
            
            if (count($errors) > 0) {
                echo "<h3>错误信息：</h3>";
                foreach ($errors as $error) {
                    echo "<p>" . $error . "</p>";
                }
            }
            
            echo "<p><a href='?action=home'>返回首页</a></p>";
        } else {
            exit("文件名格式错误，请上传Excel文件（.xls或.xlsx格式）");
        }
    }
}

/**
 * 生成学生数据统计报表（使用内存数据）
 */
function generate_students_statistics($studentsData)
{
    // 将多个报表数据组成数组，每个元素会生成一个工作表
    generate_excel(array(
        report_3112($studentsData, '小学'),
        report_3113($studentsData, '小学'),
        report_3112($studentsData, '初中'),
        report_3113($studentsData, '初中'),
        report_3112($studentsData, '普通高中'),
        report_3113($studentsData, '普通高中')
    ), '年报导出(学生)');
}

/**
 * 教基3112、3115、3118（使用内存数据）
 */
function report_3112($studentsData, $phase)
{
    // 先过滤出当前学段的数据
    $phaseData = MemoryDataQuery::filter($studentsData, array("phase" => $phase));
    
    switch ($phase) {
        case '小学':
            $grade_arr = array("一年级", "二年级", "三年级", "四年级", "五年级", "六年级");
            $temp_x = 4;
            break;
        case '初中':
            $grade_arr = array("一年级", "二年级", "三年级", "四年级");
            $temp_x = 0;
            break;
        case '普通高中':
            $grade_arr = array("一年级", "二年级", "三年级");
            $temp_x = 0;
            break;
    }
    
    $data['line'] = array();
    
    //第1行，总计
    $data['line'][0][0][0] = array("count" => "总计");
    $data['line'][0][1][0] = array("count" => "01");
    $data['line'][0][2][0] = array("count" => "填去年数据");
    
    //招生数
    $count = MemoryDataQuery::count($phaseData, array("grade" => '一年级'));
    $data['line'][0][3] = array(array("count" => $count));
    
    if ($temp_x != 0) {
        for ($i = 4; $i < 7; $i++) {
            $data['line'][0][$i][0] = array("count" => 0);
        }
        //接受过三年学前教育
        $data['line'][0][7] = $data['line'][0][3];
    }
    
    //在校生总数
    $count = MemoryDataQuery::count($phaseData);
    $data['line'][0][$temp_x + 4] = array(array("count" => $count));
    
    //在校女生总数
    $count = MemoryDataQuery::count($phaseData, array("sex" => "女"));
    $data['line'][0][$temp_x + 5] = array(array("count" => $count));
    
    //各年级在校生总数
    for ($i = 0; $i < count($grade_arr); $i++) {
        $count = MemoryDataQuery::count($phaseData, array("grade" => $grade_arr[$i]));
        $data['line'][0][$temp_x + 6 + $i] = array(array("count" => $count));
    }
    
    //预计毕业生数
    $data['line'][0][$temp_x + 6 + count($grade_arr)] = $data['line'][0][$temp_x + 6 + count($grade_arr) - 1];
    if ($phase == '初中') {
        $data['line'][0][$temp_x + 6 + count($grade_arr)] = $data['line'][0][$temp_x + 6 + count($grade_arr) - 2];
    }
    
    //第2行，女
    $data['line'][1][0][0] = array("count" => "#女");
    $data['line'][1][1][0] = array("count" => "02");
    $data['line'][1][2][0] = array("count" => "填去年数据");
    
    //招生数
    $count = MemoryDataQuery::count($phaseData, array("sex" => "女", "grade" => '一年级'));
    $data['line'][1][3] = array(array("count" => $count));
    
    if ($temp_x != 0) {
        for ($i = 4; $i < 7; $i++) {
            $data['line'][1][$i][0] = array("count" => 0);
        }
        $data['line'][1][7] = $data['line'][1][3];
    }
    
    //在校女生
    $count = MemoryDataQuery::count($phaseData, array("sex" => "女"));
    $data['line'][1][$temp_x + 4] = array(array("count" => $count));
    $data['line'][1][$temp_x + 5][0] = array("count" => "-");
    
    for ($i = 0; $i < count($grade_arr); $i++) {
        $count = MemoryDataQuery::count($phaseData, array("sex" => "女", "grade" => $grade_arr[$i]));
        $data['line'][1][$temp_x + 6 + $i] = array(array("count" => $count));
    }
    
    $data['line'][1][$temp_x + 6 + count($grade_arr)] = $data['line'][1][$temp_x + 6 + count($grade_arr) - 1];
    if ($phase == '初中') {
        $data['line'][1][$temp_x + 6 + count($grade_arr)] = $data['line'][1][$temp_x + 6 + count($grade_arr) - 2];
    }
    
    //第3行，少数民族
    $j = 2;
    $data['line'][$j][0][0] = array("count" => "#少数民族");
    $data['line'][$j][1][0] = array("count" => "03");
    $data['line'][$j][2][0] = array("count" => "填去年数据");
    
    $count = MemoryDataQuery::count($phaseData, array("minority" => '是', "grade" => '一年级'));
    $data['line'][$j][3] = array(array("count" => $count));
    
    if ($temp_x != 0) {
        for ($i = 4; $i < 7; $i++) {
            $data['line'][$j][$i][0] = array("count" => 0);
        }
        $data['line'][$j][7] = $data['line'][$j][3];
    }
    
    $count = MemoryDataQuery::count($phaseData, array("minority" => '是'));
    $data['line'][$j][$temp_x + 4] = array(array("count" => $count));
    
    $count = MemoryDataQuery::count($phaseData, array("sex" => "女", "minority" => '是'));
    $data['line'][$j][$temp_x + 5] = array(array("count" => $count));
    
    for ($i = 0; $i < count($grade_arr); $i++) {
        $count = MemoryDataQuery::count($phaseData, array("minority" => '是', "grade" => $grade_arr[$i]));
        $data['line'][$j][$temp_x + 6 + $i] = array(array("count" => $count));
    }
    
    $data['line'][$j][$temp_x + 6 + count($grade_arr)] = $data['line'][$j][$temp_x + 6 + count($grade_arr) - 1];
    if ($phase == '初中') {
        $data['line'][$j][$temp_x + 6 + count($grade_arr)] = $data['line'][$j][$temp_x + 6 + count($grade_arr) - 2];
    }
    
    //第4行，复式班
    $data['line'][3][0][0] = array("count" => "#复式班");
    $data['line'][3][1][0] = array("count" => "4");
    $data['line'][3][2][0] = array("count" => "填去年数据");
    $data['line'][3][3][0] = array("count" => "");
    if ($temp_x != 0) {
        for ($i = 4; $i < 8; $i++) {
            $data['line'][3][$i][0] = array("count" => "-");
        }
    }
    // 占位行也要补齐“预计毕业生数”列，避免固定索引导出时越界
    for ($i = 0; $i < count($grade_arr) + 3; $i++) {
        $data['line'][3][$temp_x + 4 + $i][0] = array("count" => "");
    }

    //第5行，寄宿生（源数据 G 列 boarder=是，统计口径同少数民族行）
    $j = 4;
    $data['line'][$j][0][0] = array("count" => "#寄宿生");
    $data['line'][$j][1][0] = array("count" => "5");
    $data['line'][$j][2][0] = array("count" => "填去年数据");
    
    $count = MemoryDataQuery::count($phaseData, array("boarder" => '是', "grade" => '一年级'));
    $data['line'][$j][3] = array(array("count" => $count));
    
    if ($temp_x != 0) {
        for ($i = 4; $i < 7; $i++) {
            $data['line'][$j][$i][0] = array("count" => 0);
        }
        $data['line'][$j][7] = $data['line'][$j][3];
    }
    
    $count = MemoryDataQuery::count($phaseData, array("boarder" => '是'));
    $data['line'][$j][$temp_x + 4] = array(array("count" => $count));
    
    $count = MemoryDataQuery::count($phaseData, array("sex" => "女", "boarder" => '是'));
    $data['line'][$j][$temp_x + 5] = array(array("count" => $count));
    
    for ($i = 0; $i < count($grade_arr); $i++) {
        $count = MemoryDataQuery::count($phaseData, array("boarder" => '是', "grade" => $grade_arr[$i]));
        $data['line'][$j][$temp_x + 6 + $i] = array(array("count" => $count));
    }
    
    $data['line'][$j][$temp_x + 6 + count($grade_arr)] = $data['line'][$j][$temp_x + 6 + count($grade_arr) - 1];
    if ($phase == '初中') {
        $data['line'][$j][$temp_x + 6 + count($grade_arr)] = $data['line'][$j][$temp_x + 6 + count($grade_arr) - 2];
    }
    
    //第6行至第11行：随迁子女等（占位行，指标代码 6～11；第12行农村留守儿童单独统计）
    $temp_name_arr = array("#随迁子女", "外省迁入", "本省外县迁入", "#进城务工人员随迁子女", "外省迁入", "本省外县迁入");
    for ($j = 5; $j < 11; $j++) {
        $data['line'][$j][0][0] = array("count" => $temp_name_arr[$j - 5]);
        $data['line'][$j][1][0] = array("count" => (string)($j + 1));
        $data['line'][$j][2][0] = array("count" => "填去年数据");
        $data['line'][$j][3][0] = array("count" => "");
        if ($temp_x != 0) {
            for ($i = 4; $i < 8; $i++) {
                $data['line'][$j][$i][0] = array("count" => "-");
            }
        }
        // 占位行也要补齐“预计毕业生数”列，避免固定索引导出时越界
        for ($i = 0; $i < count($grade_arr) + 3; $i++) {
            $data['line'][$j][$temp_x + 4 + $i][0] = array("count" => "");
        }
    }

    //第12行，农村留守儿童（源数据 H 列 leftbehind=是，统计口径同寄宿生）
    $j = 11;
    $data['line'][$j][0][0] = array("count" => "#农村留守儿童");
    $data['line'][$j][1][0] = array("count" => "12");
    $data['line'][$j][2][0] = array("count" => "填去年数据");
    
    $count = MemoryDataQuery::count($phaseData, array("leftbehind" => '是', "grade" => '一年级'));
    $data['line'][$j][3] = array(array("count" => $count));
    
    if ($temp_x != 0) {
        for ($i = 4; $i < 7; $i++) {
            $data['line'][$j][$i][0] = array("count" => 0);
        }
        $data['line'][$j][7] = $data['line'][$j][3];
    }
    
    $count = MemoryDataQuery::count($phaseData, array("leftbehind" => '是'));
    $data['line'][$j][$temp_x + 4] = array(array("count" => $count));
    
    $count = MemoryDataQuery::count($phaseData, array("sex" => "女", "leftbehind" => '是'));
    $data['line'][$j][$temp_x + 5] = array(array("count" => $count));
    
    for ($i = 0; $i < count($grade_arr); $i++) {
        $count = MemoryDataQuery::count($phaseData, array("leftbehind" => '是', "grade" => $grade_arr[$i]));
        $data['line'][$j][$temp_x + 6 + $i] = array(array("count" => $count));
    }
    
    $data['line'][$j][$temp_x + 6 + count($grade_arr)] = $data['line'][$j][$temp_x + 6 + count($grade_arr) - 1];
    if ($phase == '初中') {
        $data['line'][$j][$temp_x + 6 + count($grade_arr)] = $data['line'][$j][$temp_x + 6 + count($grade_arr) - 2];
    }
    
    switch ($phase) {
        case '小学':
            $data['title'] = "小学分类型学生数";
            $data['sheet_name'] = "教基3112";
            $data['columns1'] = array('指标名称', '代码', '毕业生数', '招生数', '未接受过', '一年', '两年', '三年', '在校生数', '#女', '一年级', '二年级', '三年级', '四年级', '五年级', '六年级', '预计毕业生数');
            $data['columns2'] = array('甲', '乙', '1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11', '12', '13', '14', '15');
            $data['rows'] = array();
            foreach ($data['line'] as $item) {
                $data['rows'][] = array($item[0][0]['count'], $item[1][0]['count'], $item[2][0]['count'], $item[3][0]['count'], $item[4][0]['count'], $item[5][0]['count'], $item[6][0]['count'], $item[7][0]['count'], $item[8][0]['count'], $item[9][0]['count'], $item[10][0]['count'], $item[11][0]['count'], $item[12][0]['count'], $item[13][0]['count'], $item[14][0]['count'], $item[15][0]['count'], $item[16][0]['count']);
            }
            break;
        case '初中':
            $data['title'] = "初中分类型学生数";
            $data['sheet_name'] = "教基3115";
            $data['columns1'] = array('指标名称', '代码', '毕业生数', '招生数', '在校生数', '#女', '一年级', '二年级', '三年级', '四年级', '预计毕业生数');
            $data['columns2'] = array('甲', '乙', '1', '2', '3', '4', '5', '6', '7', '8', '9');
            $data['rows'] = array();
            foreach ($data['line'] as $item) {
                $data['rows'][] = array($item[0][0]['count'], $item[1][0]['count'], $item[2][0]['count'], $item[3][0]['count'], $item[4][0]['count'], $item[5][0]['count'], $item[6][0]['count'], $item[7][0]['count'], $item[8][0]['count'], $item[9][0]['count'], $item[10][0]['count']);
            }
            break;
        case '普通高中':
            $data['title'] = "普通高中分类型学生数";
            $data['sheet_name'] = "教基3118";
            $data['columns1'] = array('指标名称', '代码', '毕业生数', '招生数', '在校生数', '#女', '一年级', '二年级', '三年级', '预计毕业生数');
            $data['columns2'] = array('甲', '乙', '1', '2', '3', '4', '5', '6', '7', '8');
            $data['rows'] = array();
            foreach ($data['line'] as $item) {
                $data['rows'][] = array($item[0][0]['count'], $item[1][0]['count'], $item[2][0]['count'], $item[3][0]['count'], $item[4][0]['count'], $item[5][0]['count'], $item[6][0]['count'], $item[7][0]['count'], $item[8][0]['count'], $item[9][0]['count']);
            }
            break;
    }
    
    return $data;
}

/**
 * 教基3113、3116、3119（使用内存数据）
 */
function report_3113($studentsData, $phase)
{
    // 先过滤出当前学段的数据
    $phaseData = MemoryDataQuery::filter($studentsData, array("phase" => $phase));
    
    $age_sql_arr = array();
    switch ($phase) {
        case '小学':
            $grade_arr = array("一年级", "二年级", "三年级", "四年级", "五年级", "六年级");
            $temp_x = 0;
            $age_sql_arr[] = array("age<=" => 5);
            for ($i = 6; $i < 15; $i++) {
                $age_sql_arr[] = array("age" => $i);
            }
            $age_sql_arr[] = array("age>=" => 15);
            break;
        case '初中':
            $grade_arr = array("一年级", "二年级", "三年级", "四年级");
            $temp_x = 0;
            $age_sql_arr[] = array("age<=" => 10);
            for ($i = 11; $i < 18; $i++) {
                $age_sql_arr[] = array("age" => $i);
            }
            $age_sql_arr[] = array("age>=" => 18);
            break;
        case '普通高中':
            $grade_arr = array("一年级", "二年级", "三年级", "四年级及以上");
            $temp_x = 1;
            $age_sql_arr[] = array("age<=" => 14);
            for ($i = 15; $i < 23; $i++) {
                $age_sql_arr[] = array("age" => $i);
            }
            $age_sql_arr[] = array("age>=" => 23);
            break;
    }
    
    $data['line'] = array();
    
    //第1行，总计
    $data['line'][0][0][0] = array("count" => "总计");
    $data['line'][0][1][0] = array("count" => "01");
    
    //招生数
    $count = MemoryDataQuery::count($phaseData, array("grade" => '一年级'));
    $data['line'][0][2] = array(array("count" => $count));
    
    //在校生总数
    $count = MemoryDataQuery::count($phaseData);
    $data['line'][0][3] = array(array("count" => $count));
    
    //在校女生总数
    $count = MemoryDataQuery::count($phaseData, array("sex" => "女"));
    $data['line'][0][4] = array(array("count" => $count));
    
    if ($phase == '普通高中') {
        $data['line'][0][5] = $data['line'][0][3];   //普通高中全日制学生等于在校生总数
    }
    
    //各年级在校生总数
    for ($i = 0; $i < count($grade_arr); $i++) {
        $count = MemoryDataQuery::count($phaseData, array("grade" => $grade_arr[$i]));
        $data['line'][0][$temp_x + 5 + $i] = array(array("count" => $count));
    }
    
    for ($i = 0; $i < count($age_sql_arr); $i++) {
        if ($i == 0) {
            $data['line'][$i + 1][0][0] = array("count" => $age_sql_arr[0]['age<='] . '岁及以下');
        } else if ($i == count($age_sql_arr) - 1) {
            $data['line'][$i + 1][0][0] = array("count" => $age_sql_arr[count($age_sql_arr) - 1]['age>='] . '岁及以上');
        } else {
            $data['line'][$i + 1][0][0] = array("count" => $age_sql_arr[$i]['age'] . '岁');
        }
        $data['line'][$i + 1][1][0] = array("count" => $i + 1);
        
        //招生数
        $where_arr = $age_sql_arr[$i] + array("grade" => '一年级');
        $count = MemoryDataQuery::count($phaseData, $where_arr);
        $data['line'][$i + 1][2] = array(array("count" => $count));
        
        //在校生总数
        $count = MemoryDataQuery::count($phaseData, $age_sql_arr[$i]);
        $data['line'][$i + 1][3] = array(array("count" => $count));
        
        //女
        $where_arr = $age_sql_arr[$i] + array("sex" => '女');
        $count = MemoryDataQuery::count($phaseData, $where_arr);
        $data['line'][$i + 1][4] = array(array("count" => $count));
        
        if ($phase == '普通高中') {
            $data['line'][$i + 1][5] = $data['line'][$i + 1][3];   //普通高中全日制学生等于在校生总数
        }
        
        for ($j = 0; $j < count($grade_arr); $j++) {
            $where_arr = $age_sql_arr[$i] + array("grade" => $grade_arr[$j]);
            $count = MemoryDataQuery::count($phaseData, $where_arr);
            $data['line'][$i + 1][$temp_x + 5 + $j] = array(array("count" => $count));
        }
    }
    
    switch ($phase) {
        case '小学':
            $data['title'] = "小学分年龄学生数";
            $data['sheet_name'] = "教基3113";
            $data['columns1'] = array('指标名称', '代码', '招生数', '在校生数', '#女', '一年级', '二年级', '三年级', '四年级', '五年级', '六年级');
            $data['columns2'] = array('甲', '乙', '1', '2', '3', '4', '5', '6', '7', '8', '9');
            $data['rows'] = array();
            foreach ($data['line'] as $item) {
                $data['rows'][] = array($item[0][0]['count'], $item[1][0]['count'], $item[2][0]['count'], $item[3][0]['count'], $item[4][0]['count'], $item[5][0]['count'], $item[6][0]['count'], $item[7][0]['count'], $item[8][0]['count'], $item[9][0]['count'], $item[10][0]['count']);
            }
            break;
        case '初中':
            $data['title'] = "初中分年龄学生数";
            $data['sheet_name'] = "教基3116";
            $data['columns1'] = array('指标名称', '代码', '招生数', '在校生数', '#女', '一年级', '二年级', '三年级', '四年级');
            $data['columns2'] = array('甲', '乙', '1', '2', '3', '4', '5', '6', '7');
            $data['rows'] = array();
            foreach ($data['line'] as $item) {
                $data['rows'][] = array($item[0][0]['count'], $item[1][0]['count'], $item[2][0]['count'], $item[3][0]['count'], $item[4][0]['count'], $item[5][0]['count'], $item[6][0]['count'], $item[7][0]['count'], $item[8][0]['count']);
            }
            break;
        case '普通高中':
            $data['title'] = "高中阶段分年龄学生数";
            $data['sheet_name'] = "教基3119";
            $data['columns1'] = array('指标名称', '代码', '招生数', '在校生数', '#女', '#全日制学生', '一年级', '二年级', '三年级', '四年级及以上');
            $data['columns2'] = array('甲', '乙', '1', '2', '3', '4', '5', '6', '7', '8');
            $data['rows'] = array();
            foreach ($data['line'] as $item) {
                $data['rows'][] = array($item[0][0]['count'], $item[1][0]['count'], $item[2][0]['count'], $item[3][0]['count'], $item[4][0]['count'], $item[5][0]['count'], $item[6][0]['count'], $item[7][0]['count'], $item[8][0]['count'], $item[9][0]['count']);
            }
            break;
    }
    
    return $data;
}

/**
 * 导入教职工（使用内存数据）
 */
function import_teachers()
{
    // 检查文件上传
    if ($_FILES["file"]["error"] > 0) {
        exit("文件上传错误：" . $_FILES["file"]["error"]);
    } else {
        // 检查文件类型
        $fileType = $_FILES["file"]["type"];
        $allowedTypes = array("application/vnd.ms-excel", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        
        if (in_array($fileType, $allowedTypes) || 
            pathinfo($_FILES["file"]["name"], PATHINFO_EXTENSION) == 'xls' || 
            pathinfo($_FILES["file"]["name"], PATHINFO_EXTENSION) == 'xlsx') {
            
            // 根据文件类型选择读取器
            $reader = IOFactory::createReaderForFile($_FILES["file"]["tmp_name"]);
            
            // 载入excel文件
            $spreadsheet = $reader->load($_FILES["file"]["tmp_name"]);
            $sheet = $spreadsheet->getActiveSheet();
            
            $highestRow = $sheet->getHighestRow();
            
            // 在内存中存储教师数据
            $teachersData = array();
            $successCount = 0;
            $errorCount = 0;
            $errors = array();
            
            // 从第2行开始读取数据（第1行是标题）
            for ($row = 2; $row <= $highestRow; $row++) {
                // 兼容 PHP 8.1+：单元格为空时 getValue() 可能返回 null，直接 trim(null) 会触发 Deprecated
                $identity_number = trim((string)$sheet->getCell('C' . $row)->getValue()); // 身份证号码
                // 跳过空行
                if (empty($identity_number)) {
                    continue;
                }
                
                $personnel_type = trim($sheet->getCell('D' . $row)->getValue()); // 人员类型
                $minority = trim($sheet->getCell('E' . $row)->getValue()); // 少数民族
                $political_status = trim($sheet->getCell('F' . $row)->getValue()); // 政治面貌
                $highest_education = trim($sheet->getCell('G' . $row)->getValue()); // 最高学历
                $job_title = trim($sheet->getCell('H' . $row)->getValue()); // 职称
                $phase = trim($sheet->getCell('I' . $row)->getValue()); // 学段
                $subject = trim($sheet->getCell('J' . $row)->getValue()); // 科目
                
                // 从身份证号码计算年龄和性别
                $age = get_age_from_id_september($identity_number);
                $sex = get_sex_from_id_chinese($identity_number);
                
                // 准备数据（存储在内存中）
                $teacherData = array(
                    'id' => $successCount + 1,
                    'identity_number' => $identity_number,
                    'minority' => $minority,
                    'party' => $political_status,
                    'education' => $highest_education,
                    'profession' => $job_title,
                    'phase' => $phase,
                    'subject' => $subject,
                    'stafftype' => $personnel_type,
                    'age' => $age,
                    'sex' => $sex
                );
                
                // 添加到内存数组
                $teachersData[] = $teacherData;
                $successCount++;
            }
            
            // 有有效行时必须先下发 Excel：若先 echo HTML，attachment 头无效，内嵌浏览器无法触发下载
            if ($successCount > 0) {
                generate_teachers_statistics($teachersData);
            }
            
            echo "<h2>导入完成</h2>";
            echo "<p>成功导入：" . $successCount . " 条记录</p>";
            echo "<p>失败：" . $errorCount . " 条记录</p>";
            if ($successCount === 0) {
                echo "<p><strong>未生成 Excel：</strong>未读取到有效数据。请确认从第 2 行起 <strong>C 列</strong>为身份证号码且非空。</p>";
            }
            
            if (count($errors) > 0) {
                echo "<h3>错误信息：</h3>";
                foreach ($errors as $error) {
                    echo "<p>" . $error . "</p>";
                }
            }
            
            echo "<p><a href='?action=home'>返回首页</a></p>";
        } else {
            exit("文件名格式错误，请上传Excel文件（.xls或.xlsx格式）");
        }
    }
}

/**
 * 生成教职工数据统计报表（使用内存数据）
 */
function generate_teachers_statistics($teachersData)
{
    // 将多个报表数据组成数组，每个元素会生成一个工作表
    generate_excel(array(
        report_4149($teachersData),
        report_4153($teachersData),
        report_4155($teachersData),
        report_4067($teachersData)
    ), '年报导出(教职工)');
}

/**
 * 教基4149（使用内存数据）
 */
function report_4149($teachersData)
{
    $stafftype_arr = array('专任教师', '行政人员', '教辅人员', '工勤人员', '其他');
    
    $data['line'] = array();
    foreach ($stafftype_arr as $item) {
        $count = MemoryDataQuery::count($teachersData, array("stafftype" => $item));
        $data['line'][0][] = array(array("count" => $count));
        
        $count = MemoryDataQuery::count($teachersData, array("stafftype" => $item, "sex" => "女"));
        $data['line'][1][] = array(array("count" => $count));
        
        $count = MemoryDataQuery::count($teachersData, array("stafftype" => $item, "minority" => "是"));
        $data['line'][2][] = array(array("count" => $count));
        
        $data['line'][3][] = array(array("count" => "-"));
    }

    // 补齐“校外教师、银龄教师”两列默认值，避免后续按固定列索引取值时出现越界告警
    $data['line'][0][] = array(array("count" => 0));
    $data['line'][0][] = array(array("count" => 0));
    $data['line'][1][] = array(array("count" => 0));
    $data['line'][1][] = array(array("count" => 0));
    $data['line'][2][] = array(array("count" => 0));
    $data['line'][2][] = array(array("count" => 0));
    $data['line'][3][] = array(array("count" => "-"));
    $data['line'][3][] = array(array("count" => "-"));
    
    for ($i = 0; $i < count($data['line']); $i++) {
        $tmp = 0;
        foreach ($data['line'][$i] as $item) {
            if ($item[0]['count'] != "-") {
                $tmp = $tmp + $item[0]['count'];
            }
        }
        array_unshift($data['line'][$i], array(array("count" => $tmp)));
    }
    
    $data['title'] = '中小学教职工';
    $data['sheet_name'] = '教基4149';
    $data['columns1'] = array('指标名称', '代码', '教职工数', '专任教师', '行政人员', '教辅人员', '工勤人员', '其他', '校外教师', '银龄教师');
    $data['columns2'] = array('甲', '乙', '1', '2', '3', '4', '5', '6', '7', '8');
    $data['rows'][] = array('总计', '01', $data['line'][0][0][0]['count'], $data['line'][0][1][0]['count'], $data['line'][0][2][0]['count'], $data['line'][0][3][0]['count'], $data['line'][0][4][0]['count'], $data['line'][0][5][0]['count'], $data['line'][0][6][0]['count'], $data['line'][0][7][0]['count']);
    $data['rows'][] = array('#女', '02', $data['line'][1][0][0]['count'], $data['line'][1][1][0]['count'], $data['line'][1][2][0]['count'], $data['line'][1][3][0]['count'], $data['line'][1][4][0]['count'], $data['line'][1][5][0]['count'], $data['line'][1][6][0]['count'], $data['line'][1][7][0]['count']);
    $data['rows'][] = array('#少数民族', '03', $data['line'][2][0][0]['count'], $data['line'][2][1][0]['count'], $data['line'][2][2][0]['count'], $data['line'][2][3][0]['count'], $data['line'][2][4][0]['count'], $data['line'][2][5][0]['count'], $data['line'][2][6][0]['count'], $data['line'][2][7][0]['count']);
    $data['rows'][] = array('#在编人员', '04');
    $data['rows'][] = array('#外籍人员', '05');
    return $data;
}

/**
 * 教基4153（使用内存数据）
 */
function report_4153($teachersData)
{
    // 先过滤出专任教师
    $teacherData = MemoryDataQuery::filter($teachersData, array("stafftype" => '专任教师'));
    
    $phase_arr = array('学前教育', '小学', '初中', '普通高中', '特殊教育学校');
    $profession_arr = array('正高级', '副高级', '中级', '助理级', '员级', '未定职级');
    
    $age_sql_arr[0] = array();
    $age_sql_arr[1] = array();
    $age_sql_arr[2] = array("age<=" => 24);
    $age_sql_arr[3] = array("age>=" => 25, "age<=" => 29);
    $age_sql_arr[4] = array("age>=" => 30, "age<=" => 34);
    $age_sql_arr[5] = array("age>=" => 35, "age<=" => 39);
    $age_sql_arr[6] = array("age>=" => 40, "age<=" => 44);
    $age_sql_arr[7] = array("age>=" => 45, "age<=" => 49);
    $age_sql_arr[8] = array("age>=" => 50, "age<=" => 54);
    $age_sql_arr[9] = array("age>=" => 55, "age<=" => 59);
    $age_sql_arr[10] = array("age>=" => 60);
    
    $data['line'] = array();
    
    $col_1_arr = array();    //指标名称
    $col_2_arr = array();    //指标代码
    $col_2_starter = 0;
    $row_where_arr = array();
    
    foreach ($phase_arr as $item) {
        $col_1_arr[] = array("count" => $item);
        $row_where_arr[] = array("phase" => $item);
        $col_2_starter++;
        $col_2_arr[] = array("count" => $col_2_starter);
        $col_1_arr[] = array("count" => " #女");
        $row_where_arr[] = array("sex" => "女", "phase" => $item);
        $col_2_starter++;
        $col_2_arr[] = array("count" => $col_2_starter);
        $col_1_arr[] = array("count" => " #少数民族");
        $row_where_arr[] = array("minority" => '是', "phase" => $item);
        $col_2_starter++;
        $col_2_arr[] = array("count" => $col_2_starter);
        foreach ($profession_arr as $item2) {
            $col_1_arr[] = array("count" => $item2);
            $row_where_arr[] = array("profession" => $item2, "phase" => $item);
            $col_2_starter++;
            $col_2_arr[] = array("count" => $col_2_starter);
        }
    }
    
    $where_arr = array();
    
    for ($row = 0; $row < count($col_2_arr); $row++) {
        for ($col = 0; $col < count($age_sql_arr); $col++) {
            $dash = false;
            reset($where_arr);
            $where_arr = $age_sql_arr[$col] + $row_where_arr[$row];
            if ($col == 1) {
                $where_arr = $where_arr + array("sex" => "女");
                if (isset($row_where_arr[$row]["sex"]) && $row_where_arr[$row]["sex"] == "女") {
                    $dash = true;
                }
            }
            
            if ($dash == true) {
                $data['line'][$row][$col] = array(array("count" => "-"));
            } else {
                $count = MemoryDataQuery::count($teacherData, $where_arr);
                $data['line'][$row][$col] = array(array("count" => $count));
            }
        }
        array_unshift($data['line'][$row], array($col_2_arr[$row]));
        array_unshift($data['line'][$row], array($col_1_arr[$row]));
    }
    
    $data['title'] = "基础教育学校专任教师分年龄情况";
    $data['sheet_name'] = "教基4153";
    $data['columns1'] = array('指标名称', '代码', '合计', '#女', '24岁及以下', '25-29岁', '30-34岁', '35-39岁', '40-44岁', '45-49岁', '50-54岁', '55-59岁', '60岁及以上');
    $data['columns2'] = array('甲', '乙', '1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11');
    $data['rows'] = array();
    foreach ($data['line'] as $item) {
        $data['rows'][] = array($item[0][0]['count'], $item[1][0]['count'], $item[2][0]['count'], $item[3][0]['count'], $item[4][0]['count'], $item[5][0]['count'], $item[6][0]['count'], $item[7][0]['count'], $item[8][0]['count'], $item[9][0]['count'], $item[10][0]['count'], $item[11][0]['count'], $item[12][0]['count']);
    }
    
    return $data;
}

/**
 * 教基4155（使用内存数据）
 */
function report_4155($teachersData)
{
    // 先过滤出专任教师
    $teacherData = MemoryDataQuery::filter($teachersData, array("stafftype" => '专任教师'));
    
    $phase_arr = array('小学', '初中');
    $education_arr = array('博士研究生', '硕士研究生', '本科', '专科', '高中阶段', '高中阶段以下');
    $subject_arr = array("道德与法制", "语文", "数学", array("英语", "日语", "俄语"), "英语", "日语", "俄语", "历史", "地理", "科学", "物理", "化学", "生物学", "信息科技", "体育与健康", array("音乐", "美术"), "音乐", "美术", "劳动", "综合实践活动", "其他", "本学年不授课");
    
    $data['line'] = array();
    
    $subject_sql_arr = array();
    foreach ($subject_arr as $item) {
        if (is_array($item)) {
            $subject_sql_arr[] = array("subject" => $item);
        } else {
            $subject_sql_arr[] = array("subject" => array($item));
        }
    }
    array_unshift($subject_sql_arr, array());
    array_unshift($subject_sql_arr, array());
    
    $col_1_arr = array();    //指标名称
    $col_2_arr = array();    //指标代码
    $col_2_starter = 0;
    $row_where_arr = array();
    
    $col_1_arr[] = array("count" => "总计");
    $row_where_arr[] = array(1 => 1);  // 匹配所有
    $col_2_starter++;
    
    foreach ($phase_arr as $item) {
        $col_1_arr[] = array("count" => $item);
        $row_where_arr[] = array("phase" => $item);
        $col_2_starter++;
        $col_2_arr[] = array("count" => $col_2_starter);
        $col_1_arr[] = array("count" => " #女");
        $row_where_arr[] = array("sex" => "女", "phase" => $item);
        $col_2_starter++;
        $col_2_arr[] = array("count" => $col_2_starter);
        $col_1_arr[] = array("count" => " #少数民族");
        $row_where_arr[] = array("minority" => '是', "phase" => $item);
        $col_2_starter++;
        $col_2_arr[] = array("count" => $col_2_starter);
        foreach ($education_arr as $item2) {
            $col_1_arr[] = array("count" => $item2);
            $row_where_arr[] = array("education" => $item2, "phase" => $item);
            $col_2_starter++;
            $col_2_arr[] = array("count" => $col_2_starter);
        }
    }
    
    $where_arr = array();
    $where_in_arr = array();
    
    for ($row = 0; $row < count($col_2_arr); $row++) {
        for ($col = 0; $col < count($subject_sql_arr); $col++) {
            $dash = false;
            reset($where_arr);
            reset($where_in_arr);
            $where_arr = $row_where_arr[$row];
            $where_in_arr = $subject_sql_arr[$col];
            if ($col == 1) {
                $where_arr = $where_arr + array("sex" => "女");
                if (isset($row_where_arr[$row]["sex"]) && $row_where_arr[$row]["sex"] == "女") {
                    $dash = true;
                }
            }
            
            if ($dash == true) {
                $data['line'][$row][$col] = array(array("count" => "-"));
            } else {
                $count = MemoryDataQuery::count($teacherData, $where_arr, $where_in_arr);
                $data['line'][$row][$col] = array(array("count" => $count));
            }
        }
        array_unshift($data['line'][$row], array($col_2_arr[$row]));
        array_unshift($data['line'][$row], array($col_1_arr[$row]));
    }
    
    $data['title'] = "义务教育阶段专任教师分课程、分学历情况";
    $data['sheet_name'] = "教基4155";
    $data['columns1'] = array('指标名称', '代码', '合计', '#女', '道德与法制', '语文', '数学', '外语', '英语', '日语', '俄语', '历史', '地理', '科学', '物理', '化学', '生物学', '信息科技', '体育与健康', '艺术', '音乐', '美术', '劳动', '综合实践活动', '其他', '本学年不授课');
    $data['columns2'] = array('甲', '乙', '1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11', '12', '13', '14', '15', '16', '17', '18', '19', '20', '21', '22', '23', '24');
    $data['rows'] = array();
    foreach ($data['line'] as $item) {
        $data['rows'][] = array($item[0][0]['count'], $item[1][0]['count'], $item[2][0]['count'], $item[3][0]['count'], $item[4][0]['count'], $item[5][0]['count'], $item[6][0]['count'], $item[7][0]['count'], $item[8][0]['count'], $item[9][0]['count'], $item[10][0]['count'], $item[11][0]['count'], $item[12][0]['count'], $item[13][0]['count'], $item[14][0]['count'], $item[15][0]['count'], $item[16][0]['count'], $item[17][0]['count'], $item[18][0]['count'], $item[19][0]['count'], $item[20][0]['count'], $item[21][0]['count'], $item[22][0]['count'], $item[23][0]['count'], $item[24][0]['count'], $item[25][0]['count']);
    }
    
    return $data;
}

/**
 * 教基4067（使用内存数据）
 */
function report_4067($teachersData)
{
    $phase_arr = array('学前教育', '小学', '初中', '普通高中', '特殊教育学校', '中等职业学校', '高等教育学校');
    $party_arr = array('中共党员', '共青团员', '民主党派', '香港', '澳门', '台湾', '华侨');
    
    $data['line'] = array();
    
    $col_1_arr = array();
    $col_2_arr = array();
    $col_2_starter = 0;
    $row_where_arr = array();
    
    $col_1_arr[] = array("count" => "教职工");
    $row_where_arr[] = array(1 => 1);
    $col_2_starter++;
    $col_2_arr[] = array("count" => $col_2_starter);
    
    $col_1_arr[] = array("count" => " #女");
    $row_where_arr[] = array("sex" => "女");
    $col_2_starter++;
    $col_2_arr[] = array("count" => $col_2_starter);
    
    $col_1_arr[] = array("count" => "专任教师");
    $row_where_arr[] = array("stafftype" => '专任教师');
    $col_2_starter++;
    $col_2_arr[] = array("count" => $col_2_starter);
    
    foreach ($phase_arr as $item) {
        $col_1_arr[] = array("count" => $item);
        $row_where_arr[] = array("phase" => $item, "stafftype" => '专任教师');
        $col_2_starter++;
        $col_2_arr[] = array("count" => $col_2_starter);
        $col_1_arr[] = array("count" => " #女");
        $row_where_arr[] = array("sex" => "女", "phase" => $item, "stafftype" => '专任教师');
        $col_2_starter++;
        $col_2_arr[] = array("count" => $col_2_starter);
    }
    
    $where_arr = array();
    for ($row = 0; $row < count($col_2_arr); $row++) {
        for ($col = 0; $col < count($party_arr); $col++) {
            reset($where_arr);
            $where_arr = $row_where_arr[$row] + array("party" => $party_arr[$col]);
            $count = MemoryDataQuery::count($teachersData, $where_arr);
            $data['line'][$row][$col] = array(array("count" => $count));
        }
        reset($where_arr);
        $where_arr = $row_where_arr[$row] + array("minority" => '是');
        $count = MemoryDataQuery::count($teachersData, $where_arr);
        $data['line'][$row][$col + 1] = array(array("count" => $count));
        array_unshift($data['line'][$row], array($col_2_arr[$row]));
        array_unshift($data['line'][$row], array($col_1_arr[$row]));
    }
    
    $data['title'] = "教职工其他情况";
    $data['sheet_name'] = "教基4067";
    $data['columns1'] = array('指标名称', '代码', '中共党员', '共青团员', '民主党派', '香港', '澳门', '台湾', '华侨', '少数民族');
    $data['columns2'] = array('甲', '乙', '1', '2', '3', '4', '5', '6', '7', '8');
    $data['rows'] = array();
    foreach ($data['line'] as $item) {
        $data['rows'][] = array($item[0][0]['count'], $item[1][0]['count'], $item[2][0]['count'], $item[3][0]['count'], $item[4][0]['count'], $item[5][0]['count'], $item[6][0]['count'], $item[7][0]['count'], $item[8][0]['count'], $item[9][0]['count']);
    }
    
    return $data;
}

/**
 * 根据数据数组生成Excel文件，每个数组元素生成一个工作表
 * 使用 PhpSpreadsheet 代替 PHPExcel
 */
function generate_excel($dataArray, $filename)
{
    // 检查数据是否有效
    if (empty($dataArray) || !is_array($dataArray)) {
        exit('数据无效，无法生成Excel文件');
    }
    
    // 创建新的Spreadsheet对象
    $spreadsheet = new Spreadsheet();
    
    // 循环处理数组中的每个元素，每个元素生成一个工作表
    foreach ($dataArray as $sheetIndex => $data) {
        // 对于除第一个外的所有工作表，需要创建新工作表
        if ($sheetIndex > 0) {
            $spreadsheet->createSheet();
        }
        
        // 设置当前活动工作表
        $spreadsheet->setActiveSheetIndex($sheetIndex);
        $sheet = $spreadsheet->getActiveSheet();
        
        // 设置工作表标题
        $sheetTitle = isset($data['sheet_name']) ? $data['sheet_name'] : '统计报表' . ($sheetIndex + 1);
        $sheet->setTitle($sheetTitle);
        
        // 设置文档属性（只在第一个工作表时设置）
        if ($sheetIndex == 0) {
            $spreadsheet->getProperties()
                ->setCreator("年报填报系统")
                ->setLastModifiedBy("年报填报系统")
                ->setTitle(isset($data['title']) ? $data['title'] : "统计报表")
                ->setSubject("数据统计")
                ->setDescription("根据数据生成的统计报表");
        }
        
        // 获取列数（用于确定标题合并范围）
        // PhpSpreadsheet的列索引从1开始，所以直接使用列数
        $maxCol = isset($data['columns1']) ? count($data['columns1']) : 10;
        $maxColLetter = Coordinate::stringFromColumnIndex($maxCol);
        
        // 写入标题行（第一行）
        if (isset($data['title']) && !empty($data['title'])) {
            // 设置标题内容
            $sheet->setCellValue('A1', $data['title']);
            // 合并标题单元格，宽度为 columns1 的数量
            $sheet->mergeCells('A1:' . $maxColLetter . '1');
            // 设置标题样式
            $titleStyle = array(
                'font' => array('bold' => true, 'size' => 14),
                'alignment' => array(
                    'horizontal' => Alignment::HORIZONTAL_CENTER,
                    'vertical' => Alignment::VERTICAL_CENTER
                ),
                'borders' => array('allBorders' => array('borderStyle' => Border::BORDER_THIN))
            );
            $sheet->getStyle('A1:' . $maxColLetter . '1')->applyFromArray($titleStyle);
        }
        
        // 表头1（第二行）
        if (isset($data['columns1']) && is_array($data['columns1'])) {
            $colIndex = 1;
            foreach ($data['columns1'] as $column) {
                $cellAddress = Coordinate::stringFromColumnIndex($colIndex) . '2';
                $sheet->setCellValue($cellAddress, $column);
                $colIndex++;
            }
        }
        
        // 表头2（第三行）
        if (isset($data['columns2']) && is_array($data['columns2'])) {
            $colIndex = 1;
            foreach ($data['columns2'] as $column) {
                $cellAddress = Coordinate::stringFromColumnIndex($colIndex) . '3';
                $sheet->setCellValue($cellAddress, $column);
                $colIndex++;
            }
        }
        
        // 数据行（从第四行开始）
        if (isset($data['rows']) && is_array($data['rows'])) {
            $rowIndex = 4;
            foreach ($data['rows'] as $row) {
                $colIndex = 1;
                foreach ($row as $cellValue) {
                    $cellAddress = Coordinate::stringFromColumnIndex($colIndex) . $rowIndex;
                    $finalValue = get_final_value($cellValue);
                    $sheet->setCellValue($cellAddress, $finalValue);
                    $colIndex++;
                }
                $rowIndex++;
            }
        }
        
        // 设置列宽
        for ($i = 1; $i <= $maxCol; $i++) {
            $colLetter = Coordinate::stringFromColumnIndex($i);
            if ($i == 1) {
                $sheet->getColumnDimension($colLetter)->setWidth(15);
            } elseif ($i == 2) {
                $sheet->getColumnDimension($colLetter)->setWidth(8);
            } else {
                $sheet->getColumnDimension($colLetter)->setWidth(12);
            }
        }
        
        // 设置样式
        $headerStyle = array(
            'font' => array('bold' => true, 'size' => 12),
            'alignment' => array(
                'horizontal' => Alignment::HORIZONTAL_CENTER,
                'vertical' => Alignment::VERTICAL_CENTER
            ),
            'borders' => array('allBorders' => array('borderStyle' => Border::BORDER_THIN))
        );
        $dataStyle = array(
            'alignment' => array(
                'horizontal' => Alignment::HORIZONTAL_CENTER,
                'vertical' => Alignment::VERTICAL_CENTER
            ),
            'borders' => array('allBorders' => array('borderStyle' => Border::BORDER_THIN))
        );
        
        // 应用表头样式（第二行和第三行）
        $sheet->getStyle('A2:' . $maxColLetter . '3')->applyFromArray($headerStyle);
        
        // 应用数据样式（从第四行开始）
        $rowCount = isset($data['rows']) ? count($data['rows']) : 0;
        if ($rowCount > 0) {
            $sheet->getStyle('A4:' . $maxColLetter . (4 + $rowCount - 1))->applyFromArray($dataStyle);
        }
    }
    
    // 清空输出缓冲，避免与 HTML/二进制响应混杂
    while (ob_get_level() > 0) {
        ob_end_clean();
    }
    
    $timePart = date('Ymd_His');
    // 导出文件名：中文业务名 + 时间戳
    $reportTitle = $filename;
    $outName = $filename . '_' . $timePart . '.xlsx';
    $isPhpDesktop = is_php_desktop_runtime();
    $outPath = null;
    
    try {
        $writer = new Xlsx($spreadsheet);
        if ($isPhpDesktop) {
            // 桌面版：写入应用根目录（与 www 同级），避免内嵌浏览器下载不稳定
            $projectRoot = realpath(dirname(__DIR__));
            if ($projectRoot === false) {
                $projectRoot = dirname(__DIR__);
            }
            $outPath = $projectRoot . DIRECTORY_SEPARATOR . $outName;
            $writer->save($outPath);
        } else {
            // 在线版：不写服务器磁盘，直接附件下载（避免无写权限、路径暴露及与代码目录混放）
            header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
            header('Content-Disposition: attachment; filename*=UTF-8\'\'' . rawurlencode($outName));
            header('Cache-Control: max-age=0');
            $writer->save('php://output');
            $spreadsheet->disconnectWorksheets();
            unset($spreadsheet);
            exit;
        }
    } catch (Throwable $e) {
        $spreadsheet->disconnectWorksheets();
        unset($spreadsheet);
        header('Content-Type: text/html; charset=UTF-8');
        echo '<!DOCTYPE html><html lang="zh-CN"><head><meta charset="UTF-8"><title>导出失败</title>';
        echo '<style>body{font-family:system-ui,sans-serif;margin:0;padding:24px;background:#fef2f2;}.box{max-width:560px;margin:0 auto;background:#fff;border:2px solid #ef4444;border-radius:12px;padding:24px;box-shadow:0 10px 40px rgba(239,68,68,.12);}.h{font-size:1.25rem;font-weight:700;color:#991b1b;margin:0 0 12px;}code{font-size:.85rem;word-break:break-all;}a{color:#b91c1c;}</style></head><body>';
        echo '<div class="box"><p class="h">导出失败</p>';
        echo '<p>' . ($isPhpDesktop ? 'Excel 未能写入应用目录' : 'Excel 未能输出') . '（请确认已启用 php_zip 等扩展）：</p><p><code>' . htmlspecialchars($e->getMessage(), ENT_QUOTES, 'UTF-8') . '</code></p>';
        if ($isPhpDesktop && $outPath !== null) {
            echo '<p>目标路径：<code>' . htmlspecialchars($outPath, ENT_QUOTES, 'UTF-8') . '</code></p>';
        }
        $failPayload = array(
            'type' => 'edureport-export',
            'ok' => false,
            'message' => $e->getMessage(),
            'path' => $outPath !== null ? $outPath : '',
        );
        echo '<script>try{if(window.parent!==window){window.parent.postMessage(' . json_encode($failPayload, JSON_UNESCAPED_UNICODE | JSON_HEX_TAG) . ',"*");}}catch(err){}</script>';
        echo '<p><a href="?action=home">返回首页</a></p></div></body></html>';
        exit;
    }
    
    $spreadsheet->disconnectWorksheets();
    unset($spreadsheet);
    
    // 仅桌面版执行到此：结果页 + postMessage
    $fileOk = is_file($outPath) && filesize($outPath) > 0;
    $successPayload = array(
        'type' => 'edureport-export',
        'ok' => true,
        'verified' => $fileOk,
        'title' => $reportTitle,
        'path' => $outPath,
        'filename' => basename($outPath),
        'mode' => 'desktop',
    );
    
    header('Content-Type: text/html; charset=UTF-8');
    echo '<!DOCTYPE html><html lang="zh-CN"><head><meta charset="UTF-8"><title>报表已生成</title>';
    echo '<style>body{font-family:system-ui,sans-serif;margin:0;padding:24px;background:#f0fdf4;}.box{max-width:560px;margin:0 auto;background:#fff;border:2px solid #22c55e;border-radius:12px;padding:24px;box-shadow:0 10px 40px rgba(34,197,94,.15);}.h{font-size:1.35rem;font-weight:700;color:#166534;margin:0 0 12px;display:flex;align-items:center;gap:8px;}.path{font-size:.9rem;word-break:break-all;background:#f7fee7;padding:12px;border-radius:8px;border:1px solid #bbf7d0;margin:12px 0;}a{color:#15803d;}</style></head><body>';
    echo '<div class="box">';
    echo '<p class="h"><span aria-hidden="true">✓</span> 导出完成</p>';
    echo '<p>报表类型：<strong>' . htmlspecialchars($reportTitle, ENT_QUOTES, 'UTF-8') . '</strong></p>';
    if ($fileOk) {
        echo '<p>文件已保存到应用根目录：</p><div class="path">' . htmlspecialchars($outPath, ENT_QUOTES, 'UTF-8') . '</div>';
    } else {
        echo '<p><strong>警告：</strong>未检测到有效文件，请检查磁盘权限与杀毒软件是否拦截。</p><div class="path">' . htmlspecialchars($outPath, ENT_QUOTES, 'UTF-8') . '</div>';
    }
    echo '<p><a href="?action=home">返回首页</a></p></div>';
    echo '<script>try{if(window.parent!==window){window.parent.postMessage(' . json_encode($successPayload, JSON_UNESCAPED_UNICODE | JSON_HEX_TAG) . ',"*");}}catch(err){}</script>';
    echo '</body></html>';
    exit;
}

