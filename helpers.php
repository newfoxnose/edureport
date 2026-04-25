<?php
/**
 * 辅助函数文件
 */


/**
 * 根据身份证号码计算年龄（以9月1日为分界线）
 */
function get_age_from_id_september($identity_number)
{
    $age = 0;
    if (!empty($identity_number) && strlen($identity_number) == 18) {
        // 提取生日信息
        $birth_year = intval(substr($identity_number, 6, 4));
        $birth_month = intval(substr($identity_number, 10, 2));
        $current_year = date('Y');
        $current_month = date('m');

        // 计算年龄的参考年份（以9月1日为分界线）
        if ($current_month < 9) {
            $reference_year = $current_year - 1;
        } else {
            $reference_year = $current_year;
        }

        // 计算年龄
        if ($birth_month <= 9) {
            $age = $reference_year - $birth_year;
        } else {
            $age = $reference_year - $birth_year - 1;
        }
    }
    return $age;
}

/**
 * 根据身份证号码获取性别（返回中文）
 */
function get_sex_from_id_chinese($identity_number)
{
    $sex = '';
    if (!empty($identity_number) && strlen($identity_number) == 18) {
        // 获取性别（身份证倒数第二位数字，奇数为男，偶数为女）
        $gender_digit = intval(substr($identity_number, 16, 1));
        if ($gender_digit % 2 == 0) {
            $sex = '女';
        } else {
            $sex = '男';
        }
    }
    return $sex;
}

/**
 * 递归获取数组的最终值（用于处理多维数组）
 */
function get_final_value($value)
{
    // 处理null值
    if (is_null($value)) {
        return '';
    }

    // 处理标量值
    if (!is_array($value)) {
        return $value;
    }

    // 处理数组
    // 如果数组有count键，直接返回count值
    if (isset($value['count'])) {
        return is_numeric($value['count']) ? $value['count'] : (string)$value['count'];
    }

    // 如果数组有0键，递归处理
    if (isset($value[0])) {
        return get_final_value($value[0]);
    }

    // 空数组返回空字符串
    return '';
}

/**
 * 去除身份证号中的所有非数字符号，只保留 0-9 和字母 X（校验码）
 * @param string $identity_number 原始身份证号
 * @return string 仅包含数字及 X 的字符串
 */
function trim_identity_number($identity_number)
{
    // 小写 x 替换为大写 X（身份证校验码）
    $identity_number = str_replace('x', 'X', $identity_number);
    // 去除所有非数字且非 X 的字符（空格、点、横线、引号等）
    return preg_replace('/[^0-9X]/', '', $identity_number);
}

//去除字符串中的空格、制表符、换行符、回车符
function trim_str($str)
{
    return preg_replace('/[\s\t\n\r]+/', '', $str);
}




/**
 * 验证中国大陆18位身份证号码是否正确
 * @param string $idCard 身份证号码
 * @return bool 验证结果：true-正确，false-错误
 */
function validateIDCard($idCard)
{
    // 基础格式验证：长度18位，只能包含数字和X（大小写均可）
    if (!preg_match('/^\d{17}[\dXx]$/', $idCard)) {
        return false;
    }

    // 验证前两位省份代码（11-15, 21-23, 31-37, 41-46, 50-54, 61-65, 81-82, 71）
    $provinceCode = substr($idCard, 0, 2);
    $validProvinces = [
        '11',
        '12',
        '13',
        '14',
        '15',      // 华北地区
        '21',
        '22',
        '23',                  // 东北地区  
        '31',
        '32',
        '33',
        '34',
        '35',
        '36',
        '37', // 华东地区
        '41',
        '42',
        '43',
        '44',
        '45',
        '46', // 华中、华南地区
        '50',
        '51',
        '52',
        '53',
        '54',      // 西南地区
        '61',
        '62',
        '63',
        '64',
        '65',      // 西北地区
        '81',
        '82',                        // 香港、澳门特别行政区
        '71'                               // 台湾省
    ];

    if (!in_array($provinceCode, $validProvinces)) {
        return false;
    }

    // 验证出生日期（格式验证 + 范围验证：1900-01-01 至当天）
    $year = substr($idCard, 6, 4);
    $month = substr($idCard, 10, 2);
    $day = substr($idCard, 12, 2);

    if (!checkdate($month, $day, $year)) {
        return false;
    }

    $birth = substr($idCard, 6, 8); // YYYYMMDD
    if ($birth < '19000101' || $birth > date('Ymd')) {
        return false;
    }

    // 验证校验码
    return validateCheckCode($idCard);
}

/**
 * 验证身份证校验码
 * @param string $idCard 身份证号码
 * @return bool 校验结果
 */
function validateCheckCode($idCard)
{
    // 加权因子
    $factors = [7, 9, 10, 5, 8, 4, 2, 1, 6, 3, 7, 9, 10, 5, 8, 4, 2];

    // 校验码对应表
    $checkCodes = ['1', '0', 'X', '9', '8', '7', '6', '5', '4', '3', '2'];

    $sum = 0;

    // 计算前17位数字的加权和
    for ($i = 0; $i < 17; $i++) {
        $sum += intval($idCard[$i]) * $factors[$i];
    }

    // 取模计算校验码索引
    $mod = $sum % 11;

    // 获取计算的校验码（转换为大写进行比较）
    $calculatedCode = $checkCodes[$mod];
    $actualCode = strtoupper($idCard[17]); // 将最后一位转为大写

    return $calculatedCode === $actualCode;
}

/**
 * 是否运行在 PHP Desktop 环境（宿主会设置 PHPDESKTOP_VERSION）
 * 在线服务器一般无此变量：用于区分「写入应用根目录」与「HTTP 附件下载」
 */
function is_php_desktop_runtime()
{
    $v = getenv('PHPDESKTOP_VERSION');
    if ($v !== false && $v !== '') {
        return true;
    }
    if (!empty($_SERVER['PHPDESKTOP_VERSION'])) {
        return true;
    }
    return false;
}
