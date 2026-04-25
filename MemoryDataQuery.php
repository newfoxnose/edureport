<?php
/**
 * 内存数据查询工具类 - 在内存数组中进行查询和统计
 */
class MemoryDataQuery
{
    /**
     * 在内存数据中查询符合条件的记录数
     */
    public static function count($data, $where_arr = null, $where_in_arr = null)
    {
        $filtered = self::filter($data, $where_arr, $where_in_arr);
        return count($filtered);
    }
    
    /**
     * 在内存数据中查询符合条件的记录
     */
    public static function filter($data, $where_arr = null, $where_in_arr = null)
    {
        $result = $data;
        
        if ($where_arr !== null && !empty($where_arr)) {
            $result = self::filterWhere($result, $where_arr);
        }
        
        if ($where_in_arr !== null && !empty($where_in_arr)) {
            $result = self::filterWhereIn($result, $where_in_arr);
        }
        
        return $result;
    }
    
    /**
     * WHERE条件过滤
     */
    private static function filterWhere($data, $where_arr)
    {
        if (empty($where_arr) || !is_array($where_arr)) {
            return $data;
        }
        
        $result = [];
        foreach ($data as $item) {
            $match = true;
            foreach ($where_arr as $key => $value) {
                // 跳过特殊键（如数字键1=>1表示匹配所有）
                if (is_numeric($key) && $key == 1 && $value == 1) {
                    continue; // 匹配所有记录
                }
                
                // 处理比较操作符
                if (strpos($key, '<=') !== false) {
                    $field = str_replace('<=', '', $key);
                    if (!isset($item[$field]) || $item[$field] > $value) {
                        $match = false;
                        break;
                    }
                } elseif (strpos($key, '>=') !== false) {
                    $field = str_replace('>=', '', $key);
                    if (!isset($item[$field]) || $item[$field] < $value) {
                        $match = false;
                        break;
                    }
                } elseif (strpos($key, '<') !== false && strpos($key, '<=') === false) {
                    $field = str_replace('<', '', $key);
                    if (!isset($item[$field]) || $item[$field] >= $value) {
                        $match = false;
                        break;
                    }
                } elseif (strpos($key, '>') !== false && strpos($key, '>=') === false) {
                    $field = str_replace('>', '', $key);
                    if (!isset($item[$field]) || $item[$field] <= $value) {
                        $match = false;
                        break;
                    }
                } else {
                    // 普通相等匹配
                    if (!isset($item[$key]) || $item[$key] != $value) {
                        $match = false;
                        break;
                    }
                }
            }
            if ($match) {
                $result[] = $item;
            }
        }
        return $result;
    }
    
    /**
     * WHERE IN条件过滤
     */
    private static function filterWhereIn($data, $where_in_arr)
    {
        if (empty($where_in_arr) || !is_array($where_in_arr)) {
            return $data;
        }
        
        $result = [];
        foreach ($data as $item) {
            $match = true;
            foreach ($where_in_arr as $key => $values) {
                if (!isset($item[$key])) {
                    $match = false;
                    break;
                }
                // 处理数组值的情况
                if (is_array($values)) {
                    // 如果values是数组，检查item[$key]是否在数组中
                    if (!in_array($item[$key], $values)) {
                        $match = false;
                        break;
                    }
                } else {
                    // 如果values不是数组，直接比较
                    if ($item[$key] != $values) {
                        $match = false;
                        break;
                    }
                }
            }
            if ($match) {
                $result[] = $item;
            }
        }
        return $result;
    }
}

