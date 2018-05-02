<?php
$sub_menu = '400300';
include_once('./_common.php');

// 상품이 많을 경우 대비 설정변경
set_time_limit ( 0 );
ini_set('memory_limit', '50M');

auth_check($auth[$sub_menu], "w");

function only_number($n)
{
    return preg_replace('/[^0-9]/', '', $n);
}

if($_FILES['excelfile']['tmp_name']) {
    $file = $_FILES['excelfile']['tmp_name'];
    
    include_once(G5_LIB_PATH.'/Excel/reader.php');
    
    $data = new Spreadsheet_Excel_Reader();
    
    // Set output Encoding.
    $data->setOutputEncoding('UTF-8');
    
    /***
     * if you want you can change 'iconv' to mb_convert_encoding:
     * $data->setUTFEncoder('mb');
     *
     **/
    
    /***
     * By default rows & cols indeces start with 1
     * For change initial index use:
     * $data->setRowColOffset(0);
     *
     **/
    
    
    
    /***
     *  Some function for formatting output.
     * $data->setDefaultFormat('%.2f');
     * setDefaultFormat - set format for columns with unknown formatting
     *
     * $data->setColumnFormat(4, '%.3f');
     * setColumnFormat - set format for column (apply only to number fields)
     *
     **/
    
    $data->read($file);
    
    /*
    
    
    $data->sheets[0]['numRows'] - count rows
    $data->sheets[0]['numCols'] - count columns
    $data->sheets[0]['cells'][$i][$j] - data from $i-row $j-column
    
    $data->sheets[0]['cellsInfo'][$i][$j] - extended info about cell
    
    $data->sheets[0]['cellsInfo'][$i][$j]['type'] = "date" | "number" | "unknown"
    if 'type' == "unknown" - use 'raw' value, because  cell contain value with format '0.00';
    $data->sheets[0]['cellsInfo'][$i][$j]['raw'] = value if cell without format
    $data->sheets[0]['cellsInfo'][$i][$j]['colspan']
    $data->sheets[0]['cellsInfo'][$i][$j]['rowspan']
    */
    
    error_reporting(E_ALL ^ E_NOTICE);
    
    $dup_it_id = array();
    $fail_it_id = array();
    $dup_count = 0;
    $total_count = 0;
    $fail_count = 0;
    $succ_count = 0;
    
    for ($i = 3; $i <= $data->sheets[0]['numRows']; $i++) {
        $total_count++;
        
        $j =2;

        
        
        // 새로운 주문번호 생성
        $od_id = get_uniqid();
/*        
        od_id
        mb_id
        it_id
        it_name
        it_sc_type
        it_sc_method
        it_sc_price
        it_sc_minimum
        it_sc_qty
        ct_status
        ct_price
        ct_point
        ct_point_use
        ct_stock_use
        ct_option
        ct_qty
        ct_notax
        io_id
        io_type
        io_price
        ct_time
        ct_ip
        ct_send_cost
        ct_direct
        ct_select
        ct_select_time 
*/        
        
        $sc_od              = clean_xss_tags(addslashes($data->sheets[0]['cells'][$i][$j++]));
        $sc_date            = clean_xss_tags(addslashes($data->sheets[0]['cells'][$i][$j++]));
        $sc_join            = clean_xss_tags(addslashes($data->sheets[0]['cells'][$i][$j++]));
        $sc_delivery_company= clean_xss_tags(addslashes($data->sheets[0]['cells'][$i][$j++]));
        $sc_delivery_num    = clean_xss_tags(addslashes($data->sheets[0]['cells'][$i][$j++]));
        $sc_delivery_reg    = clean_xss_tags(addslashes($data->sheets[0]['cells'][$i][$j++]));
        $sc_od_name         = clean_xss_tags(addslashes($data->sheets[0]['cells'][$i][$j++]));
        $sc_item_id         = clean_xss_tags(addslashes($data->sheets[0]['cells'][$i][$j++]));
        $sc_item_name       = clean_xss_tags(addslashes($data->sheets[0]['cells'][$i][$j++]));
        $sc_item_op_id      = clean_xss_tags(addslashes($data->sheets[0]['cells'][$i][$j++]));
        $sc_item_op_name    = clean_xss_tags(addslashes($data->sheets[0]['cells'][$i][$j++]));
        $sc_od_qty          = clean_xss_tags(addslashes($data->sheets[0]['cells'][$i][$j++]));
        $sc_od_price        = clean_xss_tags(addslashes($data->sheets[0]['cells'][$i][$j++]));
        $sc_od_b_name       = clean_xss_tags(addslashes($data->sheets[0]['cells'][$i][$j++]));
        $sc_od_b_hp         = clean_xss_tags(addslashes($data->sheets[0]['cells'][$i][$j++]));
        $sc_od_b_zip        = preg_replace('/[^0-9]/', '', addslashes($data->sheets[0]['cells'][$i][$j++]));
        $sc_od_b_zip1       = substr($od_b_zip, 0, 3);
        $sc_od_b_zip2       = substr($od_b_zip, 3);
        $sc_od_b_addr1      = clean_xss_tags(addslashes($data->sheets[0]['cells'][$i][$j++]));
        $sc_od_r            = clean_xss_tags(addslashes($data->sheets[0]['cells'][$i][$j++]));
        $sc_od_memo         = clean_xss_tags(addslashes($data->sheets[0]['cells'][$i][$j++]));
        $sc_od_company      = clean_xss_tags(addslashes($data->sheets[0]['cells'][$i][$j++]));
        $sc_od_mid          = clean_xss_tags(addslashes($data->sheets[0]['cells'][$i][$j++]));
        $sc_od_uni          = clean_xss_tags(addslashes($data->sheets[0]['cells'][$i][$j++]));
        $it_id              = clean_xss_tags(addslashes($data->sheets[0]['cells'][$i][$j++]));
        $sc_delivery_delay  = clean_xss_tags(addslashes($data->sheets[0]['cells'][$i][$j++]));
        
        // 상품정보
        /*
        $sql = " select * from {$g5['g5_shop_item_table']} where it_id = '$it_id' ";
        $it = sql_fetch($sql);
        if(!$it['it_id'])
            alert('상품정보가 존재하지 않습니다.');
        */
//장바구니에 넣기        

        // 장바구니에 Insert
        // 바로구매일 경우 장바구니가 체크된것으로 강제 설정
        if($sw_direct) {
            $ct_select = 1;
            $ct_select_time = G5_TIME_YMDHIS;
        } else {
            $ct_select = 0;
            $ct_select_time = '0000-00-00 00:00:00';
        }

        // 장바구니에 Insert
        $comma = '';
        $sql = " INSERT INTO {$g5['g5_shop_cart_table']}
                        ( od_id, mb_id, it_id, it_name, it_sc_type, it_sc_method, it_sc_price, it_sc_minimum, it_sc_qty, ct_status, ct_price, ct_point, ct_point_use, ct_stock_use, ct_option, ct_qty, ct_notax, io_id, io_type, io_price, ct_time, ct_ip, ct_send_cost, ct_direct, ct_select, ct_select_time )
                    VALUES ";
        $sql .= "( '$tmp_cart_id', '{$member['mb_id']}', '{$it_id}', '".addslashes($it['it_name'])."', '{$it['it_sc_type']}', '{$it['it_sc_method']}', '{$it['it_sc_price']}', '{$it['it_sc_minimum']}', '{$it['it_sc_qty']}', '준비', '{$it['it_price']}', '$point', '0', '0', '$io_value', '$ct_qty', '{$it['it_notax']}', '$io_id', '$io_type', '$io_price', '".G5_TIME_YMDHIS."', '$REMOTE_ADDR', '$ct_send_cost', '$sw_direct', '$ct_select', '$ct_select_time' )";
            sql_query($sql);
    }
}
    ?>