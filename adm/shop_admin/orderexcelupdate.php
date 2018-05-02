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

        $j =3;

        // 새로운 주문번호 생성
        $od_id = get_uniqid();
        $od_name          = clean_xss_tags(addslashes($data->sheets[0]['cells'][$i][$j++]));
        $od_email         = get_email_address(addslashes($data->sheets[0]['cells'][$i][$j++]));
        $od_tel           = clean_xss_tags(addslashes($data->sheets[0]['cells'][$i][$j++]));
        $od_hp            = clean_xss_tags(addslashes($data->sheets[0]['cells'][$i][$j++]));
        $od_zip           = preg_replace('/[^0-9]/', '', addslashes($data->sheets[0]['cells'][$i][$j++]));
        $od_zip1          = substr($od_zip, 0, 3);
        $od_zip2          = substr($od_zip, 3);
        $od_addr1         = clean_xss_tags(addslashes($data->sheets[0]['cells'][$i][$j++]));
        $od_addr2         = clean_xss_tags(addslashes($data->sheets[0]['cells'][$i][$j++]));
        $od_addr3         = clean_xss_tags(addslashes($data->sheets[0]['cells'][$i][$j++]));
        $od_addr_jibeon   =  addslashes($data->sheets[0]['cells'][$i][$j++]);
        $od_addr_jibeon   = preg_match("/^(N|R)$/", $od_addr_jibeon) ? $od_addr_jibeon : '';
        $od_deposit_name  = clean_xss_tags(addslashes($data->sheets[0]['cells'][$i][$j++]));
        $od_b_name        = clean_xss_tags(addslashes($data->sheets[0]['cells'][$i][$j++]));
        $od_b_tel         = clean_xss_tags(addslashes($data->sheets[0]['cells'][$i][$j++]));
        $od_b_hp          = clean_xss_tags(addslashes($data->sheets[0]['cells'][$i][$j++]));
        $od_b_zip           = preg_replace('/[^0-9]/', '', addslashes($data->sheets[0]['cells'][$i][$j++]));
        $od_b_zip1          = substr($od_b_zip, 0, 3);
        $od_b_zip2          = substr($od_b_zip, 3);
        $od_b_addr1       = clean_xss_tags(addslashes($data->sheets[0]['cells'][$i][$j++]));
        $od_b_addr2       = clean_xss_tags(addslashes($data->sheets[0]['cells'][$i][$j++]));
        $od_b_addr3       = clean_xss_tags(addslashes($data->sheets[0]['cells'][$i][$j++]));
        $od_b_addr_jibeon =  addslashes($data->sheets[0]['cells'][$i][$j++]);
        $od_b_addr_jibeon = preg_match("/^(N|R)$/", $od_b_addr_jibeon) ? $od_b_addr_jibeon : '';
        $od_memo          = clean_xss_tags(addslashes($data->sheets[0]['cells'][$i][$j++]));
        $cart_count       = clean_xss_tags(addslashes($data->sheets[0]['cells'][$i][$j++]));
        $tot_ct_price     = clean_xss_tags(addslashes($data->sheets[0]['cells'][$i][$j++]));
        $od_tax_flag      = $default['de_tax_flag_use'];
        
/*        
        if(!$it_id || !$ca_id || !$it_name) {
            $fail_count++;
            continue;
        }
/*
        // it_id 중복체크
        $sql2 = " select count(*) as cnt from {$g5['g5_shop_item_table']} where it_id = '$it_id' ";
        $row2 = sql_fetch($sql2);
        if($row2['cnt']) {
            $fail_it_id[] = $it_id;
            $dup_it_id[] = $it_id;
            $dup_count++;
            $fail_count++;
            continue;
        }

        // 기본분류체크
        $sql2 = " select count(*) as cnt from {$g5['g5_shop_category_table']} where ca_id = '$ca_id' ";
        $row2 = sql_fetch($sql2);
        if(!$row2['cnt']) {
            $fail_it_id[] = $it_id;
            $fail_count++;
            continue;
        }
*/
        


        
        // 주문서에 입력
        $sql = " insert {$g5['g5_shop_order_table']}
            set od_id             = '$od_id',
                mb_id             = '{$member['mb_id']}',
                od_pwd            = '$od_pwd',
                od_name           = '$od_name',
                od_email          = '$od_email',
                od_tel            = '$od_tel',
                od_hp             = '$od_hp',
                od_zip1           = '$od_zip1',
                od_zip2           = '$od_zip2',
                od_addr1          = '$od_addr1',
                od_addr2          = '$od_addr2',
                od_addr3          = '$od_addr3',
                od_addr_jibeon    = '$od_addr_jibeon',
                od_b_name         = '$od_b_name',
                od_b_tel          = '$od_b_tel',
                od_b_hp           = '$od_b_hp',
                od_b_zip1         = '$od_b_zip1',
                od_b_zip2         = '$od_b_zip2',
                od_b_addr1        = '$od_b_addr1',
                od_b_addr2        = '$od_b_addr2',
                od_b_addr3        = '$od_b_addr3',
                od_b_addr_jibeon  = '$od_b_addr_jibeon',
                od_deposit_name   = '$od_deposit_name',
                od_memo           = '$od_memo',
                od_cart_count     = '$cart_count',
                od_cart_price     = '$tot_ct_price',
                od_cart_coupon    = '$tot_it_cp_price',
                od_send_cost      = '$od_send_cost',
                od_send_coupon    = '$tot_sc_cp_price',
                od_send_cost2     = '$od_send_cost2',
                od_coupon         = '$tot_od_cp_price',
                od_receipt_price  = '$od_receipt_price',
                od_receipt_point  = '$od_receipt_point',
                od_bank_account   = '$od_bank_account',
                od_receipt_time   = '$od_receipt_time',
                od_misu           = '$od_misu',
                od_pg             = '$od_pg',
                od_tno            = '$od_tno',
                od_app_no         = '$od_app_no',
                od_escrow         = '$od_escrow',
                od_tax_flag       = '$od_tax_flag',
                od_tax_mny        = '$od_tax_mny',
                od_vat_mny        = '$od_vat_mny',
                od_free_mny       = '$od_free_mny',
                od_status         = '주문',
                od_shop_memo      = '',
                od_hope_date      = '$od_hope_date',
                od_time           = '".G5_TIME_YMDHIS."',
                od_ip             = '$REMOTE_ADDR',
                od_settle_case    = '$od_settle_case',
                od_test           = '{$default['de_card_test']}'
                ";
        echo $sql;    
        sql_query($sql);
        $succ_count++;
    }
}

$g5['title'] = '상품 엑셀일괄등록 결과';
include_once(G5_PATH.'/head.sub.php');
?>

<div class="new_win">
    <h1><?php echo $g5['title']; ?></h1>

    <div class="local_desc01 local_desc">
        <p>상품등록을 완료했습니다.</p>
    </div>

    <dl id="excelfile_result">
        <dt>총상품수</dt>
        <dd><?php echo number_format($total_count); ?></dd>
        <dt>완료건수</dt>
        <dd><?php echo number_format($succ_count); ?></dd>
        <dt>실패건수</dt>
        <dd><?php echo number_format($fail_count); ?></dd>
        <?php if($fail_count > 0) { ?>
        <dt>실패상품코드</dt>
        <dd><?php echo implode(', ', $fail_it_id); ?></dd>
        <?php } ?>
        <?php if($dup_count > 0) { ?>
        <dt>상품코드중복건수</dt>
        <dd><?php echo number_format($dup_count); ?></dd>
        <dt>중복상품코드</dt>
        <dd><?php echo implode(', ', $dup_it_id); ?></dd>
        <?php } ?>
    </dl>

    <div class="btn_win01 btn_win">
        <button type="button" onclick="window.close();">창닫기</button>
    </div>

</div>

<?php
include_once(G5_PATH.'/tail.sub.php');
?>