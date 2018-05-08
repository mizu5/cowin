<?php
$sub_menu = '400400';
include_once('./_common.php');

//print_r2($_GET); exit;
/*
check_admin_token();

for ($i=0; $i<count($_POST['chk']); $i++)
{
    // 실제 번호를 넘김
    $k     = $_POST['chk'][$i];
    $od_id = $_POST['od_id'][$k];

    $od = sql_fetch(" select * from {$g5['g5_shop_order_table']} where od_id = '$od_id' ");
    if (!$od) continue;

    // 주문상태가 주문이 아니면 건너뜀
    if($od['od_status'] != '주문') continue;

    $data = serialize($od);

    $sql = " insert {$g5['g5_shop_order_delete_table']} set de_key = '$od_id', de_data = '".addslashes($data)."', mb_id = '{$member['mb_id']}', de_ip = '{$_SERVER['REMOTE_ADDR']}', de_datetime = '".G5_TIME_YMDHIS."' ";
    sql_query($sql, true);

    // cart 테이블의 상품 상태를 삭제로 변경
    $sql = " update {$g5['g5_shop_cart_table']} set ct_status = '삭제' where od_id = '$od_id' and ct_status = '주문' ";
    sql_query($sql);

    $sql = " delete from {$g5['g5_shop_order_table']} where od_id = '$od_id' ";
    sql_query($sql);
}
*/

$where = array();

$doc = strip_tags($doc);

$sort1 = in_array($sort1, array('od_id', 'od_cart_price', 'od_receipt_price', 'od_cancel_price', 'od_misu', 'od_cash')) ? $sort1 : '';
$sort2 = in_array($sort2, array('desc', 'asc')) ? $sort2 : 'desc';

$sel_field = get_search_string($sel_field);
if( !in_array($sel_field, array('od_id', 'mb_id', 'od_name', 'od_tel', 'od_hp', 'od_b_name', 'od_b_tel', 'od_b_hp', 'od_deposit_name', 'od_invoice')) ){   //검색할 필드 대상이 아니면 값을 제거
    $sel_field = '';
}
echo "sel_field:".$sel_field;
$od_status = get_search_string($od_status);
$search = get_search_string($search);
if(! preg_match("/^[0-9]{4}-(0[1-9]|1[0-2])-(0[1-9]|[1-2][0-9]|3[0-1])$/", $fr_date) ) $fr_date = '';
if(! preg_match("/^[0-9]{4}-(0[1-9]|1[0-2])-(0[1-9]|[1-2][0-9]|3[0-1])$/", $to_date) ) $to_date = '';

$sql_search = "";
if ($search != "") {
    if ($sel_field != "") {
        $where[] = " $sel_field like '%$search%' ";
    }
    
    if ($save_search != $search) {
        $page = 1;
    }
}

if ($od_status) {
    switch($od_status) {
        case '전체취소':
            $where[] = " od_status = '취소' ";
            break;
        case '부분취소':
            $where[] = " od_status IN('주문', '입금', '준비', '배송', '완료') and od_cancel_price > 0 ";
            break;
        default:
            $where[] = " od_status = '$od_status' ";
            break;
    }
    
    switch ($od_status) {
        case '주문' :
            $sort1 = "od_id";
            $sort2 = "desc";
            break;
        case '입금' :   // 결제완료
            $sort1 = "od_receipt_time";
            $sort2 = "desc";
            break;
        case '배송' :   // 배송중
            $sort1 = "od_invoice_time";
            $sort2 = "desc";
            break;
    }
}

if ($od_settle_case) {
    $where[] = " od_settle_case = '$od_settle_case' ";
}

if ($od_misu) {
    $where[] = " od_misu != 0 ";
}

if ($od_cancel_price) {
    $where[] = " od_cancel_price != 0 ";
}

if ($od_refund_price) {
    $where[] = " od_refund_price != 0 ";
}

if ($od_receipt_point) {
    $where[] = " od_receipt_point != 0 ";
}

if ($od_coupon) {
    $where[] = " ( od_cart_coupon > 0 or od_coupon > 0 or od_send_coupon > 0 ) ";
}

if ($od_escrow) {
    $where[] = " od_escrow = 1 ";
}

if ($fr_date && $to_date) {
    $where[] = " od_time between '$fr_date 00:00:00' and '$to_date 23:59:59' ";
}

if ($where) {
    $sql_search = ' where '.implode(' and ', $where);
}

if ($sel_field == "")  $sel_field = "od_id";
if ($sort1 == "") $sort1 = "od_id";
if ($sort2 == "") $sort2 = "desc";

$sql_common = " from {$g5['g5_shop_order_table']} $sql_search ";

$sql  = " select *,
            (od_cart_coupon + od_coupon + od_send_coupon) as couponprice
           $sql_common
           order by $sort1 $sort2";

           $result = sql_query($sql);
           for ($i=0; $row=sql_fetch_array($result); $i++)
           {
               if (!$row) continue;
               
               // 주문상태가 주문이 아니면 건너뜀
             //  if($row['od_status'] != '주문') continue;
               
               $od_id = $row['od_id'];
               $data = serialize($row);
               $sql = " insert {$g5['g5_shop_order_delete_table']} set de_key = '$od_id', de_data = '".addslashes($data)."', mb_id = '{$member['mb_id']}', de_ip = '{$_SERVER['REMOTE_ADDR']}', de_datetime = '".G5_TIME_YMDHIS."' ";
               sql_query($sql, true);
               
               // cart 테이블의 상품 상태를 삭제로 변경
               $sql = " update {$g5['g5_shop_cart_table']} set ct_status = '삭제' where od_id = '$od_id'";
               sql_query($sql);
               
               $sql = " delete from {$g5['g5_shop_order_table']} where od_id = '$od_id' ";
               sql_query($sql);
           }
           
           
$qstr  = "sort1=$sort1&amp;sort2=$sort2&amp;sel_field=$sel_field&amp;search=$search";
$qstr .= "&amp;od_status=$search_od_status";
$qstr .= "&amp;od_settle_case=$od_settle_case";
$qstr .= "&amp;od_misu=$od_misu";
$qstr .= "&amp;od_cancel_price=$od_cancel_price";
$qstr .= "&amp;od_receipt_price=$od_receipt_price";
$qstr .= "&amp;od_receipt_point=$od_receipt_point";
$qstr .= "&amp;od_receipt_coupon=$od_receipt_coupon";

goto_url("./orderlist.php?$qstr");
?>