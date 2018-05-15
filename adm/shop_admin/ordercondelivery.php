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
if($od_wm && $od_tm){
	$where[] = " sm in ('wmp','tm') ";
}

if ($od_wm && !$od_tm){
	$where[] = " sm = 'wmp' ";
}
if ($od_tm && !$od_wm){
	$where[] = " sm = 'tm' ";
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
           error_reporting(E_ALL ^ E_NOTICE);
           
           $fail_od_id = array();
           $total_count = 0;
           $fail_count = 0;
           $succ_count = 0;
           
           
           for ($i=0; $row=sql_fetch_array($result); $i++)
           {
           	$total_count++;
               if (!$row) continue;
               // 주문정보
               $od_id = $row[od_id];
               if (!$od_id) {
               	$fail_count++;
               	$fail_od_id[] = $od_id;
               	continue;
               }
               
               if($row['od_status'] != '준비') {
               	$fail_count++;
               	$fail_od_id[] = $od_id;
               	continue;
               }
               $od_invoice = "000-00000";
               $od_delivery_company= "대한통운";
               $delivery['invoice'] = $od_invoice;
               $delivery['invoice_time'] = G5_TIME_YMDHIS;
               $delivery['delivery_company'] = $od_delivery_company;
               
               // 주문정보 업데이트
               order_update_delivery($od_id, $od['mb_id'], '배송', $delivery);
               change_status($od_id, '준비', '배송');
               
               $succ_count++;
               // 주문상태가 주문이 아니면 건너뜀
             //  if($row['od_status'] != '주문') continue;
           }
           
           
$qstr  = "sort1=$sort1&amp;sort2=$sort2&amp;sel_field=$sel_field&amp;search=$search";
$qstr .= "&amp;od_status=배송";
$qstr .= "&amp;od_settle_case=$od_settle_case";
$qstr .= "&amp;od_misu=$od_misu";
$qstr .= "&amp;od_wm=$od_wm";
$qstr .= "&amp;od_tm=$od_tm";
$qstr .= "&amp;od_cancel_price=$od_cancel_price";
$qstr .= "&amp;od_receipt_price=$od_receipt_price";
$qstr .= "&amp;od_receipt_point=$od_receipt_point";
$qstr .= "&amp;od_receipt_coupon=$od_receipt_coupon";

//goto_url("./orderlist.php?$qstr");

$g5['title'] = '엑셀 배송일괄처리 결과';
include_once(G5_PATH.'/head.sub.php');
?>

<div class="new_win">
    <h1><?php echo $g5['title']; ?></h1>

    <div class="local_desc01 local_desc">
        <p>배송일괄처리를 완료했습니다.</p>
    </div>

    <dl id="excelfile_result">
        <dt>총배송건수</dt>
        <dd><?php echo number_format($total_count); ?></dd>
        <dt class="result_done">완료건수</dt>
        <dd class="result_done"><?php echo number_format($succ_count); ?></dd>
        <dt class="result_fail">실패건수</dt>
        <dd class="result_fail"><?php echo number_format($fail_count); ?></dd>
        <?php if($fail_count > 0) { ?>
        <dt>실패주문코드</dt>
        <dd><?php echo implode(', ', $fail_od_id); ?></dd>
        <?php } ?>
    </dl>

    <div class="btn_confirm01 btn_confirm">
        <button type="button" onclick="opener.opener.location.href='./orderlist.php?<?php echo $qstr?>';opener.close();window.close();">창닫기</button>
        
    </div>

</div>

<?php
include_once(G5_PATH.'/tail.sub.php');
?>