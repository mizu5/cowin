<?php
$sub_menu = '400400';
include_once('./_common.php');

$where = array();

$doc = strip_tags($doc);

$sort1 = in_array($sort1, array('od_id', 'od_cart_price', 'od_receipt_price', 'od_cancel_price', 'od_misu', 'od_cash')) ? $sort1 : '';
$sort2 = in_array($sort2, array('desc', 'asc')) ? $sort2 : 'desc';

$sel_field = get_search_string($sel_field);
if( !in_array($sel_field, array('od_id', 'mb_id', 'od_name', 'od_tel', 'od_hp', 'od_b_name', 'od_b_tel', 'od_b_hp', 'od_deposit_name', 'od_invoice')) ){   //검색할 필드 대상이 아니면 값을 제거
	$sel_field = '';
}

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



$qstr  = "sort1=$sort1&amp;sort2=$sort2&amp;sel_field=$sel_field&amp;search=$search";
$qstr .= "&amp;od_status=".urlencode($od_status);
$qstr .= "&amp;od_settle_case=$od_settle_case";
$qstr .= "&amp;od_misu=$od_misu";
$qstr .= "&amp;od_wm=$od_wm";
$qstr .= "&amp;od_tm=$od_tm";
$qstr .= "&amp;od_cancel_price=$od_cancel_price";
$qstr .= "&amp;od_receipt_price=$od_receipt_price";
$qstr .= "&amp;od_receipt_point=$od_receipt_point";
$qstr .= "&amp;od_receipt_coupon=$od_receipt_coupon";

auth_check($auth[$sub_menu], "w");

$g5['title'] = '엑셀 배송일괄처리';
include_once(G5_PATH.'/head.sub.php');
?>
<script>
	$(document).ready(function(){
		$('.order_con_wrap').hide();

	});
	function showExeDelivery(){
		$('.order_con_wrap').fadeIn();
	}	
</script>
<div class="new_win">
    <h1><?php echo $g5['title']; ?></h1>

    <div class="local_desc01 local_desc">
        <p>
            엑셀파일을 이용하여 배송정보를 일괄등록할 수 있습니다.<br>
            형식은 <strong>배송처리용 엑셀파일</strong>을 다운로드하여 배송 정보를 입력하시면 됩니다.<br>
            수정 완료 후 엑셀파일을 업로드하시면 배송정보가 일괄등록됩니다.<br>
            엑셀파일을 저장하실 때는 <strong>Excel 97 - 2003 통합문서 (*.xls)</strong> 로 저장하셔야 합니다.<br>
            주문상태가 준비이고 미수금이 0인 주문에 한해 엑셀파일이 생성됩니다.
        </p>

        <p  class="local_ov01 local_ov">
            <a href="<?php echo G5_ADMIN_URL;?>/shop_admin/orderdeliveryexcel.php?<?php echo $qstr; ?>" onclick="showExeDelivery();"  class="ov_a">택배 일괄등록용 엑셀파일 다운로드</a>


             <a href="./ordercondelivery.php?<?php echo $qstr;?>" id="order_con_delivery" class="ov_a  order_con_wrap">주문내역 일괄 배송처리</a>
        </p>     
    </div>

    <form name="forderdelivery" method="post" action="./orderdeliveryupdate.php" enctype="MULTIPART/FORM-DATA" autocomplete="off">

    <div id="excelfile_upload">
        <label for="excelfile">파일선택</label>
        <input type="file" name="excelfile" id="excelfile">
    </div>

    <div id="excelfile_input">
        <input type="checkbox" name="od_send_mail" value="1" id="od_send_mail" checked="checked">
        <label for="od_send_mail">배송안내 메일</label>
        <input type="checkbox" name="send_sms" value="1" id="od_send_sms" checked="checked">
        <label for="od_send_sms">배송안내 SMS</label>
        <input type="checkbox" name="send_escrow" value="1" id="od_send_escrow">
        <label for="od_send_escrow">에스크로배송등록</label>
    </div>

    <div class="btn_confirm01 btn_confirm" style="text-align:center">
        <input type="submit" value="배송정보 등록" class="btn_submit">
        <button type="button" onclick="window.close();">닫기</button>
    </div>

    </form>

</div>
<script>
$(function(){
	// 위메프 엑셀배송처리창
	$("#order_con_delivery").on("click", function() {
	    var opt = "width=600,height=450,left=10,top=10";
	    var input = confirm('검색하신 주문내역의 상품을 배송처리합니다.');
	    if(input){window.open(this.href, "win_del", opt);}
	    return false;
	});    
});
</script>
<?php
include_once(G5_PATH.'/tail.sub.php');
?>