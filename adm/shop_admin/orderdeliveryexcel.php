<?php
$sub_menu = '400400';
include_once('./_common.php');

auth_check($auth[$sub_menu], "w");

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


if(!@sql_num_rows($result))
    alert_close('배송처리할 주문 내역이 없습니다.');

function column_char($i) { return chr( 65 + $i ); }

if (phpversion() >= '5.2.0') {
    include_once(G5_LIB_PATH.'/PHPExcel.php');
    
    $headers = array( '배송자명', '우편번호', '배송지주소', '소셜커머스', '주문자전화', '배송자전화', '', '상품명', '배송메세지','총상품수','');
    $widths  = array(18, 15, 50, 15, 15, 15, 15, 50, 30,15,15);
    $header_bgcolor = 'FFABCDEF';
    $last_char = column_char(count($headers) - 1);

    for($i=1; $row=sql_fetch_array($result); $i++) {
    	$sql = " select *
    	from {$g5['g5_shop_cart_table']}
    	where od_id = {$row['od_id']} and ct_status = '준비'";
    	$result2 = sql_query($sql);

    	$total_option ='';
    	$total_qty = 0;
    	$sm = smByName($row[sm]);
    	for($j=1; $cart_row=sql_fetch_array($result2); $j++) {
    		if($j==1){
    			$total_option .=$sm;
    		}
    		$total_option .= ' ★['.$cart_row['it_id'].'] '.$cart_row['ct_option'].' '.$cart_row['ct_qty'].'개 ';   
    		$total_qty += $cart_row['ct_qty'];
    	}    	

    	$total_op_strlen = mb_strlen($total_option,'utf-8');	//상품 옵션 문자 길이 체크
    	$d_len = 400;											//긴문자 설정
    	$total_count = ceil($total_op_strlen/$d_len);			//긴문자수로 환산한 열카운트
    	$k = 0;
    	while($k < $total_count){
    		if($total_count > 1){
    			//$option_print = substr($total_option,$k*$d_len+1,$k*$d_len+$d_len);
    			$option_print= mb_substr($total_option, $k*$d_len+1 , $k*$d_len+$d_len,'utf-8');
    		}else{
    			$option_print = $total_option;
    		}
    	if($k<1){
    	if(!empty($row['od_zip2'])){
    		$zip = $row['od_zip1'].'-'.$row['od_zip2'];
    	}else{
    		$zip = $row['od_zip1'];
    	}
    	$sm = smByName($row[sm]);
        $rows[] = 
        array($row['od_b_name'],
        					' '.$zip,
        					print_address($row['od_b_addr1'], $row['od_b_addr2'], $row['od_b_addr3'], $row['od_b_addr_jibeon']),
        					' '.$sm,
                          	' '.$row['od_hp'],
        					' '.$row['od_b_hp'],
        					'1',
        					' '.$option_print,
        					' '.$row['od_memo'],
        					' 총:  '.$total_qty.'개',
        					'신용');
    	}else{
    		$rows[] =
    		array('"','"','"','"','"','"','"',''.$option_print,'"','"','"');
    	}
                    $k++;
    	

    $data = array_merge(array($headers), $rows);
    	}
    }
    $excel = new PHPExcel();
    $excel->setActiveSheetIndex(0)->getStyle( "A1:${last_char}1" )->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setARGB($header_bgcolor);
    $excel->setActiveSheetIndex(0)->getStyle( "A:$last_char" )->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)->setWrapText(true);
    foreach($widths as $i => $w) $excel->setActiveSheetIndex(0)->getColumnDimension( column_char($i) )->setWidth($w);
    $excel->getActiveSheet()->fromArray($data,NULL,'A1');

    header("Content-Type: application/octet-stream");
    header("Content-Disposition: attachment; filename=\"deliverylist-".date("ymd", time()).".xls\"");
    header("Cache-Control: max-age=0");

    $writer = PHPExcel_IOFactory::createWriter($excel, 'Excel5');
    $writer->save('php://output');
} else {
    /*================================================================================
    php_writeexcel http://www.bettina-attack.de/jonny/view.php/projects/php_writeexcel/
    =================================================================================*/

    include_once(G5_LIB_PATH.'/Excel/php_writeexcel/class.writeexcel_workbook.inc.php');
    include_once(G5_LIB_PATH.'/Excel/php_writeexcel/class.writeexcel_worksheet.inc.php');

    $fname = tempnam(G5_DATA_PATH, "tmp-deliverylist.xls");
    $workbook = new writeexcel_workbook($fname);
    $worksheet = $workbook->addworksheet();

    // Put Excel data
    $data = array('주문번호', '주문자명', '주문자전화1', '주문자전화2', '배송자명', '배송지전화1', '배송지전화2', '배송지주소', '배송회사', '운송장번호');
    $data = array_map('iconv_euckr', $data);

    $col = 0;
    foreach($data as $cell) {
        $worksheet->write(0, $col++, $cell);
    }

    for($i=1; $row=sql_fetch_array($result); $i++) {
        $row = array_map('iconv_euckr', $row);

        $worksheet->write($i, 0, ' '.$row['od_id']);
        $worksheet->write($i, 1, $row['od_name']);
        $worksheet->write($i, 2, ' '.$row['od_tel']);
        $worksheet->write($i, 3, ' '.$row['od_hp']);
        $worksheet->write($i, 4, $row['od_b_name']);
        $worksheet->write($i, 5, ' '.$row['od_b_tel']);
        $worksheet->write($i, 6, ' '.$row['od_b_hp']);
        $worksheet->write($i, 7, print_address($row['od_b_addr1'], $row['od_b_addr2'], $row['od_b_addr3'], $row['od_b_addr_jibeon']));
        $worksheet->write($i, 8, $row['od_delivery_company']);
        $worksheet->write($i, 9, $row['od_invoice']);
    }

    $workbook->close();

    header("Content-Type: application/x-msexcel; name=\"deliverylist-".date("ymd", time()).".xls\"");
    header("Content-Disposition: inline; filename=\"deliverylist-".date("ymd", time()).".xls\"");
    $fh=fopen($fname, "rb");
    fpassthru($fh);
    unlink($fname);
}
?>