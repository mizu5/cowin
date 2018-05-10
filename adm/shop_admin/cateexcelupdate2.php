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
    $sc_join_id = 0;	//합포장 번호
    $pre_sc_join_id = 0;	//합포자 전 카트번호 
    $old_od = false;
    $row_num_one = max(count($data->sheets[0]['cells']), 0);
    $cols_num_one = $data->sheets[0]['numCols'];
    
    for ($i = 1; $i <= $row_num_one; $i++) {  	
             
        $j =1;         
        unset($e_meta);
        unset($e_val);
        for ($j = 1; $j <= $cols_num_one; $j++) {
        if($i == 1) {
        	if($data->sheets[0]['cells'][$i][$j]) {
        		$k++;
        		$e_key[$j] = $data->sheets[0]['cells'][$i][$j];
        	}
        } else {
        	$e_meta[$e_key[$j]] = $data->sheets[0]['cells'][$i][$j];
        	$e_val[$j] = $data->sheets[0]['cells'][$i][$j];
        }
  	  }
        if($i > 1){
        	$total_count++;  
        if($sm=="wmp"){
        	$sc_od              = clean_xss_tags(addslashes($e_meta['주문번호']));	//소셜주문번호
        	$sc_date            = clean_xss_tags(addslashes($e_meta['주문일시']));	//주문일시
        	$sc_join_id         = clean_xss_tags(addslashes($e_meta['합포장번호']));	//합포장번호
        	$sc_delivery_company= clean_xss_tags(addslashes($e_meta['택배사']));	//택배사
        	$sc_delivery_num    = clean_xss_tags(addslashes($e_meta['운송장번호']));	//운송장번호
        	$sc_delivery_reg    = clean_xss_tags(addslashes($e_meta['운송장등록일시']));	//운송장등록
        	$sc_od_name         = clean_xss_tags(addslashes($e_meta['구매자이름']));	//구매자 이름
        	$sc_item_id         = clean_xss_tags(addslashes($e_meta['상품ID']));	//상품ID
        	$sc_item_name       = clean_xss_tags(addslashes($e_meta['상품명']));	//상품명
        	$sc_item_op_id      = clean_xss_tags(addslashes($e_meta['옵션ID']));	//옵션ID
        	$sc_item_op_name    = clean_xss_tags(addslashes($e_meta['옵션명']));	//옵션명
        	$sc_od_qty          = (int)clean_xss_tags(addslashes($e_meta['구매수량']));	//구매수량
        	$sc_od_price        = (int)clean_xss_tags(addslashes($e_meta['구매금액']));	//구매금액
        	$sc_od_price 		= $sc_od_price / $sc_od_qty;							//위메프 구매금액을 단가로 전환 
        	$sc_od_b_name       = clean_xss_tags(addslashes($e_meta['받는분 이름']));	//받는분 이름
        	$sc_od_b_hp         = clean_xss_tags(addslashes($e_meta['받는분 휴대폰']));	//받는분 휴대폰
        	$sc_od_b_zip        = preg_replace('/[^0-9-]/', '', addslashes($e_meta['우편번호']));//받는분우편번호
        	$zip =explode('-' , $sc_od_b_zip);
        	$sc_od_b_zip1       = $zip[0];
        	if(!is_null($zip[1])){
        	$sc_od_b_zip2       = $zip[1];
        	}else{
        		unset($sc_od_b_zip2);
        	}
        	$sc_od_b_addr1      = clean_xss_tags(addslashes($e_meta['주소']));	//받는분 주소
        	$sc_od_r            = clean_xss_tags(addslashes($e_meta['배송비 유형']));	//배송비유형
        	if($sc_od_r=="무료"){
        		$it_sc_type = 1;	//무료배송
        	}else{
        		$it_sc_type = 3;	//유료배송
        		$it_sc_method = 0;
        	}
        	$sc_od_memo         = clean_xss_tags(addslashes($e_meta['배송메시지']));	//배송메세지
        	$sc_od_company      = clean_xss_tags(addslashes($e_meta['업체명']));	//업체명
        	$sc_od_mid          = clean_xss_tags(addslashes($e_meta['구매자MID']));	//구매자MID
        	$sc_od_uni          = clean_xss_tags(addslashes($e_meta['특이사항']));	//특이사항
        	$it_id              = clean_xss_tags(addslashes($e_meta['업체상품코드']));	//업체상품코드
        	$sc_delivery_delay  = clean_xss_tags(addslashes($e_meta['배송지연기준일']));	//배송지연기준일
/*        	
        $sc_od              = clean_xss_tags(addslashes($data->sheets[0]['cells'][$i][$j++]));	//소셜주문번호
        $sc_date            = clean_xss_tags(addslashes($data->sheets[0]['cells'][$i][$j++]));	//주문일시
        $sc_join_id         = clean_xss_tags(addslashes($data->sheets[0]['cells'][$i][$j++]));	//합포장번호  
        $sc_delivery_company= clean_xss_tags(addslashes($data->sheets[0]['cells'][$i][$j++]));	//택배사
        $sc_delivery_num    = clean_xss_tags(addslashes($data->sheets[0]['cells'][$i][$j++]));	//운송장번호
        $sc_delivery_reg    = clean_xss_tags(addslashes($data->sheets[0]['cells'][$i][$j++]));	//운송장등록
        $sc_od_name         = clean_xss_tags(addslashes($data->sheets[0]['cells'][$i][$j++]));	//구매자 이름
        $sc_item_id         = clean_xss_tags(addslashes($data->sheets[0]['cells'][$i][$j++]));	//상품ID
        $sc_item_name       = clean_xss_tags(addslashes($data->sheets[0]['cells'][$i][$j++]));	//상품명
        $sc_item_op_id      = clean_xss_tags(addslashes($data->sheets[0]['cells'][$i][$j++]));	//옵션ID
        $sc_item_op_name    = clean_xss_tags(addslashes($data->sheets[0]['cells'][$i][$j++]));	//옵션명
        $sc_od_qty          = (int)clean_xss_tags(addslashes($data->sheets[0]['cells'][$i][$j++]));	//구매수량
        $sc_od_price        = (int)clean_xss_tags(addslashes($data->sheets[0]['cells'][$i][$j++]));	//구매금액
        $sc_od_b_name       = clean_xss_tags(addslashes($data->sheets[0]['cells'][$i][$j++]));	//받는분 이름
        $sc_od_b_hp         = clean_xss_tags(addslashes($data->sheets[0]['cells'][$i][$j++]));	//받는분 휴대폰
        $sc_od_b_zip        = preg_replace('/[^0-9]/', '', addslashes($data->sheets[0]['cells'][$i][$j++]));//받는분우편번호
        $sc_od_b_zip1       = substr($od_b_zip, 0, 3);
        $sc_od_b_zip2       = substr($od_b_zip, 3);
        $sc_od_b_addr1      = clean_xss_tags(addslashes($data->sheets[0]['cells'][$i][$j++]));	//받는분 주소
        $sc_od_r            = clean_xss_tags(addslashes($data->sheets[0]['cells'][$i][$j++]));	//배송비유형
        if($sc_od_r=="무료"){
        	$it_sc_type = 1;	//무료배송
        }else{
        	$it_sc_type = 3;	//유료배송
        	$it_sc_method = 0;
        }
        $sc_od_memo         = clean_xss_tags(addslashes($data->sheets[0]['cells'][$i][$j++]));	//배송메세지
        $sc_od_company      = clean_xss_tags(addslashes($data->sheets[0]['cells'][$i][$j++]));	//업체명
        $sc_od_mid          = clean_xss_tags(addslashes($data->sheets[0]['cells'][$i][$j++]));	//구매자MID
        $sc_od_uni          = clean_xss_tags(addslashes($data->sheets[0]['cells'][$i][$j++]));	//특이사항
        $it_id              = clean_xss_tags(addslashes($data->sheets[0]['cells'][$i][$j++]));	//업체상품코드
        $sc_delivery_delay  = clean_xss_tags(addslashes($data->sheets[0]['cells'][$i][$j++]));	//배송지연기준일
*/        
        }else if($sm=='tm'){
        	$sc_od              = clean_xss_tags(addslashes($e_meta['주문번호']));	//소셜주문번호
        	$sc_date            = clean_xss_tags(addslashes($e_meta['최종처리일']));	//주문일시
        	$sc_join_id         = clean_xss_tags(addslashes($e_meta['딜번호']));	//합포장번호
        	$sc_delivery_company= clean_xss_tags(addslashes($e_meta['택배사']));	//택배사
        	$sc_delivery_num    = clean_xss_tags(addslashes($e_meta['운송장번호']));	//운송장번호
        	$sc_delivery_reg    = clean_xss_tags(addslashes($e_meta['운송장등록시점']));	//운송장등록
        	$sc_od_name         = clean_xss_tags(addslashes($e_meta['주문자명']));	//구매자 이름
        	$sc_od_id         = clean_xss_tags(addslashes($e_meta['아이디']));	//구매자 ID(티몬만 있음)
        	$sc_item_id         = clean_xss_tags(addslashes($e_meta['딜번호']));	//상품ID
        	$sc_item_name       = clean_xss_tags(addslashes($e_meta['상품명']));	//상품명
        	$sc_item_op_id      = clean_xss_tags(addslashes($e_meta['옵션번호']));	//옵션ID
        	$sc_item_op_name    = clean_xss_tags(addslashes($e_meta['옵션명']));	//옵션명
        	$sc_od_qty          = (int)clean_xss_tags(addslashes($e_meta['구매수량']));	//구매수량
        	$sc_od_price        = (int)clean_xss_tags(addslashes($e_meta['판매단가']));	//판매단가
        	$sc_od_b_name       = clean_xss_tags(addslashes($e_meta['수취인명']));	//받는분 이름
        	$sc_od_b_hp         = clean_xss_tags(addslashes($e_meta['수취인연락처']));	//받는분 휴대폰
        	$sc_od_b_zip        = preg_replace('/[^0-9-]/', '', addslashes($e_meta['수취인우편번호']));//받는분우편번호
        	$zip =explode('-' , $sc_od_b_zip);
        	$sc_od_b_zip1       = $zip[0];
        	if(!is_null($zip[1])){
        		$sc_od_b_zip2       = $zip[1];
        	}else{
        		unset($sc_od_b_zip2);
        	}
        	$sc_od_b_addr1      = clean_xss_tags(addslashes($e_meta['수취인주소']));	//받는분 주소
        	//$sc_od_r            = clean_xss_tags(addslashes($e_meta['배송비 유형']));	//배송비유형
        	$sc_od_r            = "무료";	//배송비유형
        	if($sc_od_r=="무료"){
        		$it_sc_type = 1;	//무료배송
        	}else{
        		$it_sc_type = 3;	//유료배송
        		$it_sc_method = 0;
        	}
        	$sc_od_memo         = clean_xss_tags(addslashes($e_meta['배송요청메모']));	//배송메세지
        	$sc_od_company      = clean_xss_tags(addslashes($e_meta['업체명']));	//업체명
        	$sc_od_mid          = clean_xss_tags(addslashes($e_meta['아이디']));	//구매자MID
        	$sc_od_uni          = clean_xss_tags(addslashes($e_meta['특이사항']));	//특이사항
        	$it_id              = clean_xss_tags(addslashes($e_meta['파트너 CODE1']));	//업체상품코드
        	$sc_delivery_delay  = clean_xss_tags(addslashes($e_meta['지연신고일']));	//배송지연기준일
/*        	
        	$sc_dill 			= clean_xss_tags(addslashes($data->sheets[0]['cells'][$i][$j++]));	//딜번호
        	$sc_od              = clean_xss_tags(addslashes($data->sheets[0]['cells'][$i][$j++]));	//소셜주문번호
        	$sc_date            = clean_xss_tags(addslashes($data->sheets[0]['cells'][$i][$j++]));	//주문일시
        	$sc_join_id         = clean_xss_tags(addslashes($data->sheets[0]['cells'][$i][$j++]));	//합포장번호
        	$sc_delivery_company= clean_xss_tags(addslashes($data->sheets[0]['cells'][$i][$j++]));	//택배사
        	$sc_delivery_num    = clean_xss_tags(addslashes($data->sheets[0]['cells'][$i][$j++]));	//운송장번호
        	$sc_delivery_reg    = clean_xss_tags(addslashes($data->sheets[0]['cells'][$i][$j++]));	//운송장등록
        	$sc_od_name         = clean_xss_tags(addslashes($data->sheets[0]['cells'][$i][$j++]));	//구매자 이름
        	$sc_item_id         = clean_xss_tags(addslashes($data->sheets[0]['cells'][$i][$j++]));	//상품ID
        	$sc_item_name       = clean_xss_tags(addslashes($data->sheets[0]['cells'][$i][$j++]));	//상품명
        	$sc_item_op_id      = clean_xss_tags(addslashes($data->sheets[0]['cells'][$i][$j++]));	//옵션ID
        	$sc_item_op_name    = clean_xss_tags(addslashes($data->sheets[0]['cells'][$i][$j++]));	//옵션명
        	$sc_od_qty          = (int)clean_xss_tags(addslashes($data->sheets[0]['cells'][$i][$j++]));	//구매수량
        	$sc_od_price        = (int)clean_xss_tags(addslashes($data->sheets[0]['cells'][$i][$j++]));	//구매금액
        	$sc_od_b_name       = clean_xss_tags(addslashes($data->sheets[0]['cells'][$i][$j++]));	//받는분 이름
        	$sc_od_b_hp         = clean_xss_tags(addslashes($data->sheets[0]['cells'][$i][$j++]));	//받는분 휴대폰
        	$sc_od_b_zip        = preg_replace('/[^0-9]/', '', addslashes($data->sheets[0]['cells'][$i][$j++]));//받는분우편번호
        	$sc_od_b_zip1       = substr($od_b_zip, 0, 3);
        	$sc_od_b_zip2       = substr($od_b_zip, 3);
        	$sc_od_b_addr1      = clean_xss_tags(addslashes($data->sheets[0]['cells'][$i][$j++]));	//받는분 주소
        	$sc_od_r            = clean_xss_tags(addslashes($data->sheets[0]['cells'][$i][$j++]));	//배송비유형
        	if($sc_od_r=="무료"){
        		$it_sc_type = 1;	//무료배송
        	}else{
        		$it_sc_type = 3;	//유료배송
        		$it_sc_method = 0;
        	}
        	$sc_od_memo         = clean_xss_tags(addslashes($data->sheets[0]['cells'][$i][$j++]));	//배송메세지
        	$sc_od_company      = clean_xss_tags(addslashes($data->sheets[0]['cells'][$i][$j++]));	//업체명
        	$sc_od_mid          = clean_xss_tags(addslashes($data->sheets[0]['cells'][$i][$j++]));	//구매자MID
        	$sc_od_uni          = clean_xss_tags(addslashes($data->sheets[0]['cells'][$i][$j++]));	//특이사항
        	$it_id              = clean_xss_tags(addslashes($data->sheets[0]['cells'][$i][$j++]));	//업체상품코드
        	$sc_delivery_delay  = clean_xss_tags(addslashes($data->sheets[0]['cells'][$i][$j++]));	//배송지연기준일
*/        	
        }
        
        
        $sql = "select od_id,od_cart_count,od_cart_price from {$g5['g5_shop_order_table']} where sc_od_id = '{$sc_od}' and sc_join_id = '{$sc_join_id}'";
        /*$sql = "select SO.od_id AS od_id,SO.od_cart_count AS od_cart_count,SO.od_cart_price AS od_cart_price,SC.sc_od_id AS sc_od_id, SC.it_name AS it_name, SC.ct_option AS ct_option from {$g5['g5_shop_order_table']} 
AS SO RIGHT OUTER JOIN  {$g5['g5_shop_cart_table']} AS SC ON SC.od_id = SO.od_id where SC.sc_od_id = '{$sc_od}' and SC.sc_join_id = '{$sc_join_id}'and SC.ct_option NOT IN ('{$sc_item_op_name}')";
*/
//        echo $sql;        exit;
        $row = sql_fetch($sql); 
        echo "ccount:".$row[od_cart_count];
        echo "cprice:".$row[od_cart_price];

        if(!$row[od_id]){
        	$od_id = get_uniqid();
        	// 주문서에 입력
        	$sql = " insert {$g5['g5_shop_order_table']}
        	set od_id             = '$od_id',
        	mb_id             = '{$sc_od_mid}',
        	sc_od_id		  =  '{$sc_od}',
        	sc_join_id		  =  '{$sc_join_id}',
        	od_pwd            = '1111',
        	od_name           = '$sc_od_name',
        	od_email          = '',
        	od_tel            = '',
        	od_hp             = '$sc_od_b_hp',
        	od_zip1           = '$sc_od_b_zip1',
        	od_zip2           = '$sc_od_b_zip2',
        	od_addr1          = '$sc_od_b_addr1',
        	od_addr2          = '',
        	od_addr3          = '',
        	od_addr_jibeon    = '',
        	od_b_name         = '$sc_od_b_name',
        	od_b_tel          = '',
        	od_b_hp           = '$sc_od_b_hp',
        	od_b_zip1         = '$sc_od_b_zip1',
        	od_b_zip2         = '$sc_od_b_zip2',
        	od_b_addr1        = '$sc_od_b_addr1',
        	od_b_addr2        = '',
        	od_b_addr3        = '',
        	od_b_addr_jibeon  = '',
        	od_deposit_name   = '',
        	od_memo           = '$sc_od_memo',
        	od_cart_count     = '$sc_od_qty',
        	od_cart_price     = '$sc_od_price',
        	od_cart_coupon    = '',
        	od_send_cost      = '',
        	od_send_coupon    = '',
        	od_send_cost2     = '',
        	od_coupon         = '',
        	od_receipt_price  = '$od_receipt_price',
        	od_receipt_point  = '$od_receipt_point',
        	od_bank_account   = '$od_bank_account',
        	od_receipt_time   = '$od_receipt_time',
        	od_misu           = '0',
        	od_pg             = '$od_pg',
        	od_tno            = '$od_tno',
        	od_app_no         = '$od_app_no',
        	od_escrow         = '',
        	od_tax_flag       = '',
        	od_tax_mny        = '',
        	od_vat_mny        = '',
        	od_free_mny       = '',
        	od_status         = '준비',
        	od_shop_memo      = '',
        	od_hope_date      = '$od_hope_date',
        	od_time           = '".G5_TIME_YMDHIS."',
        	od_ip             = '$REMOTE_ADDR',
        	od_settle_case    = '무통장',
        	od_test           = '',
			sm 				= '$sm'	
        	";
        	sql_query($sql);
        	//$total_qty = 0;
        	//$total_price = 0;        	
        }else{
        	$od_id = $row[od_id];
        	$old_od = true;
        }
        
        // 상품정보
        /*
        $sql = " select * from {$g5['g5_shop_item_table']} where it_id = '$it_id' ";
        $it = sql_fetch($sql);
        if(!$it['it_id'])
            alert('상품정보가 존재하지 않습니다.');
        */
   

        // 장바구니에 Insert
        // 바로구매일 경우 장바구니가 체크된것으로 강제 설정
        if($sw_direct) {
            $ct_select = 1;
            $ct_select_time = G5_TIME_YMDHIS;
        } else {
            $ct_select = 0;
            $ct_select_time = '0000-00-00 00:00:00';
        }
		$sql = "SELECT  od_id from {$g5['g5_shop_cart_table']} where sc_od_id = '{$sc_od}' and ct_option IN ('{$sc_item_op_name}')";	//카트에 중복 등록 방지
		//echo $sql; exit;
		$row_du = sql_fetch($sql);
        // 장바구니에 Insert
        $comma = '';
        if(!$row_du[od_id]){
        $sql = " INSERT INTO {$g5['g5_shop_cart_table']}
                        ( od_id,sc_od_id,sc_join_id, mb_id, it_id, it_name, it_sc_type, it_sc_method , ct_status, ct_price, ct_option, ct_qty, ct_time, ct_ip, ct_direct, ct_select_time, sm )
                    VALUES ";
        $sql .= "( '{$od_id}'
				, '{$sc_od}'
				, '{$sc_join_id}'
        		, '{$sc_od_mid}'						
				, '{$it_id}'							
				, '".addslashes($sc_item_name)."'		
				, '{$it_sc_type}'									
				, '{$it_sc_method}'
				, '준비'
				, '{$sc_od_price}'		
				, '{$sc_item_op_name}'				
				, '$sc_od_qty'									
				, '".G5_TIME_YMDHIS."'
				, '$REMOTE_ADDR'
				, '0'	
				, '$ct_select_time' 
				, '{$sm}')";
        //echo $sql; exit;
            sql_query($sql);
            if($old_od){
           	$sc_qty = 0;
           	$sc_price = 0;
            $sc_qty = (int)$row[od_cart_count]+$sc_od_qty;
            $sc_price = (int)$row[od_cart_price]+$sc_od_price;
            $sql  = "update {$g5['g5_shop_order_table']}
            set od_cart_count             = '$sc_qty',
            od_cart_price           	  =	'$sc_price'
			where od_id = {$od_id}";
//            echo $sql;
//            exit; 
            sql_query($sql);
            }
            $succ_count++;
            //print_r2($sql);exit;
        }//if(!$row_du[od_id]){
        } 
    }
}

$g5['title'] = '주문 엑셀일괄등록 결과';
include_once(G5_PATH.'/head.sub.php');
?>

<div class="new_win">
    <h1><?php echo $g5['title']; ?></h1>

    <div class="local_desc01 local_desc">
        <p>위메프 주문 등록을 완료했습니다.</p>
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
        <button type="button" onclick="opener.location.reload();window.close();">창닫기</button>
    </div>

</div>

<?php
include_once(G5_PATH.'/tail.sub.php');
?>