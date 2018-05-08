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
    
    for ($i = 2; $i <= $data->sheets[0]['numRows']; $i++) {
        $total_count++;
        
        $j =1;

        
        
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
        
        $sc_od              = clean_xss_tags(addslashes($data->sheets[0]['cells'][$i][$j++]));	//소셜주문번호
        $sc_date            = clean_xss_tags(addslashes($data->sheets[0]['cells'][$i][$j++]));	//주문일시
        $sc_join_id            = clean_xss_tags(addslashes($data->sheets[0]['cells'][$i][$j++]));	//합포장번호        
       if($sc_join_id != $pre_sc_join_id){    
       	
       	// 주문서에 입력
       	$sql = " insert {$g5['g5_shop_order_table']}
       	set od_id             = '$od_id',
       	mb_id             = '{$sc_od_mid}',
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
       	od_cart_count     = '$total_qty',
       	od_cart_price     = '$total_price',
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
       	od_test           = ''
       	";
       	sql_query($sql);
       	
       	
       	
       	$od_id = get_uniqid();   
       	$total_qty = 0;
       	$total_price = 0;
       }

        $pre_sc_join_id = $sc_join_id;
        
        $sc_delivery_company= clean_xss_tags(addslashes($data->sheets[0]['cells'][$i][$j++]));	//택배사
        $sc_delivery_num    = clean_xss_tags(addslashes($data->sheets[0]['cells'][$i][$j++]));	//운송장번호
        $sc_delivery_reg    = clean_xss_tags(addslashes($data->sheets[0]['cells'][$i][$j++]));	//운송장등록
        $sc_od_name         = clean_xss_tags(addslashes($data->sheets[0]['cells'][$i][$j++]));	//구매자 이름
        $sc_item_id         = clean_xss_tags(addslashes($data->sheets[0]['cells'][$i][$j++]));	//상품ID        
        $sc_item_name       = clean_xss_tags(addslashes($data->sheets[0]['cells'][$i][$j++]));	//상품명
        $sc_item_op_id      = clean_xss_tags(addslashes($data->sheets[0]['cells'][$i][$j++]));	//옵션ID
        $sc_item_op_name    = clean_xss_tags(addslashes($data->sheets[0]['cells'][$i][$j++]));	//옵션명
        $sc_od_qty          = clean_xss_tags(addslashes($data->sheets[0]['cells'][$i][$j++]));	//구매수량
        $total_qty += (int)$sc_od_qty;
        $sc_od_price        = clean_xss_tags(addslashes($data->sheets[0]['cells'][$i][$j++]));	//구매금액
        $total_price += (int)$sc_od_price;
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

        // 장바구니에 Insert
        $comma = '';
        $sql = " INSERT INTO {$g5['g5_shop_cart_table']}
                        ( od_id,sc_od_id,sc_join_id, mb_id, it_id, it_name, it_sc_type, it_sc_method , ct_status, ct_price, ct_option, ct_qty, ct_time, ct_ip, ct_direct, ct_select_time )
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
				, '$ct_select_time' )";
            sql_query($sql);
            $succ_count++;
            //print_r2($sql);exit;
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