<?php
/*
echo time("2018-04-24  11:56:00 PM");
echo date("Y-m-d",time("2018-04-24  11:56:00 PM"));
*/
$total_option= "착한가격★생활용품 146종- 046. 실리콘리모컨커버기|02. L 1개 | 080. 토끼선정리기2set|랜덤 1개 | 091. 세면대머리카락필터|01. 소형 24개 3개 |";
$total_op_strlen = mb_strlen($total_option,'utf-8');
echo $total_op_strlen;
$total_count = ceil($total_op_strlen/50);
$k = 0;
while($k < $total_count){
	$str = mb_substr($total_option, $k*50+1 , $k*50+50,'utf-8');
	echo $str."<br>";
	$k++;
}

$sc_od_b_zip = "502-859";
$sc_od_b_zip        = preg_replace('/[^0-9-]/', '',$sc_od_b_zip);//받는분우편번호
$zip =explode('-' , $sc_od_b_zip);
echo $zip[0].'-'.$zip[1];

?>