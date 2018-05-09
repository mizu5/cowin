<?php
$sub_menu = '400400';
include_once('./_common.php');

auth_check($auth[$sub_menu], "w");


include_once(G5_PATH.'/head.sub.php');

//echo $sm;
?>

<div class="new_win">
<div style="text-align:center;padding:10px 0;">
<?php 
if($sm == "wmp"){
echo "<img src='/img/wmp_logo.png'/>";
$sm_title = "위메프";
}else if($sm == "tm"){
echo "<img src='/img/tm_logo.png'/>";
$sm_title = "티몬";
}

$g5['title'] = $sm_title.'엑셀파일로  주문 일괄 등록';
?>
</div>
    <h1><?php echo $g5['title']; ?></h1>

    <div class="local_desc01 local_desc">
        <p>
            엑셀파일을 이용하여 주문 일괄등록<br>
            형식은 <strong>주문일괄등록용 엑셀파일</strong>을 다운로드하여 상품 정보를 입력하시면 됩니다.<br>
            수정 완료 후 엑셀파일을 업로드하시면 주문 일괄등록됩니다.<br>
            엑셀파일을 저장하실 때는 <strong>Excel 97 - 2003 통합문서 (*.xls)</strong> 로 저장하셔야 합니다.
        </p>

        <p>
            <a href="<?php echo G5_URL; ?>/<?php echo G5_LIB_DIR; ?>/Excel/itemexcel.xls">주문일괄등록용 엑셀파일 다운로드</a>
        </p>
    </div>

    <form name="fitemexcel" method="post" action="./cateexcelupdate2.php?sm=<?php echo $sm?>" enctype="MULTIPART/FORM-DATA" autocomplete="off">

    <div id="excelfile_upload">
        <label for="excelfile">파일선택</label>
        <input type="file" name="excelfile" id="excelfile">
    </div>

    <div class="win_btn btn_confirm">
        <input type="submit" value="주문 엑셀파일 등록" class="btn_submit btn">
        <button type="button" onclick="window.close();" class="btn_close btn">닫기</button>
    </div>

    </form>

</div>

<?php
include_once(G5_PATH.'/tail.sub.php');
?>