<?php
/**
* 2017-03-13 (월) 11:07:04 
* www/test/php_excel.php
*
*/
include $_SERVER["DOCUMENT_ROOT"]."/lib/PHPExcel-1.8/Classes/PHPExcel.php";
include_once $_SERVER["DOCUMENT_ROOT"]."/common.php";

$UpFile	= $_FILES["upfile"];
$UpFileName = $UpFile["name"];

$UpFilePathInfo = pathinfo($UpFileName);
$UpFileExt		= strtolower($UpFilePathInfo["extension"]);

if($UpFileExt != "xls" && $UpFileExt != "xlsx") {
	echo "엑셀파일만 업로드 가능합니다. (xls, xlsx 확장자의 파일포멧)";
	exit;
}

//-- 읽을 범위 필터 설정 (아래는 A열만 읽어오도록 설정함  => 속도를 중가시키기 위해)
class MyReadFilter implements PHPExcel_Reader_IReadFilter
{
	public function readCell($column, $row, $worksheetName = '') {
		// Read rows 1 to 7 and columns A to E only

		if (in_array($column,array('B','U'))) {
			return true;
		}
		return false;
	}
}

//-- 읽을 범위 필터 설정 (아래는 A열만 읽어오도록 설정함  => 속도를 중가시키기 위해)
class MyReadFilter2 implements PHPExcel_Reader_IReadFilter
{
	public function readCell($column, $row, $worksheetName = '') {
		// Read rows 1 to 7 and columns A to E only

		if (in_array($column,createColumnsArray('AZ'))) {
			return true;
		}
		return false;
	}
}

$filterSubset = new MyReadFilter();
$filterSubset2 = new MyReadFilter2();
//업로드된 엑셀파일을 서버의 지정된 곳에 옮기기 위해 경로 적절히 설정
$upload_path = $_SERVER["DOCUMENT_ROOT"]."/data/uploads/Excel_".date("Ymd");



if (!file_exists($upload_path)) {
    mkdir($upload_path, 0777, true);
}

$upfile_path = $upload_path."/".date("Ymd_His")."_".$UpFileName;
if(is_uploaded_file($UpFile["tmp_name"])) {

	if(!move_uploaded_file($UpFile["tmp_name"],$upfile_path)) {
		echo "업로드된 파일을 옮기는 중 에러가 발생했습니다.";
		exit;
	}

	//파일 타입 설정 (확자자에 따른 구분)
	$inputFileType = 'Excel2007';
	if($UpFileExt == "xls") {
		$inputFileType = 'Excel5';	
	}

	//1.1. 엑셀리더 초기화
	$objReader = PHPExcel_IOFactory::createReader($inputFileType);

	//1.2. 데이터만 읽기(서식을 모두 무시해서 속도 증가 시킴)
	$objReader->setReadDataOnly(true);	

	//1.3. 범위 지정(위에 작성한 범위필터 적용)
	$objReader->setReadFilter($filterSubset);

	//1.4. 업로드된 엑셀 파일 읽기
	$objPHPExcel = $objReader->load($upfile_path);

	//1.5. 첫번째 시트로 고정
	$objPHPExcel->setActiveSheetIndex(0);

	//1.6. 고정된 시트 로드
	$objWorksheet = $objPHPExcel->getActiveSheet();

	//1.7. 시트의 지정된 범위 데이터를 모두 읽어 배열로 저장
	$sheetData = $objPHPExcel->getActiveSheet()->toArray(null,true,true,true);
	$total_rows = count($sheetData);
	


	//2.1. 엑셀리더 초기화
	$objReader2 = PHPExcel_IOFactory::createReader($inputFileType);

	//2.2. 데이터만 읽기(서식을 모두 무시해서 속도 증가 시킴)
	$objReader2->setReadDataOnly(true);	

	//2.3. 범위 지정(위에 작성한 범위필터 적용)
	$objReader2->setReadFilter($filterSubset2);

	//2.4. 업로드된 엑셀 파일 읽기
	$objPHPExcel2 = $objReader->load($upfile_path);

	//2.5. 첫번째 시트로 고정
	$objPHPExcel2->setActiveSheetIndex(0);

	//2.6. 고정된 시트 로드
	$objWorksheet2 = $objPHPExcel2->getActiveSheet();

	//2.7. 시트의 지정된 범위 데이터를 모두 읽어 배열로 저장
	$sheetData2 = $objPHPExcel2->getActiveSheet()->toArray(null,true,true,true);

	$key_code=array();
	foreach($sheetData as $rows) 
	{
		if($rows["U"]=="자동배차")
		{
			$key_code[]=$rows["B"];
		}
	}
	$sql=array();
	$keys=join(",",$key_code);
	$sql[]="select po_odno from drive_point ";
	$sql[]="where po_odno not in ({$keys});";
	echo join("",$sql);
	$query=mysql_query(join("",$sql));
/*
	while($list=mysql_fetch_assoc($query))
	{
		echo $list['po_odno'];
	}
	*/
/*
	foreach($key_code as $k)
	{
		echo $k;
	}


	foreach($sheetData2 as $rows) 
	{
			$key_code[]=$rows["B"];
			$sql=array();
			$date = str_replace("/", "-",$rows["J"]);
			$po_datetime=date("Y")."-".$date." ".$rows["K"];
			$po_mb_point=get_point($rows["G"])+1000;
			$sql[]="insert into drive_point set ";
			$sql[]="po_odno='".$rows["B"]."',";
			$sql[]="po_datetime='".$po_datetime."',";
			$sql[]="po_count='".$rows["F"]."',";
			$sql[]="mb_hp='".$rows["G"]."',";
			$sql[]="po_point='1000',";
			$sql[]="po_mb_point='{$po_mb_point}';";
			echo $rows["AM"];
			echo join("",$sql);
			$result=mysql_query(join("",$sql));
		
	}
	*/	
}


?>