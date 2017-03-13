<?php
include_once "data/dbconfig.php";
$connect_db = mysql_connect(T1_MYSQL_HOST, T1_MYSQL_USER, T1_MYSQL_PASSWORD) or die('MySQL Connect Error!!!');
$select_db  = mysql_select_db(T1_MYSQL_DB, $connect_db) or die('MySQL DB Error!!!');
mysql_query(" set names utf8 ");

function get_point($mb_hp)
{
	$sql="select * from drive_point where mb_hp='{$mb_hp}';";
	$row=mysql_fetch_assoc(mysql_query($sql));
	return $row['po_mb_point'];
}


/*
createColumnsArray 
PHPexcel reage
A-AZ
ex)  createColumnsArray('AZ');
*/
function createColumnsArray($end_column, $first_letters = '')
{
  $columns = array();
  $length = strlen($end_column);
  $letters = range('A', 'Z');

  // Iterate over 26 letters.
  foreach ($letters as $letter) {
      // Paste the $first_letters before the next.
      $column = $first_letters . $letter;

      // Add the column to the final array.
      $columns[] = $column;

      // If it was the end column that was added, return the columns.
      if ($column == $end_column)
          return $columns;
  }

  // Add the column children.
  foreach ($columns as $column) {
      // Don't itterate if the $end_column was already set in a previous itteration.
      // Stop iterating if you've reached the maximum character length.
      if (!in_array($end_column, $columns) && strlen($column) < $length) {
          $new_columns = createColumnsArray($end_column, $column);
          // Merge the new columns which were created with the final columns array.
          $columns = array_merge($columns, $new_columns);
      }
  }

  return $columns;
}


?>