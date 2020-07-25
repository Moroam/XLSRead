<!DOCTYPE html>
<html lang=ru>
<meta charset=utf-8>
<head>
  <title>Project - Import XLS</title>
</head>

<link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.4.1/css/bootstrap.min.css" integrity="sha384-Vkoo8x4CGsO3+Hhxv8T/Q5PaXtkKtu6ug5TOeNV6gBiFeWPGFN9MuhOf23Q9Ifjh" crossorigin="anonymous">

<style type="text/css">
  form {
    width: 50%;
    min-width: 20em;
    max-width: 100em;
  }
</style>

<body>


<?php
require '../vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\Spreadsheet;

require_once 'xlsread.class.php';

error_reporting(E_ALL);
if(ini_set('display_errors', 1)===false)
  echo "ERROR INI SET";

?>
<form method='post' enctype='multipart/form-data' class="border rounded mx-auto p-3 mt-5">
  <h3 class="text-center">Импорт файлов - XLSRead</h3>

  <div class="form-group">
    <input type="file" class="form-control-file" id="inputfile" name='inputfile[]' multiple>
  </div>

  <div class="text-center">
      <input type='submit' class="btn btn-primary"   name='btnimport' value='Импорт' />
      <input type='reset'  class="btn btn-secondary" name='reset'     value='Очистить' />
  </div>
</form>

<?php



if (!isset($_POST['btnimport'])){
  goto END;
}

$cntMaxFileUploads = (int)(ini_get('max_file_uploads'));
$cntF = count($_FILES['inputfile']['name']);
if($cntF > $cntMaxFileUploads){
  echo "<div style='text-align: center;'><h3>Выбрано СЛИШКОМ много - <b>$cntF</b>!!! Будут отработаны только первые $cntMaxFileUploads файлов.</h3></div>";
  $cntF = $cntMaxFileUploads;
}

?>
<div style="text-align: center;">
  <div style="margin:0 auto;">
    <h4>Результат импорта</h4>
    <b>Количество файлов: </b><?=$cntF?>
  </div>


  <?php
  for($i=0; $i < $cntF; ){
    $F = XLSRead::file_by_num('inputfile', $i++);

    $xlsarray = XLSRead::xls_to_array($F);

    echo '<h5>' . $i . '. File - "' . $F['name'] . '"</h5>';
    echo '<table class="table table-striped table-bordered table-sm">';
    foreach ($xlsarray as $row_num => $row) {
      echo "<tr>";
      foreach ($row as $col_name => $value) {
        echo "<td>" . $value . "</td>";
      }
      echo "</tr>";
    }
    echo "<table>";


  }
  ?>

</div>


<?php
END:
?>

</body>
</html>
