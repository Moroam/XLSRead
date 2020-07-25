<?php

/**
 * Class XLSRead
 *
 * @version 1.0.0
 */

class XLSRead {
  static protected $error = "";
  static public function get_error(){
    return self::$error;
  }

  /**
   * Static function for checking file is it XLS table
   *
   * @param file $FILE     file wich we check
   * @return bool
   */
  static public function is_xls_file($FILE){
    $xls_mtype=['application/vnd.ms-excel',
          'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
          'application/excel',
          'application/msexcel',
          'application/x-excel',
          'application/x-msexcel',
          'application/x-ms-excel',
          'application/xls',
          'text/xml',
          'text/csv',
          'application/vnd.ms-excel.sheet.macroEnabled.12',
          'application/vnd.oasis.opendocument.spreadsheet',  #ods - open doc sheet
          'application/wps-office.xlsx',
          'application/wps-office.xls'
          ];
    $ext_type = ['xlsx','xls','xlsm','csv','xml','ods'];

    self::$error = "";

    if ($FILE['error'] == UPLOAD_ERR_OK ) { //проверка на наличие ошибок
      $mtype = $FILE['type'];
      if(!in_array($mtype,$xls_mtype)){
        self::$error = 'File '.$FILE['name'].' has wrong mime-type: '.$FILE['type'];
        return false;
      }
      $ext = pathinfo($FILE['name'],PATHINFO_EXTENSION);
      if(!in_array($ext,$ext_type)){
        self::$error = 'File '.$FILE['name'].' has wrong extantion: '.$ext;
        return false;
      }
    } else {
      switch ($FILE['error']) {
        case UPLOAD_ERR_FORM_SIZE:
        case UPLOAD_ERR_INI_SIZE:
          $msg = 'File '.$FILE['name'].' Size exceed';
          break;
        case UPLOAD_ERR_NO_FILE:
          $msg = 'FIle '.$FILE['name'].' Not selected';
          break;
        default:
          $msg = 'Something is wrong with file '.$FILE['name'];
      }
      self::$error = $msg;
      return false;
    }

    return true;
  } #END is_xls_file

  /**
   * Static function for reading XLS file in ARRAY
   *
   * @param file $FILE               file wich we read
   * @param string $B                last column
   * @param string $sheet            number worksheet for reading
   * @param string $ExplicitColumns  list
   * @return bool or mysql_result
   */
  # $sheet -> number of the sheet to read, by default select active sheet
  static public function xls_to_array($FILE, string $B='', string $sheet='', string $ExplicitColumns = "", bool $unsetSpreadSheet = TRUE){
    if(!self::is_xls_file($FILE)){
      return false;
    }

    $mtype = $FILE['type'];
    $ext = ucfirst(pathinfo($FILE['name'], PATHINFO_EXTENSION));
    # create reader for XLS or XLSX files - by MIME type
    if($mtype=='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      || $mtype=='application/vnd.ms-excel.sheet.macroEnabled.12'
      || $mtype=='application/wps-office.xlsx')
      $reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
    elseif(($mtype=='application/vnd.ms-excel' || $mtype=='text/csv') && $ext=='Csv')
      $reader = new \PhpOffice\PhpSpreadsheet\Reader\Csv();
    elseif($mtype=='application/vnd.oasis.opendocument.spreadsheet' && $ext=='Ods')
      $reader = new \PhpOffice\PhpSpreadsheet\Reader\Ods();
    else
      $reader = new \PhpOffice\PhpSpreadsheet\Reader\Xls();

    if($ExplicitColumns == "")
      $reader->setReadDataOnly(TRUE);

    $inputFileName = $FILE['tmp_name'];

    $spreadsheet = $reader->load($inputFileName);

    if( $sheet != '' && $sheet >= $spreadsheet->getSheetCount()){
      return [];
    }

    if( $sheet == '' || $ext == 'Csv'){
      $worksheet = $spreadsheet->getActiveSheet();
    } else{
      $worksheet = $spreadsheet->getSheet($sheet);
    }

    $lastRow = $worksheet->getHighestRow();
    $lastCol = ($B == '' ? $worksheet->getHighestColumn() : $B);

    if($lastRow == 1  && $worksheet->getHighestColumn() == 'A'){
      return [];
    }

    if($ExplicitColumns != ''){
      $columns = explode(',', $ExplicitColumns);
      foreach ($columns as $col)
        $spreadsheet->getActiveSheet()->getStyle($col.':'.$col)
            ->getNumberFormat()
            ->setFormatCode('#');
    }


    $dataArray = $worksheet
        ->rangeToArray(
            'A1:'.$lastCol.$lastRow,     // The worksheet range that we want to retrieve
            NULL,        // Value that should be returned for empty cells
            TRUE,        // Should formulas be calculated (the equivalent of getCalculatedValue() for each cell)
            TRUE,        // Should values be formatted (the equivalent of getFormattedValue() for each cell)
            TRUE         // Should the array be indexed by cell row and cell column
        );

    if($unsetSpreadSheet){
      unset($spreadsheet);
    }

    return $dataArray;
  }#end xls_to_array

  /**
   * Delete column from associate array by key name
   *
   * @param array &$array
   * @param string $key
   */
  static public function array_delete_col(&$array, string $key) {
      return array_walk($array, function (&$v) use ($key) {
          unset($v[$key]);
      });
  }

  /**
   * Get column name by number for XLS files
   *
   * @param int $nom column number
   * @return string
   */
  static public function column_name_by_number( int $nom ) {
      return $nom > 26 ? chr(64 + intdiv($nom, 26)).chr(64 + $nom % 26) : chr(64 + $nom);
  }

  /**
   * Get column number by name for XLS files
   *
   * @param string $let column name
   * @return number
   */
  static public function column_number_by_name( string $let ) {
    $ord = ord(strtolower($let[0])) - 96;
      return strlen($let) > 1 ? $ord * 26 + ord(strtolower($let[1])) - 96 : $ord;
  }

  /**
   * Select one file information from global array $_FILES with multiple files
   * @param string $inputfile name input type="file"
   * @param int $num file number
   * @return array
   */
  static public function file_by_num(string $inputfile, int $num = 0 ) : array {
    return array_combine(
      array_keys($_FILES[$inputfile]),
      array_column($_FILES[$inputfile], $num)
    );
  }


}
