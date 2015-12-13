<?php

require_once('PhpOffice/PHPExcel/IOFactory.php');
require_once('PhpOffice/PHPExcel/PHPExcel.php');

class Xls2Doc {

  function process($template_doc, $fname, $out_dir) {

    $values = $this->readSheet($fname);
    if (empty($values)) {
      return false;
    }

    $x = explode('.', basename($fname));
    $values['fname'] = $x[0];

    $templateProcessor = new \PhpOffice\PhpWord\TemplateProcessor($template_doc);

    foreach($values as $key => $value) {
      $templateProcessor->setValue($key, htmlspecialchars($value));
    }
    $save_as = $out_dir.$x[0].'.'.basename($template_doc);
    if (is_file($save_as)) {
      $dt = (new DateTime('now', new DateTimeZone('america/toronto')))->format('Ymd_His');
      $ext = pathinfo($save_as, PATHINFO_EXTENSION);
      $bu_fname = $save_as.".$dt.$ext";
      // die("$dt  $save_as  $bu_fname");
      rename($save_as, $bu_fname);
    }
    $templateProcessor->saveAs($save_as);

    return $save_as;

  }

  function readSheet($fname) {
    $res = [];
    $objPHPExcel = PHPExcel_IOFactory::load($fname);
    // scan col A to find "A".
    for($row=1; $row<500; $row++) {
      $cellValue = $objPHPExcel->getActiveSheet()->getCell('A'.$row)
        ->getValue();
      if (strtoupper($cellValue)=='A') {
        $k_val = $objPHPExcel->getActiveSheet()->getCell('B'.$row)
          ->getValue();
        $v_val = $objPHPExcel->getActiveSheet()->getCell('C'.$row)
          ->getValue();
        if ($k_val) {
          $res[$k_val] = $v_val;
        }
      }
    }
    return $res;
  }

}
