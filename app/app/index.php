<?php

spl_autoload_register(function($class) {
  $dir = __DIR__.'/';
  // die($dir.$class);
  if (is_file($file = $dir.$class.'.php')
    || is_file($file = $dir.strtolower($class).'.php')
    || is_file($file = strtolower($dir.$class).'.php')
    ) {
      return require($file);
    }
});

// filter ms backup file ~*
function good_glob($mask) {
  return array_filter( glob($mask), function($fname) {
    return strpos( basename($fname), '~') === false;
  });
}

function alert($type, $msg) {
  return "<div class=\"alert alert-$type\">$msg</div>";
}

$conf = require('conf.php');
$log = [];

// print_r($_POST);
// die();

///////// gather templates

$template_dir = $conf->doc_root.'templates/';

$templates_xls = good_glob($template_dir.'*.xls*');
$templates_doc = good_glob($template_dir.'*.doc*');

///////// process request

if (isset($_POST['qa-new-qa'])) {
  $template_fname = $_POST['qa-new-template'];
  $fname = $conf->doc_root.'in/'
    . str_replace(' ', '_', $_POST['qa-name'])
    . '.'.$template_fname;

  $template_xls = $template_dir.$template_fname;

  if (file_exists($fname)) {
    $log[] = alert('danger', "$fname exists already!");
  } elseif (!copy($template_xls, $fname)) {
    $log[] = alert('danger', "$fname creation from $template_xls failed.");
  } else {
    $log[] = alert('success', "$fname created from $template_xls.");
  }

}

if (isset($_POST['qa-process'])) {
  $template_doc = $template_dir.$_POST['qa-process-template'];
  $inbox = glob( $conf->doc_root.'in/*.xls*');
  include_once('xls2doc.php');
  $xls2doc = new Xls2Doc();

  foreach($_POST['qa-inbox-fnames'] as $fname) {
    if ($out_fname = $xls2doc->process(
        $template_doc,
        $conf->doc_root.'in/'.$fname,
        $conf->doc_root.'out/')
    ) {
      // rename($fname, $conf->doc_root.'done/'.basename($fname));
      $log[] = alert('success', "$out_fname created from $template_doc.");
    } else {
      $log[] = alert('danger', $xls2doc->lastError());
    }

  }

}


////////// gather file lists

$inbox_dir = good_glob( $conf->doc_root.'in/*.xls*');
$out_dir = good_glob( $conf->doc_root.'out/*.doc*');

// echo $conf->doc_root;
// print_r($current_dir);
// die();

// $inbox_files = implode('<br>', array_map('basename', $inbox_dir));
$out_files = implode('<br>', array_map('basename', $out_dir));

$process_enabled = empty($templates_doc) || empty($inbox_dir)
  ? ' disabled'
  : '';

$title = 'XLS to DOC convertor';

$log = implode('', $log);

$inbox_checkboxes = '<div>'
  .implode('', array_map(function($fname) {
    return '<label><input type="checkbox" name="qa-inbox-fnames[]" value="'
      .basename($fname).'" /> '.basename($fname).'</label><br/>';
  }, $inbox_dir))
  .'</div>';

$xls_radios = '<div class="radio-group">'
  .implode('', array_map(function($fname) {
    return '<label><input type="radio" name="qa-new-template" value="'
      .basename($fname).'" /> '.basename($fname).'</label><br/>';
  }, $templates_xls))
  .'</div>';

$doc_radios = '<div class="radio-group">'
  .implode('', array_map(function($fname) {
    return '<label><input type="radio" name="qa-process-template" value="'
      .basename($fname).'" /> '.basename($fname).'</label><br/>';
  }, $templates_doc))
  .'</div>';

echo <<<EOT

<!DOCTYPE html>
<head>
  <!-- Latest compiled and minified CSS -->
  <link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.6/css/bootstrap.min.css" integrity="sha384-1q8mTJOASx8j1Au+a5WDVnPi2lkFfwwEAa8hDDdjZlpLegxhjVME1fgjWPGmkzs7" crossorigin="anonymous">
  <script src="//code.jquery.com/jquery-1.11.3.min.js"></script>
  <style>
    .btn.pull-right {
      margin: 12px 0;
    }
    input.pull-right {
      margin-right: -12px;
    }
    .err {
      font-size: 300%;
      color: red;
      display: none;
      vertical-align: middle;
    }
  </style>
</head>
<body>

<div class="container">
  <h1>$title</h1>

  <a id="btn-refresh" href="" class="btn btn-primary pull-right">Refresh</a>


  <div class="row">
    <div class="col-xs-12">
      $log
    </div>
  </div>

  <form action="" method="post" id="qa-form-new" >
    <input type="hidden" name="qa-new-qa" value="1" />

    <div class="row alert alert-warning" >
      <div class="col-xs-6">
        <h4>Q&A templates <span id="err-new-template" class="err">*</span></h4>
        <span id="err-new-template" class="err">*</span>
        $xls_radios
      </div>
      <div class="col-xs-6">
        <div class="pull-right">
          <label>Questionaire name
            <span id="err-new-name" class="err">*</span><br/>
            <input  type="textbox" id="qa-name" name="qa-name" value="" />
          </label><br/><br/>
          <button class="btn btn-primary" id="btn-qa-name">Create new questionaire</button>
        </div>
      </div>

    </div>

  </form>

  <form action="" method="post" id="qa-form-process">

    <div class="row">
      <div class="col-xs-12">
        <div class="panel">
          <h3>In <span id="err-inbox-fnames" class="err">*</span></h3>
          $inbox_checkboxes
        </div>
      </div>
    </div>

    <div class="row alert alert-warning">
      <div class="col-xs-12">
        <input type="hidden" name="qa-process" value="1" />
        <h4>Output templates <span id="err-process-template" class="err">*</span></h4>
        $doc_radios
      </div>
      <p>
        <button id="btn-qa-process" class="btn btn-primary{$process_enabled} pull-right">Process Inbox</button>
      </p>
    </div>

  </form>

  <div class="row">
    <div class="col-xs-12">

      <p>
      </p>
      <div class="panel">
        <h3>Output - Ready for review</h3>
        <div>
          $out_files
        </div>
      </div>
    </div>

  </div>

  <script>
  $(document).ready( function() {

    $('#btn-qa-name').on('click', function() {
      if (! $("#qa-name").val()) {
        $('#err-new-name').show();
        return false;
      }
      if (! $('input:radio[name=qa-new-template]:checked').val()) {
        $('#err-new-template').show();
        return false;
      }
      $('#qa-form-new').submit();
    });

    $('#btn-qa-process').on('click', function() {

      try {

        // alert($('input:checkbox[name=qa-inbox-fnames:checked').val());
        if (! $('#qa-form-process input[type=checkbox]:checked').val() ) {
          $('#err-inbox-fnames').show();
          return false;
        }

        // alert($('input:radio[name=qa-process-template]:checked').val());
        if (! $('input:radio[name=qa-process-template]:checked').val()) {
          $('#err-process-template').show();
          return false;
        }
        $('#qa-form-process').submit();

      } catch(e) {
        console.log(e);
        return false;
      }

    });



  });

  </script>

</div>

</body>

EOT;
