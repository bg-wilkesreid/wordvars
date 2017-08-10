<?php

namespace WordVars;

class WordVars
{
  private $templatepath;
  private $outputpath;
  private $codes;

  public function __construct($templatepath = "template.docx", $outputpath = "output.docx") {
    $this->templatepath = $templatepath;
    $this->outputpath = $outputpath;
  }

  public function setCodes($codes) {
    $this->codes = $codes;
  }

  public function getCodes() {
    return $this->codes;
  }

  public function setCalculatedCodes($calculated_codes) {
    foreach ($calculated_codes as $calculated_code=>$calcvalue) {
      foreach ($this->codes as $code=>$codevalue) {
        $calcvalue = preg_replace('/#\['.$code.'\]/', $codevalue, $calcvalue);
      }
      $this->codes[$calculated_code] = $calcvalue;
    }
  }

  public function go() {
    copy($this->templatepath, $this->outputpath);
    $zip = new \ZipArchive;
    $zip->open($this->outputpath);

    $content = $zip->getFromName('word/document.xml');

    // Do replacement
    foreach ($this->codes as $code=>$value) {
      $content = preg_replace('%#\[(?:</w:t></w:r>(?:<w:proofErr[^/]*/>)?<w:r[^>]*><w:t[^>]*>)?'.$code.'(?:</w:t></w:r><w:proofErr[^>]*/><w:r[^>]*><w:t[^>]*>)?\]%', $value, $content);
    }

    $zip->deleteName('word/document.xml');
    $zip->addFromString('word/document.xml', $content);

    $zip->close();
  }
}
