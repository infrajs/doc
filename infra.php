<?php
namespace infrajs\doc;
use infrajs\infra\Infra;

$conf=&Infra::config('doc');
$conf=array_merge(Mht::$conf, Docx::$conf, $conf);
Docx::$conf=$conf;
Mht::$conf=$conf;
