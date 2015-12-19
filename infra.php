<?php
namespace infrajs\doc;

use infrajs\event\Event;
use infrajs\path\Path;
use infrajs\infra\Infra;

$conf=&Config::get('doc');
$conf=array_merge(Mht::$conf, Docx::$conf, $conf);
Docx::$conf=$conf;
Mht::$conf=$conf;

Event::handler('oninstall', function () {
	Path::mkdir(Docx::$conf['cache']);
	Path::mkdir(Mht::$conf['cache']);
});