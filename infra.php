<?php
namespace infrajs\doc;

use infrajs\event\Event;
use infrajs\path\Path;

Event::handler('oninstall', function () {
	Path::mkdir(Docx::$conf['cache']);
});