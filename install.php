<?php

$dirs = infra_dirs();
if (!is_dir($dirs['cache'].'docx/')) {
	mkdir($dirs['cache'].'docx/');
}
if (!is_dir($dirs['cache'].'mht/')) {
	mkdir($dirs['cache'].'mht/');
}
