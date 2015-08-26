<?php

infra_test(true);

use itlife\files\Docx;
use itlife\files\Mht;
use itlife\files\Xlsx;
use itlife\infra\ext\Ans;

$text = Docx::get('*files/tests/resources/test.docx');
if (!$text || strlen($text) != 1388) {
	return Ans::err($ans, 'Cant read file .docx');
}
$preview = Docx::preview('*files/tests/resources/test.docx');
if (sizeof($preview) != 12) {
	return Ans::err($ans, 'Cant read preview test.docx '.sizeof($preview));
}
if (sizeof($preview['links']) != 4) {
	return Ans::err($ans, 'Cant read links test.docx');
}
if (sizeof($preview['images']) != 1) {
	return Ans::err($ans, 'Cant read images test.docx');
}
if (strlen($preview['preview']) != 199) {
	return Ans::err($ans, 'Cant read test.docx preview');
}

$name = 'test.tpl';
$preview = Mht::preview('*files/tests/resources/'.$name);
if (sizeof($preview) != 12) {
	return Ans::err($ans, 'Cant read preview '.$name.' '.sizeof($preview));
}
if (sizeof($preview['links']) != 1) {
	return Ans::err($ans, 'Cant read links '.$name.' '.sizeof($preview['links']));
}
if (sizeof($preview['images']) != 2) {
	return Ans::err($ans, 'Cant read images '.$name.' '.sizeof($preview['images']));
}
if (strlen($preview['preview']) != 899) {
	return Ans::err($ans, 'Cant read '.$name.' preview '.strlen($preview['preview']));
}

$name = 'test.tpl';
$text = Mht::get('*files/tests/resources/'.$name);
if (strlen($text) != 2891) {
	return Ans::err($ans, 'Cant read '.$name.' '.strlen($text));
}

$name = 'test.html';
$text = Mht::get('*files/tests/resources/'.$name);
if (strlen($text) != 1073) {
	return Ans::err($ans, 'Cant read '.$name.' '.strlen($text));
}

$data = Xlsx::init('*files/tests/resources/test.xlsx');

if (!$data) {
	return Ans::err($ans, 'Cant read test.xlsx');
}
$data = Xlsx::get('*files/tests/resources/test.xls');

if (sizeof($data['childs'][0]['data']) != 30) {
	return Ans::err($ans, 'Cant read test.xls '.sizeof($data['childs'][0]['data']));
}

return Ans::ret($ans, 'tpl, mht, docx, xls, xlsx read ok!');
