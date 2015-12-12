<?php

namespace infrajs\doc;
use infrajs\access\Access;
use infrajs\ans\Ans;

Access::test(true);

$text = Docx::get('*files/tests/resources/test.docx');


if (!$text || mb_strlen($text) != 1056) {
	return Ans::err($ans, 'Cant read file .docx mb_strlen '.mb_strlen($text));
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
if (mb_strlen($preview['preview']) != 119) {
	return Ans::err($ans, 'Cant read test.docx preview '.mb_strlen($preview['preview']));
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
if (mb_strlen($preview['preview']) != 521) {
	return Ans::err($ans, 'Cant read '.$name.' preview '.strlen($preview['preview']));
}

$name = 'test.tpl';
$text = Mht::get('*files/tests/resources/'.$name);
if (mb_strlen($text) != 1935) {
	return Ans::err($ans, 'Cant read '.$name.' '.strlen($text));
}

$name = 'test.html';
$text = Mht::get('*files/tests/resources/'.$name);
if (strlen($text) != 1073) {
	return Ans::err($ans, 'Cant read '.$name.' '.strlen($text));
}


return Ans::ret($ans, 'tpl, mht, docx read ok!');
