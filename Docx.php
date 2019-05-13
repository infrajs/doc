<?php
namespace infrajs\doc;

use infrajs\cache\Cache as OldCache;
use akiyatkin\boo\Cache;
use infrajs\path\Path;
use infrajs\load\Load;

class Docx
{
	public static $conf=array(
		"imgmaxwidth" => 1000,
		"previewlen" => 150,
		'cache'=>'!doc/'
	);
	public static function getPreview($html)
	{
		$temphtml = strip_tags($html, '<p><a><b><i>');
		//preg_match('/^(<p.*>.{'.$previewlen.'}.*<\/p>)/U',$temphtml,$match);
		preg_match('/(<p.*>.{1}.*<\/p>)/U', $temphtml, $match);
		if (sizeof($match) > 1) {
			$preview = $match[1];
		} else {
			$preview = $html;
		}
		$preview = preg_replace('/<h1[^>]*>.*<\/h1>/iU', '', $preview);
		$preview = preg_replace('/<img[^>]*>/iU', '', $preview);
		$preview = preg_replace('/<p[^>]*>\s*<\/p>/iU', '', $preview);
		$preview = preg_replace("/\s+/", ' ', $preview);
		$preview = trim($preview);
		return $preview;
	}
	public static function preview($src)
	{
		$param = self::parse($src);

		$data = Load::srcInfo($src);
		unset($data['query']);
		//$data = Load::nameInfo($data['file']);

		
		$preview = Docx::getPreview($param['html']);
		
		/*preg_match('/<img.*src=["\'](.*)["\'].*>/U', $param['html'], $match);
		if ($match && $match[1]) {
			$img = $match[1];
		} else {
			$img = false;
		}*/
		$filetime = filemtime(Path::theme($src));
		$data['modified'] = $filetime;
		if (!empty($param['links'])) {
			$data['links'] = $param['links'];
		}
		if (!empty($param['heading'])) {
			$data['heading'] = $param['heading'];
		}
		//title - depricated
		if (!empty($data['name'])) {
			$data['title'] = $data['name'];
		}
		/*if ($img) {
			$data['img'] = $img;
		}*/
		if (!empty($param['images'])) {
			$data['images'] = $param['images'];
		}

		$data['preview'] = $preview;

		return $data;
	}
	public static function get($src)
	{
		//От логика infra-com отказались
		/*
			Содержмое txt не может повлиять на работу infrajs
			Если мы хотим иметь такую новость, которая будет менять сайт, нужно заложить эту логику в layers.json в данные
		*/
		$param = self::parse($src);

		return $param['html'];
	}
	/**
	 * Кэширумеая функция, основной разбор.
	 */
	public static function parse($src)
	{
		$args = array($src);
		$param = Cache::exec('Разбор документов Word', function ($src) {
			$conf = Docx::$conf;
			$imgmaxwidth = $conf['imgmaxwidth'];
			$previewlen = $conf['previewlen'];

			$cachename = Path::encode($src);
			$cacheFolder = Path::mkdir(Docx::$conf['cache'].$cachename.'/');
		
//В винде ингда вылетает о шибка что нет прав удалить какой-то файл в папке и как следствие саму папку
			//Обновление страницы проходит уже нормально
			//Полагаю в линукс такой ошибки не будет хз почему возникает
			OldCache::fullrmdir($cacheFolder);
			

			$path=Path::theme($src);


			if (!$path) return array('html'=>false);
			$xmls = docx_getTextFromZippedXML($path, 'word/document.xml', $cacheFolder);
			
			$rIds = array();
			$param = array('folder' => $cacheFolder, 'imgmaxwidth' => $imgmaxwidth, 'previewlen' => $previewlen, 'rIds' => $rIds);
			if ($xmls[0]) {
				$xmlar = docx_dom_to_array($xmls[0]);
				$xmlar2 = docx_dom_to_array($xmls[1]);

				foreach ($xmlar2['Relationships']['Relationship'] as $v) {
					$rIds[$v['Id']] = $v['Target'];
				}

				$param['rIds'] = $rIds;
				$html = docx_each($xmlar, '\\infrajs\\doc\\docx_analyse', $param);
			} else {
				$param['rIds'] = array();
				$html = '';
			}

			$param['html'] = $html;
			unset($param['rIds']);

			unset($param['type']);
			unset($param['imgmaxwidth']);
			unset($param['previewlen']);
			unset($param['isli']);
			unset($param['isul']);
			unset($param['imgnum']);
			unset($param['folder']);
			return $param;
		}, $args, ['akiyatkin\boo\Cache','getModifiedTime'], array($src));

		return $param;
	}
}



function docx_full_del_dir($directory)
{
	if (!$directory) return;
	$dir = opendir($directory);
	if (!$dir) return;
	while ($file = readdir($dir)) {
		if (is_file($directory.'/'.$file)) {
			unlink($directory.'/'.$file);
		} elseif (is_dir($directory.'/'.$file) && $file !== '.' && $file !== '..') {
			docx_full_del_dir($directory.'/'.$file);
		}
	}
	closedir($dir);
	rmdir($directory);
}
function docx_getTextFromZippedXML($archiveFile, $contentFile, $cacheFolder)
{
	// Создаёт "реинкарнацию" zip-архива...

	$zip = new \ZipArchive();

	// И пытаемся открыть переданный zip-файл
	if ((int) phpversion() > 6) {
		$archiveFile = realpath($archiveFile);
		$archiveFile = Path::tofs($archiveFile);

		$cacheFolder = realpath($cacheFolder);
		$cacheFolder = Path::tofs($cacheFolder);

		if (!empty($_SERVER['WINDIR'])) { //Только для Виндовс
			$archiveFile = Path::toutf($archiveFile);
			//$cacheFolder = Path::toutf($cacheFolder);
		}
	}
	if ($zip->open($archiveFile) === true) {

		$zip->extractTo($cacheFolder);
		// В случае успеха ищем в архиве файл с данными
		$xml = false;
		$xml2 = false;
		$file = $contentFile;

		if (($index = $zip->locateName($file)) !== false) {
			// Если находим, то читаем его в строку
			$content = $zip->getFromIndex($index);
			// После этого подгружаем все entity и по возможности include'ы других файлов
			$xml = new \DOMDocument();
			$xml->loadXML($content, LIBXML_NOENT | LIBXML_XINCLUDE | LIBXML_NOERROR | LIBXML_NOWARNING);
		}
		$file = 'word/_rels/document.xml.rels';
		if (($index = $zip->locateName($file)) !== false) {
			// Если находим, то читаем его в строку
			$content = $zip->getFromIndex($index);
			$xml2 = new \DOMDocument();
			$xml2->loadXML($content, LIBXML_NOENT | LIBXML_XINCLUDE | LIBXML_NOERROR | LIBXML_NOWARNING);
		//@ - https://bugs.php.net/bug.php?id=41398 Strict Standards:  Non-static method DOMDocument::loadXML() should not be called statically
		}
		$zip->close();

		return array($xml,$xml2);
	}
}

function docx_dom_to_array($root)
{
	$result = array();
	if ($root->hasAttributes()) {
		$attrs = $root->attributes;

		foreach ($attrs as $i => $attr) {
			$result[$attr->name] = $attr->value;
		}
	}

	$children = $root->childNodes;
	if ($children->length == 1) {
		$child = $children->item(0);
		if ($child->nodeType == XML_TEXT_NODE) {
			$result['_value'] = $child->nodeValue;
			if (count($result) == 1) {
				return $result['_value'];
			} else {
				return $result;
			}
		}
	}

	$group = array();
	for ($i = 0; $i < $children->length; ++$i) {
		$child = $children->item($i);
		$name = $child->nodeName;
		if ($name == 'w:hyperlink') {
			$name = 'w:r';
			$child->setAttribute('hyperlink', '1');
		}
		if ($name == 'w:tbl') {
			$name = 'w:p';
			$child->setAttribute('tbl', '1');
		}

		if (!isset($result[$name])) {
			$result[$name] = docx_dom_to_array($child);
		} else {
			if (!isset($group[$name])) {
				$tmp = $result[$name];
				$result[$name] = array($tmp);
				$group[$name] = 1;
			}
			$result[$name][] = docx_dom_to_array($child);
		}
	}

	return $result;
}
function docx_each(&$el, $callback, &$param, $key = false)
{
	//Бежим в какие узлы для анализа заходим в какие нет
	$tagelnext = array('w:document','w:body');//Проходные без анализа
	$tagel = array();//Узлы в этом массиве должны быть обработаны
	//<br>
	$tagelnext = array_merge($tagelnext, array('w:p', 'w:r'));
	$tagel = array_merge($tagel, array('w:br'));
	//Картинка
	$tagelnext = array_merge($tagelnext, array('w:r', 'w:drawing'));
	$tagel = array_merge($tagel, array('wp:anchor', 'wp:inline', 'w:pict'));
	//p table h1 h2 h3 h4
	$tagelnext = array_merge($tagelnext, array());
	$tagel = array_merge($tagel, array('w:p'));
	//a
	$tagelnext = array_merge($tagelnext, array());
	$tagel = array_merge($tagel, array('w:r'));
	//b i u
	$tagelnext = array_merge($tagelnext, array());
	$tagel = array_merge($tagel, array('w:p', 'w:r'));
	//Список как абзац
	$tagelnext = array_merge($tagelnext, array());
	$tagel = array_merge($tagel, array('w:p'));//У списка есть [w:pPr][w:numPr]
	//Текст
	$tagelnext = array_merge($tagelnext, array('w:p', 'w:r'));
	$tagel = array_merge($tagel, array('w:t'));

//Таблицы table
	$tagelnext = array_merge($tagelnext, array());
	$tagel = array_merge($tagel, array('w:p', 'w:tr', 'w:tc'));

	$h = '';
	foreach ($el as $k => &$val) {
		if (is_integer($k)) {
			$h .= docx_each($val, $callback, $param, $key);
		} elseif (in_array($k, $tagel)) {
			if (is_array($val) && isset($val[0])) {
				foreach ($val as $kk => &$vv) {
					$h .= call_user_func_array($callback, array(&$vv, $k, &$param, $key));
				}
			} else {
				$h .= call_user_func_array($callback, array(&$val, $k, &$param, $key));
			}
		} elseif (in_array($k, $tagelnext)) {
			$h .= docx_each($val, $callback, $param, $k);
		}
	}
	if ($key === false) {
		//Специально для </ul>
		$h .= call_user_func_array($callback, array(array(), '', &$param, false));
	}

	return $h;
}
function docx_analyse($el, $key, &$param, $keyparent)
{
	$tag = array('','');
	$isli = false;
	$isheading = false;
	$h = '';
	$t = '';
	//Таблицы

	if (is_array($el) && isset($el['tbl']) && $el['tbl'] == '1') {
		$param['istable'] = true;
		$tag = array("<table class='table table-striped'>\n",'</table>');
	} elseif ($key === 'w:tr' && !empty($param['istable'])) {
		$tag = array("<tr>\n",'</tr>');
	//}else if($key==='w:p'&&$param['istable']){
	} elseif ($key === 'w:tc' && !empty($param['istable'])) {
		$tag = array('<td>','</td>');
	} elseif ($key == 'w:pict' && !empty($el['v:shape'])) {
		$rid = $el['v:shape']['v:imagedata']['id'];
		$src = $param['folder'].'word/'.$param['rIds'][$rid];
		$style = $el['v:shape']['style'];
		if (preg_match('/:right/', $style)) {
			$align = 'right';
		} else {
			$align = 'left';
		}
		if (empty($param['images'])) {
			$param['images'] = array();
		}
		$param['images'][] = array('src' => Path::toutf($src));
		//$tag=array('<img align="'.$align.'" src="'.$src.'">','');
		$tag = array('<div style="background-color:gray; color:white; font-weight:normal; padding:5px; font-size:14px; float:'.$align.'">Некорректно<br>добавленная<br>картинка</div>','');
	//Картинки
	} elseif ($keyparent === 'w:drawing') {
		if (empty($param['imgnum'])) $param['imgnum'] = 0;
		$imgnum = ++$param['imgnum'];

		//$origsrc=$el['wp:docPr']['descr'];
		$inline = ($key == 'wp:inline');
		$align = empty($el['wp:positionH']['wp:align']);
		if ($align !== 'left') $align = 'right';

		$width = ceil($el['wp:extent']['cx'] / 8000);
		$height = ceil($el['wp:extent']['cy'] / 8000);

		if ($width > $param['imgmaxwidth']) {
			$width = $param['imgmaxwidth'];
			$height = '';
		}
		$src = $param['folder'].'word/media/image'.$imgnum;
		if (is_file($src.'.jpeg')) {
			$src .= '.jpeg';
		} elseif (is_file($src.'.jpg')) {
			$src .= '.jpg';
		} elseif (is_file($src.'.png')) {
			$src .= '.png';
		} elseif (is_file($src.'.gif')) {
			$src .= '.gif';
		} else {
			$src .= '.wtf';
		}

		if (isset($el['wp:docPr']['title'])) {
			$alt = $el['wp:docPr']['title'];
		} else {
			$alt = '';
		}

		if (empty($param['images'])) $param['images'] = array();
		$param['images'][] = array('src' => Path::toutf($src));

		$src = '/-imager/?src='.Path::toutf($src);
		if ($height) {
			$src .= '&h='.$height;
		}
		if ($width) {
			$src .= '&w='.$width;
		}

		$tag = '<img src="'.$src.'"';

//if($height)$tag.=' height="'.$height.'px"';
		//if($width)$tag.=' width="'.$width.'px"';
		if ($alt) {
			$tag .= ' alt="'.$alt.'"';
		} else {
			$tag .= ' alt=""';
		}
		if (!$inline && $align) {
			$tag .= ' class="img-thumbnail '.$align.'"';
		} else {
			$tag .= ' class="img-thumbnail"';
		}

		$tag .= '>';
		$tag = array($tag,'');

		if (isset($el['wp:docPr']) && isset($el['wp:docPr']['a:hlinkClick'])) {
			//Ссылка на самой картинке
			$r = $el['wp:docPr']['a:hlinkClick']['id'];
			$link = $param['rIds'][$r];
			$tag[0] = '<a href="'.$link.'">'.$tag[0];
			$tag[1] = '</a>';
		}
	//Список
	} elseif ($key === 'w:p' && isset($el['w:pPr']['w:numPr'])) {

		$isli = true;
		$param['isli'] = true;
		//$v = isset($el['w:pPr']['w:numPr']['w:numId'])?$el['w:pPr']['w:numPr']['w:numId']:'';
		$v = $el['rsidRPr'];
		if (!isset($param['isul']) || $param['isul'] !== $v) {
			if (!empty($param['isul']) && $param['isul'] !== $v) $h .= "</ul>\n"; //Раньше был список
			$param['isul'] = $v;
			$h .= "<ul>\n";
			
		}
		$tag = array('<li>','</li>');
	//h1 h2 h3 h4
	} elseif ($key === 'w:p' && !empty($el['rsidR']) && isset($el['w:pPr']['w:pStyle']['val']) && in_array($el['w:pPr']['w:pStyle']['val'], array(1, 2, 3, 4, 5, 6))) {
		$isheading = true;
		$v = $el['w:pPr']['w:pStyle']['val'];
		$tag = array('<h'.$v.'>','</h'.$v.">\n");
	//Абзац
	} elseif ($key === 'w:p' && !empty($el['rsidR'])) {
		$tag = array('<p>',"</p>\n");
	//a
	} elseif ($key === 'w:r' && !empty($el['history'])) {
		$href = $param['rIds'][$el['id']];
		$tag = array('<a href="'.$href.'">','</a>');
	//b i u
	} elseif ($key === 'w:r' && !empty($el['w:rPr']) &&
		   (isset($el['w:rPr']['w:i']) || isset($el['w:rPr']['w:b']) || isset($el['w:rPr']['w:u']))) {
		if (isset($el['w:rPr']['w:i'])) {
			$tag[0] .= '<i>';
		}
		if (isset($el['w:rPr']['w:b'])) {
			$tag[0] .= '<b>';
		}
		if (isset($el['w:rPr']['w:u'])) {
			$tag[0] .= '<u>';
		}

		if (isset($el['w:rPr']['w:u'])) {
			$tag[1] .= '</u>';
		}
		if (isset($el['w:rPr']['w:b'])) {
			$tag[1] .= '</b>';
		}
		if (isset($el['w:rPr']['w:i'])) {
			$tag[1] .= '</i>';
		}
	//<i>
	} elseif ($key === 'w:r' && isset($el['w:rPr']['w:i'])) {
		$tag = array('<i>','</i>');
	//<b>
	} elseif ($key === 'w:r' && isset($el['w:rPr']['w:b'])) {
		$tag = array('<b>','</b>');
	//<br>
	} elseif ($key === 'w:br') {
		$tag = array('<br>','');
	}
	//Список
	if (!empty($param['isul']) && !$isli && $key == 'w:p') {
		//li это абзац и проверяем только на уровне абзацев
		//Есть метка что мы в ul и нет что в li
		$param['isul'] = false;
		$h .= "\n</ul>\n";
	}

//=====================
	if ($key === 'w:t') {
		//Текст
		if (is_string($el)) {
			$t .= $el;
		} else {
			$t .= $el['_value'];
		}

		$h .= $tag[0].$t;
	} else {
		//Вложенность
		$hr = docx_each($el, '\\infrajs\\doc\\docx_analyse', $param, $key);
		if ($tag[0] == '<p>' && preg_match("/\{.*\}/", $hr)) {
			$t = strip_tags($hr);
			if ($t{0} == '{' && $t{strlen($t) - 1} == '}') {
				$t = substr($t, 1, strlen($t) - 2);
				$t = explode(':', $t);
				if (sizeof($t) == 2) {
					$name = $t[0];
					$val = $t[1];

					if ($name == 'div') { //envdiv {div:tadam}
						//чтобы обработать env нужно уже загрузить этот слой к этому времени env обработаны
						//$hr='<script>if(window.infra)infra.when(infrajs,"onshow",function(){ infrajs.envSet("'.$t[1].'",true)});</script>';
						$tag[0] = '<div id="'.$val.'">';
						$tag[1] = '</div>';
						$hr = '';
					}
				}
			}
		}
		$h .= $tag[0];//Открывающий тэг
		//<a>
		if ($isheading && empty($param['heading'])) {
			$param['heading'] = strip_tags($hr);
		}
		if ($key === 'w:r' && !empty($el['history'])) {
			if (empty($param['links'])) $param['links'] = array();
			$href = $param['rIds'][$el['id']];
			$param['links'][] = array('href' => $href,'title' => strip_tags($hr));
		}
		$h .= $hr;
	}
	//=====================

	//Таблицы
	if (is_array($el) && isset($el['tbl']) && $el['tbl'] == '1') {
		$param['istable'] = false;
	//Список
	} elseif ($isli) {
		//Вышли из какого-то li
		$param['isli'] = false;
	} elseif ($isheading) {
		//Вышли из какого-то li
		$isheading = false;
	}

	$h .= $tag[1];//Закрывающий тэг


	return $h;
}
function docx_get($src, $type = 'norm', $re = false)
{
	return Docx::get($src, $type, $re);
}

