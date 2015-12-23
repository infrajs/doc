<?php
namespace infrajs\doc;

use infrajs\cache\Cache;
use infrajs\path\Path;
use infrajs\load\Load;


class Mht
{
	public static function get($src)
	{
		$param=self::parse($src);
		return $param['html'];
	}
	public static function preview($src)
	{
		$param=self::parse($src);

		$data = Load::srcInfo($src);
		$data = Load::nameInfo($data['file']);


		$temphtml = strip_tags($param['html'], '<p><b><strong><i>');
		$temphtml=preg_replace('/\n/', ' ', $temphtml);
		preg_match('/(<p.*>.{1}.*<\/p>)/U', $temphtml, $match);
		if (sizeof($match) > 1) {
			$preview = $match[1];
		} else {
			$preview = $param['html'];
		}
		$preview = preg_replace('/<h1.*<\/h1>/U', '', $preview);
		$preview = preg_replace('/<img.*>/U', '', $preview);
		$preview = preg_replace('/<p.*>\s*<\/p>/iU', '', $preview);
		$preview = preg_replace("/\s+/", ' ', $preview);
		$preview = trim($preview);

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

		if (!empty($param['images'])) {
			$data['images'] = $param['images'];
		}

		$data['preview'] = $preview;


		return $data;
	}
	public static function parse($src)
	{
		$src=Path::theme($src);
		if (!$src) {
			return;
		}

		$args = array($src);
		return Cache::exec(array($src), 'mhtparse', function ($src) {
			$conf = Docx::$conf;
			$imgmaxwidth = $conf['imgmaxwidth'];
			$previewlen = $conf['previewlen'];

			$filename=Path::theme($src);
			$fdata=Load::srcInfo($src);
			if ($fdata['ext']=='php') {
				$data = Load::loadTEXT($filename);
			} else {
				$data = file_get_contents($filename);
			}

			$ans = array();
			if ($fdata['ext']=='mht') {
				$p = explode('/', $filename);
				$fname = array_pop($p);
				$fnameext = $fname;
				//$fname=basename($filename);


				preg_match("/^(\d*)/", $fname, $match);
				$date = $match[0];
				$fname = Path::toutf(preg_replace('/^\d*\s+/', '', $fname));
				$fname = preg_replace('/\.\w{0,4}$/', '', $fname);

				$ar = preg_split('/------=_NextPart_.*/', $data);
				if (sizeof($ar) > 1) {
					//На первом месте идёт информация о ворде...
					unset($ar[0]);
					unset($ar[sizeof($ar) - 1]);
				}
				$ar = array_values($ar);

				$folder = Path::mkdir(Docx::$conf['cache'].md5($src).'/');
				$html = '';
				for ($i = 0, $l = sizeof($ar); $i < $l; ++$i) {
					if (!$ar[$i]) {
						continue;
					}
					$d = preg_split("/\n/", $ar[$i], 6);

					$j = -1;
					do {
						++$j;
					} while (@$d[$j][0] !== 'C' && $j <= 5);

					if ($j >= 5) {
						/*
							не нашли
							Content-Location: file:///C:/0FCF1655/file9909.files/header.htm
							Content-Transfer-Encoding: quoted-printable
							Content-Type: text/html; charset="us-ascii"
						*/
						continue;
					}

					$location = preg_replace('/Content-Location: /', '', $d[$j]);
					$location = trim($location);
					$encoding = preg_replace('/Content-Transfer-Encoding: /', '', $d[$j + 1]);
					$type = preg_replace('/Content-Type: /', '', $d[$j + 2]);
					$content = $d[5];
					$name = basename($location);
					if (preg_match("/text\/html/", $type) || preg_match('/Subject:/', $type)) {
						$html .= $content;
					} else {
						@file_put_contents($folder.$name, base64_decode($content));//Сохраняем картинку или тп...
					}
				}

				if (!$html) {
					$html = '';
				}
				$html = preg_replace("/=\r\n/", '', $html);
				$html = preg_replace("/\s+/", ' ', $html);
				$html = preg_replace("/^.*<body .*>\s*/U", '', $html, 1);
				$html = preg_replace("/\s*<\/body>.*/", '', $html, 1);

				$images = array();

				preg_match_all('/src=3D".*\.files\/(image.+)"/U', $html, $match, PREG_PATTERN_ORDER);

				for ($i = 0, $l = sizeof($match[1]); $i < $l; $i = $i + 2) {
					$min=$match[1][$i + 1];
					if (!$min) {
						$min = $match[1][$i];
					}
					$images[$min] = $match[1][$i];//Каждая следующая картинка есть уменьшенная копия предыдущей оригинального размера
				}


				$html = preg_replace("/<\!--.*-->/U", '', $html);

				$html = preg_replace("/<!\[if !vml\]>/", '', $html);
				$html = preg_replace("/<!\[endif\]>/", '', $html);

				$html = preg_replace('/=3D/', '=', $html);

				$html = preg_replace('/align="right"/', 'align="right" class="right"', $html);
				$html = preg_replace('/align="left"/', 'align="left" class="left"', $html);
				$html = preg_replace('/align=right/', 'align="right" class="right"', $html);
				$html = preg_replace('/align=left/', 'align="left" class="left"', $html);



				$html = Path::toutf($html);//Виндовые файлы хранятся в cp1251

				$folder = Path::toutf($folder);
				$html = preg_replace('/ src=".*\/(.*)"/U', ' src="'.$folder.'${1}"', $html);


				$html = preg_replace('/<span class=SpellE>(.*)<\/span>/U', '${1}', $html);
				$html = preg_replace('/<span lang=.*>(.*)<\/span>/U', '${1}', $html);
				$html = preg_replace('/<span class=GramE>(.*)<\/span>/U', '${1}', $html);
				$html = preg_replace("/<span style='mso.*>(.*)<\/span>/U", '${1}', $html);
				$html = preg_replace("/<span style='mso.*>(.*)<\/span>/U", '${1}', $html);
				$html = preg_replace("/<span style='mso.*>(.*)<\/span>/U", '${1}', $html);
				$html = preg_replace("/<span style='mso.*>(.*)<\/span>/U", '${1}', $html);
				$html = preg_replace('/ class=MsoNormal/U', '', $html);
				$html = preg_replace('/<a name="_.*>(.*)<\/a>/U', '${1}', $html);

		//Приводим к единому виду маркерные списки
				$patern = '/<p class=MsoListParagraphCxSp(\w+) .*>(.*)<\/p>/U';
				$count = 3;
				do {
					preg_match($patern, $html, $match);
					if (sizeof($match) == $count) {
						$pos = strtolower($match[1]);
						$text = $match[2];
						$text = preg_replace('/^.*(<\/span>)+/U', '', $text, 1);
						$text = '<li>'.$text.'</li>';
						if ($pos == 'first') {
							$text = '<ul>'.$text;
						}
						if ($pos == 'last') {
							$text = $text.'</ul>';
						}
						$html = preg_replace($patern, $text, $html, 1);
					} else {
						break;
					}
				} while (sizeof($match) == $count);




				$title = $fname;


				$patern = '/<img(.*)>/U';
				$count = 2;
				do {
					preg_match($patern, $html, $match);
					if (sizeof($match) == $count) {
						$sfind = $match[1];
									//$sfind='<img src="/image.asdf">';
									preg_match("/width=(\d*)/", $sfind, $match2);

						$w = trim($match2[1]);
						preg_match("/height=(\d*)/", $sfind, $match2);
						$h = trim($match2[1]);

						if (!$w || $w > $imgmaxwidth) {
							$w = $imgmaxwidth;
						}

						preg_match('/src="(.*\/)(image.*)"/U', $sfind, $match2);
						$path = trim($match2[1]);
						$small = $match2[2];

						preg_match('/alt="(.*)".*/U', $sfind, $match2);
						$alt = trim(@$match2[1]);
						$alt = html_entity_decode($alt, ENT_QUOTES, 'utf-8');

						preg_match('/align="(.*)".*/U', $sfind, $match2);
						$align = trim($match2[1]);
						$align = html_entity_decode($align, ENT_QUOTES, 'utf-8');

						$big = $images[$small];
						if (!$big) {
							$big = $small;
						}

						$isbig = preg_match('/#/', $alt);
						if ($isbig) {
							$alt = preg_replace('/#/', '', $alt);
						}
						//$i="<IMG title='$alt' src='?-imager/imager.php?w=$w&h=$h&src=".($path.$big)."' align='$align' class='$align' alt='$alt'>";
						$i = "<IMG src='?-imager/imager.php?w=$w&h=$h&src=".($path.$big)."' align='$align' class='$align'>";
						//urlencode решает проблему с ie7 когда иллюстрации с адресом содержащим пробел не показываются
						if ($isbig) {
							$i = "<a target='about:blank' href='?-imager/imager.php?src=".urlencode($path.$big)."'>$i</a>";
						}
						//$i.='<textarea style="width:500px; height:300px">'.$i.'</textarea>';
						$html = preg_replace($patern, $i, $html, 1);
					} else {
						break;
					}
				} while (sizeof($match) == $count);

				$patern = "/###\{(.*)\}###/U";//js код
				do {
					preg_match($patern, $html, $match);

					if (sizeof($match) > 0) {
						$param = $match[1];
						$param = strip_tags($param);
						$param = html_entity_decode($param, ENT_QUOTES, 'utf-8');
						$param = preg_replace('/(‘|’)/', "'", $param);
						$param = preg_replace('/(“|«|»|”)/', '"', $param);
						$html = preg_replace($patern, $param, $html, 1);
					} else {
						break;
					}
				} while (sizeof($match) > 1);

				$patern = "/####.*<table.*>(.*)<\/table>.*####/U";
				do {
					preg_match($patern, $html, $match);
					if (sizeof($match) > 0) {
						$param = $match[1];
						$param = preg_replace('/style=".*"/U', '', $param);
						$param = preg_replace("/style='.*'/U", '', $param);
						$html = preg_replace($patern, '<table class="table table-striped">'.$param.'</table>', $html, 1);
					} else {
						break;
					}
				} while (sizeof($match) > 1);


				$ans['images']=array();
				foreach ($images as $img) {
					$ans['images'][]=array('src'=>$folder.$img);
				}
			} else {
				$html=$data;
				$images = array();
				preg_match_all('/<img.*src="(.*)".*>/U', $html, $match, PREG_PATTERN_ORDER);
				for ($i = 0, $l = sizeof($match[1]); $i < $l; $i ++) {
					$images[] = array('src'=>$match[1][$i]);//Каждая следующая картинка есть уменьшенная копия предыдущей оригинального размера
				}
				$ans['images']=$images;
			}
			$r = preg_match('/<h.*>(.*)<\/h.>/U', $html, $match);
			if ($r) {
				$heading = strip_tags($match[1]);
			} else {
				$heading = false;
			}
			$ans['heading']=$heading;

			preg_match_all('/<a.*href="(.*)".*>(.*)<\/a>/U', $html, $match);
			$links = array();
			foreach ($match[1] as $k => $v) {
				$title = strip_tags($match[2][$k]);
				if (!$title) {
					continue;
				}
				$links[] = array('title' => $title, 'href' => $match[1][$k]);
			}


			$ans['links']=$links;


			$html = trim($html);


			$html=html_entity_decode($html, ENT_COMPAT, 'UTF-8');
			$html=preg_replace('/ /U', '', $html);//bugfix списки в mht порождаются адский символ. в eval-е скрипта недопустим.
			$ans['html'] = $html;
			return $ans;
		}, $args);
	}
}
