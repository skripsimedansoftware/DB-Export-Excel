<?php
require_once('vendor/autoload.php');

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

function nice_date($bad_date = '', $format = FALSE)
{
	if (empty($bad_date))
	{
		return 'Unknown';
	}
	elseif (empty($format))
	{
		$format = 'U';
	}

	// Date like: YYYYMM
	if (preg_match('/^\d{6}$/i', $bad_date))
	{
		if (in_array(substr($bad_date, 0, 2), array('19', '20')))
		{
			$year  = substr($bad_date, 0, 4);
			$month = substr($bad_date, 4, 2);
		}
		else
		{
			$month  = substr($bad_date, 0, 2);
			$year   = substr($bad_date, 2, 4);
		}

		return date($format, strtotime($year.'-'.$month.'-01'));
	}

	// Date Like: YYYYMMDD
	if (preg_match('/^\d{8}$/i', $bad_date, $matches))
	{
		return DateTime::createFromFormat('Ymd', $bad_date)->format($format);
	}

	// Date Like: MM-DD-YYYY __or__ M-D-YYYY (or anything in between)
	if (preg_match('/^(\d{1,2})-(\d{1,2})-(\d{4})$/i', $bad_date, $matches))
	{
		return date($format, strtotime($matches[3].'-'.$matches[1].'-'.$matches[2]));
	}

	// Any other kind of string, when converted into UNIX time,
	// produces "0 seconds after epoc..." is probably bad...
	// return "Invalid Date".
	if (date('U', strtotime($bad_date)) === '0')
	{
		return 'Invalid Date';
	}

	// It's probably a valid-ish date format already
	return date($format, strtotime($bad_date));
}

$pdo = new PDO('mysql:host=localhost;dbname=skripsi-max-miner','root','');

$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();

$headers = array('No', 'ID Transaksi', 'Items', 'Harga Total', 'Tanggal');
$range_A_Z = range('A', 'Z');
$data = $pdo->query('SELECT * FROM `order`', PDO::FETCH_ASSOC)->fetchAll();

for ($rows = 1; $rows < (count($data)+2); $rows++)
{
	for ($i = 0; $i < count($headers); $i++)
	{
		if ($rows == 1)
		{
			$sheet->setCellValue($range_A_Z[$i].$rows, $headers[$i]);
		}
		else
		{
			$cart = $pdo->query('SELECT * FROM `cart` WHERE `order_id` = '.$data[$rows-2][array_keys($data[$rows-2])[0]].' ', PDO::FETCH_ASSOC)->fetchAll();
			$cart = array_map(function($cart) {
				return $cart['name'];
			}, $cart);

			if ($i == 0)
			{
				$sheet->setCellValue($range_A_Z[$i].$rows, ($rows-1)); // numbering
			}
			elseif ($i == 2)
			{
				$sheet->setCellValue($range_A_Z[$i].$rows, implode(', ', $cart)); // custom column for item
			}
			elseif ($i == 4)
			{
				$sheet->setCellValue($range_A_Z[$i].$rows, nice_date($data[$rows-2][array_keys($data[$rows-2])[$i]], 'd/m/Y')); // custom column for date
			}
			else
			{
				$sheet->setCellValue($range_A_Z[$i].$rows, $data[$rows-2][array_keys($data[$rows-2])[$i]]);
			}
		}
	}
}

$writer = new Xlsx($spreadsheet);
$writer->save('export.xlsx');
?>