# ci_office_excel
This is a library for codeigniter to read and generate excel document.

### How to use
>First
```
$this->load->library('excel');
```

>Read

```PHP
$objLoad = PHPExcel_IOFactory::load('example.xls');
$array_data = array();
foreach ($objLoad->getWorksheetIterator() as $worksheet) {
	$worksheetTitle     = $worksheet->getTitle();
	$highestRow         = $worksheet->getHighestRow(); // e.g. 10
	$highestColumn      = $worksheet->getHighestColumn(); // e.g 'F'
	$highestColumnIndex = PHPExcel_Cell::columnIndexFromString($highestColumn);

	for ($row = 2; $row <= $highestRow; ++ $row) {
		for ($col = 0; $col < $highestColumnIndex; ++ $col) {
		  $cell = $worksheet->getCellByColumnAndRow($col, $row);
		  $val = $cell->getValue();
		  $array_data[$row][$col] = $val;
		}
	}
}
```


>Generate
```PHP
$i=2;
$this->load->library('excel');
$filename='example.xls'; //save our workbook as this file name
header('Content-Type: application/vnd.ms-excel'); //mime type
header('Content-Disposition: attachment;filename="'.$filename.'"'); //tell browser what's the file name
header('Cache-Control: max-age=0'); //no cache

// Set Header on table
$this->excel->setActiveSheetIndex(0);
$this->excel->getActiveSheet()->setTitle('Phone Book');
$this->excel->getActiveSheet()->setCellValue('A1', 'id');
$this->excel->getActiveSheet()->getStyle('A1')->getFont()->setBold(true);
$this->excel->getActiveSheet()->getStyle('A1')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

$this->excel->getActiveSheet()->setCellValue('B1', 'name');
$this->excel->getActiveSheet()->getStyle('B1')->getFont()->setBold(true);
$this->excel->getActiveSheet()->getStyle('B1')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

$this->excel->getActiveSheet()->setCellValue('C1', 'number');
$this->excel->getActiveSheet()->getStyle('C1')->getFont()->setBold(true);
$this->excel->getActiveSheet()->getStyle('C1')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

// write on table
$this->excel->getActiveSheet()->setCellValue('A2', '1');
$this->excel->getActiveSheet()->setCellValue('B2', 'George Lovato');
$this->excel->getActiveSheet()->setCellValue('C2', '21000187');

$this->excel->getProperties()->setCreator("author");
$this->excel->getProperties()->setLastModifiedBy("author");
$this->excel->getProperties()->setTitle($filename);
$this->excel->getProperties()->setSubject("Phone Book");
$this->excel->getProperties()->setDescription(base_url());
$this->excel->getProperties()->setCategory("Phone Book");

$objWriter = PHPExcel_IOFactory::createWriter($this->excel, 'Excel5');
$objWriter->save('php://output');
exit;
```
