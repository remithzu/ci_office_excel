<?php
defined('BASEPATH') OR exit('No direct script access allowed');
/**
* 
*/
require_once APPPATH."/libraries/PHPExcel.php";
class Excel extends PHPExcel
{
	function __construct()
	{
		parent::__construct();
	}
}

/* End of file Excel.php */
/* Location: ./application/libraries/Excel.php */
