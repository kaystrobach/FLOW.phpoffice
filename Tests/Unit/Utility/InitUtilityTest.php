<?php

namespace KayStrobach\PhpOffice\Utility;


class PhpOfficeUtilityTest  {

	protected $supportedLibraries = array(
		'PHPExcel',
		'PHPPowerPoint',
		'PHPProject',
		'PHPVisio',
		'PHPWord'
	);

	/**
	 * Test if valid Objects are returned
	 *
	 * @test
	 */
	public function initTest() {
		/**
		 * @var Tx_Phpoffice_Utility_InitUtility $initObject
		 */
		$initObject = t3lib_div::makeInstance('Tx_Phpoffice_Utility_InitUtility');
		foreach($this->supportedLibraries as $library) {
			$object = $initObject->init($library);
			if($library === NULL) {
				$this->fail('Got no object of type: ' . $library);
			} else {
				$classname = get_class($object);
				$this->assertSame($library, $classname, 'Wrong classname returned, expected ' . $library . ' got ' . $classname);
			}
		}
	}

}
?>