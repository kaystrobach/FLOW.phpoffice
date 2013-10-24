<?php

namespace KayStrobach\PhpOffice\Utility;

class PHPOfficeUtility {
	/**
	 * contains the already initilized libaries
	 */
	protected $initialized = array();

	/**
	 *
	 *
	 * @param null $library
	 * @return null
	 * @throws Exception
	 */
	public function init($library = NULL) {
		$path         = __DIR__ . '/../../../../Resources/Private/PHP/' . $library . '/Classes/' . $library . '.php';
		$fallbackPath = __DIR__ . '/../../../../Resources/Private/PHP/' . $library . '/src/' . $library . '.php';
		if(!in_array($library, $this->initialized)) {
			if(file_exists($path)) {
				include_once($path);
				if(class_exists($library)) {
					return new $library();
				}
			} elseif(file_exists($fallbackPath)) {
				include_once($fallbackPath);
				if(class_exists($library)) {
					return new $library();
				}
			} else {
				throw new Exception('Can´t find library "' . $library . '" in "' . $path . '"');
			}
		}
		return NULL;
	}
}