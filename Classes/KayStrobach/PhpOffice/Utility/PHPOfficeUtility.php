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
	 * @throws \Exception
	 */
	public function init($library = NULL) {
		$path         = FLOW_PATH_PACKAGES . 'Libraries/phpoffice/' . $library . '/Classes/' . $library . '.php';
		$fallbackPath = FLOW_PATH_PACKAGES . 'Libraries/phpoffice/' . $library . '/src/' . $library . '.php';
		$classname    = '\\' . $library;
		if(!in_array($library, $this->initialized)) {
			if(file_exists($path)) {
				include_once($path);
				if(class_exists($classname)) {
					return new $classname();
				}
			} elseif(file_exists($fallbackPath)) {
				include_once($fallbackPath);
				if(class_exists($classname)) {
					return new $classname();
				}
			} else {
				throw new \Exception('Can´t find library "' . $library . '" in "' . $path . '"');
			}
		}
		return NULL;
	}
}
