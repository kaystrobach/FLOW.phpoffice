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
	public function init($library = NULL)
    {
		$path         = FLOW_PATH_PACKAGES . 'Libraries/phpoffice/' . strtolower($library) . '/Classes/' . $library . '.php';
		$fallbackPath = FLOW_PATH_PACKAGES . 'Libraries/phpoffice/' . strtolower($library) . '/src/' . $library . '.php';
		$classname    = '\\' . $library;
		if (!class_exists($classname)) {
			if (!in_array($library, $this->initialized)) {
				if (file_exists($path)) {
					include_once($path);
				} elseif (file_exists($fallbackPath)) {
					include_once($fallbackPath);
				} else {
					throw new \Exception('Can´t find library "' . $library . '" in "' . $path . '"');
				}
			}
		}
		if (class_exists($classname)) {
			return new $classname();
		}
		return NULL;
	}
}
