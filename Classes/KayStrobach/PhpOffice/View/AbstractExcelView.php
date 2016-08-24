<?php

namespace KayStrobach\PhpOffice\View;

use TYPO3\Flow\Annotations as Flow;
use TYPO3\Flow\Mvc\Exception\StopActionException;

/**
 * Class AbstractExcelView
 *
 * Basiert auf http://www.networkteam.com/blog/post/verwendung-von-custom-views-in-flow.html
 *
 * @package SBS\LaPo\View
 */
class AbstractExcelView extends \TYPO3\Flow\Mvc\View\AbstractView
{
    /**
     * define allowed view options
     * @var array
     */
    protected $supportedOptions = array(
        'templateRootPaths' => array(null, 'Path(s) to the template root. If NULL, then $this->options["templateRootPathPattern"] will be used to determine the path', 'array'),
        'partialRootPaths' => array(null, 'Path(s) to the partial root. If NULL, then $this->options["partialRootPathPattern"] will be used to determine the path', 'array'),
        'layoutRootPaths' => array(null, 'Path(s) to the layout root. If NULL, then $this->options["layoutRootPathPattern"] will be used to determine the path', 'array'),
    );

    /**
     * @Flow\Inject
     * @var \TYPO3\Flow\Persistence\PersistenceManagerInterface
     */
    protected $persistenceManager;

    /**
     * @var null
     */
    protected $excelTemplate = null;

    /**
     * @var string
     */
    protected $locale = 'de_de';

    /**
     * @var string
     */
    protected $pathSegment = 'export-';

    /**
     * @var int
     */
    protected $firstRow = 2;

    /**
     * @var array
     */
    protected $columnTypes = array();

    /**
     * @Flow\Inject
     * @var \TYPO3\Flow\Utility\Environment
     */
    protected $environment;

    /**
     * @param string or resource $path
     */
    public function setTemplatePath($path)
    {
        $this->excelTemplate = $path;
    }

    /**
     * Renders the view
     *
     * @throws \Exception
     * @return string The rendered view
     * @api
     */
    public function render()
    {
        /**
         * wird aus dem Paket Packages/Libraries/kaystrobach/phpoffice geladen
         *
         * @var \PHPExcel $PHPExcel
         */

        $loader = new \KayStrobach\PhpOffice\Utility\PHPOfficeUtility('PHPExcel');
        $loader->init('PHPExcel');
        $validlocale = \PHPExcel_Settings::setLocale($this->locale);

        if (!$validlocale) {
            throw new \Exception('No valid locale set');
        }

        $tempFileName = $this->environment->getPathToTemporaryDirectory() . crc32($this->excelTemplate) . '.xlsx';
        copy($this->excelTemplate, $tempFileName);

        $excelFileObject = \PHPExcel_IOFactory::load($tempFileName);
        $excelFileObject->setActiveSheetIndex(0);
        $rowNumber = $this->firstRow;
        $values = $this->renderValues($excelFileObject, $this->firstRow);
        foreach ($values as $row) {
            /** höhe der zeile automatisch, damit der inhalt nicht abgeschnitten wird */
            $excelFileObject->getActiveSheet()->getRowDimension($rowNumber)->setRowHeight(-1);
            if (is_array($row)) {
                $columnNumber = 0;
                foreach ($row as $value) {
                    if (array_key_exists($columnNumber, $this->columnTypes)) {
                        $excelFileObject->getActiveSheet()->setCellValueExplicitByColumnAndRow($columnNumber, $rowNumber, $value, \PHPExcel_Cell_DataType::TYPE_STRING);
                    } else {
                        $excelFileObject->getActiveSheet()->setCellValueByColumnAndRow($columnNumber, $rowNumber, $value);
                    }
                    $columnNumber++;
                }
                $rowNumber++;
            } else {
                throw new \Exception('AbstractExcelView, it seems, that your row is not an array ...');
            }
        }

        header('Content-type: application/ms-excel');
        header('Content-Disposition: attachment;filename="' . $this->pathSegment . $this->getFormatedDateNow() . '.xlsx"');
        header('Cache-Control: max-age=0');

        $objWriter = \PHPExcel_IOFactory::createWriter($excelFileObject, 'Excel2007');
        ob_start();
        $objWriter->save('php://output');
        echo ob_get_clean();
        throw new StopActionException();

    }

    /**
     * @param \PHPExcel $excelFileObject
     * @param $firstRow
     * @return array
     */
    public function renderValues(\PHPExcel $excelFileObject, $firstRow)
    {
        return array();
    }

    protected function getFormatedDateNow()
    {
        $date = new \DateTime('now');
        return $date->format('d.m.Y-H_i_s');
    }
}
