<?php

namespace KayStrobach\PhpOffice\View;

use Neos\Flow\Annotations as Flow;
use Neos\Flow\Mvc\Exception\StopActionException;

/**
 * Class AbstractExcelView
 *
 * Based on http://www.networkteam.com/blog/post/verwendung-von-custom-views-in-flow.html
 */
class AbstractExcelView extends \Neos\Flow\Mvc\View\AbstractView
{
    /**
     * define allowed view options
     * @var array
     */
    protected $supportedOptions = array(
        'templateRootPaths' => array(null, 'Path(s) to the template root. If NULL, then $this->options["templateRootPathPattern"] will be used to determine the path', 'array'),
        'partialRootPaths' => array(null, 'Path(s) to the partial root. If NULL, then $this->options["partialRootPathPattern"] will be used to determine the path', 'array'),
        'layoutRootPaths' => array(null, 'Path(s) to the layout root. If NULL, then $this->options["layoutRootPathPattern"] will be used to determine the path', 'array'),
        'writer' => array('Excel2007', 'Defines which writer should be used', 'string'),
        'fileExtension' => array('xlsx', 'file extension for download', 'string'),
    );

    /**
     * @Flow\Inject
     * @var \Neos\Flow\Persistence\PersistenceManagerInterface
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
     * PHP Excel Writer name
     * @var string
     */
    protected $writer = 'Excel2007';

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
     * @var \Neos\Flow\Utility\Environment
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

        $loader = new \KayStrobach\PhpOffice\Utility\PHPOfficeUtility();
        $loader->init('PHPExcel');
        $validLocale = \PHPExcel_Settings::setLocale($this->locale);

        if (!$validLocale) {
            throw new \Exception('No valid locale set');
        }

        $tempFileName = $this->createTempFileFromTemplate();
        $excelFileObject = $this->resetTemplate($tempFileName);

        $this->renderValuesIntoTemplate($excelFileObject);

        header('Content-type: application/ms-excel');
        header('Content-Disposition: attachment;filename="' . $this->pathSegment . $this->getFormatedDateNow() . '.' . $this->getOption('fileExtension') . '"');
        header('Cache-Control: max-age=0');

        $objWriter = \PHPExcel_IOFactory::createWriter($excelFileObject, $this->getOption('writer'));
        $this->configureWriter($objWriter);
        ob_start();
        $objWriter->save('php://output');
        return ob_get_clean();
    }

    protected function createTempFileFromTemplate(): string
    {
        $tempFileName = $this->environment->getPathToTemporaryDirectory() . crc32($this->excelTemplate) . '.xlsx';
        copy($this->excelTemplate, $tempFileName);
        return $tempFileName;
    }

    protected function resetTemplate($file): \PHPExcel
    {
        $excelFileObject = \PHPExcel_IOFactory::load($file);
        $excelFileObject->setActiveSheetIndex(0);
        return $excelFileObject;
    }

    /**
     * @param \PHPExcel $excelFileObject
     * @throws \InvalidArgumentException
     * @throws \PHPExcel_Exception
     */
    protected function renderValuesIntoTemplate(\PHPExcel $excelFileObject): void
    {
        $values = $this->renderValues($excelFileObject, $this->firstRow);
        $rowNumber = $this->firstRow;
        foreach ($values as $row) {
            $activeSheet = $excelFileObject->getActiveSheet();
            /** autoheight of cell to avoid cutting content visibility */
            $activeSheet->getRowDimension($rowNumber)->setRowHeight(-1);
            if (\is_array($row)) {
                $columnNumber = 0;
                foreach ($row as $value) {
                    if (\array_key_exists($columnNumber, $this->columnTypes)) {
                        $activeSheet->setCellValueExplicitByColumnAndRow($columnNumber, $rowNumber, $value, \PHPExcel_Cell_DataType::TYPE_STRING);
                    } else {
                        $activeSheet->setCellValueByColumnAndRow($columnNumber, $rowNumber, $value);
                    }
                    $columnNumber++;
                }
                $rowNumber++;
            } else {
                throw new \InvalidArgumentException('AbstractExcelView, it seems, that your row is not an array ...');
            }
        }
    }

    /**
     * @param \PHPExcel $excelFileObject
     * @param int $firstRow
     * @return array
     */
    public function renderValues(\PHPExcel $excelFileObject, int $firstRow): array
    {
        return array();
    }

    /**
     * @return string
     * @throws \Exception
     */
    protected function getFormatedDateNow(): string
    {
        $date = new \DateTime('now');
        return $date->format('d.m.Y-H_i_s');
    }
    
    protected function configureWriter(\PHPExcel_Writer_IWriter $writer)
    {
        return $writer;
    }
}
