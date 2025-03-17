<?php

namespace KayStrobach\PhpOffice\View;

use Neos\Flow\Annotations as Flow;
use Neos\Flow\Http\Component\SetHeaderComponent;
use Neos\Flow\Mvc\Exception\StopActionException;
use Neos\Flow\Mvc\View\AbstractView;
use Neos\Flow\Persistence\PersistenceManagerInterface;
use Neos\Flow\Utility\Environment;
use PhpOffice\PhpSpreadsheet\Cell\CellAddress;
use PhpOffice\PhpSpreadsheet\Cell\DataType;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Settings;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\IWriter;
use PhpOffice\PhpSpreadsheet\Exception;

/**
 * Class AbstractExcelView
 *
 * Based on http://www.networkteam.com/blog/post/verwendung-von-custom-views-in-flow.html
 */
class AbstractExcelView extends AbstractView
{
    /**
     * define allowed view options
     * @var array
     */
    protected $supportedOptions = array(
        'templateRootPaths' => [null, 'Path(s) to the template root. If NULL, then $this->options["templateRootPathPattern"] will be used to determine the path', 'array'],
        'partialRootPaths' => [null, 'Path(s) to the partial root. If NULL, then $this->options["partialRootPathPattern"] will be used to determine the path', 'array'],
        'layoutRootPaths' => [null, 'Path(s) to the layout root. If NULL, then $this->options["layoutRootPathPattern"] will be used to determine the path', 'array'],
        'writer' => ['Xlsx', 'Defines which writer should be used', 'string'],
        'fileExtension' => ['xlsx', 'file extension for download', 'string'],
        'exportAdditionalFileName' => [null, 'additional file name component, will be set automatically, if null', 'string']
    );

    /**
     * @Flow\Inject
     * @var PersistenceManagerInterface
     */
    protected $persistenceManager;

    /**
     * name of template file
     * @var string
     */
    protected $excelTemplate;

    /**
     * @var string
     */
    protected $locale = 'de_de';

    /**
     * PHP Excel Writer name
     * @var string
     */
    protected $writer = 'Xlsx';

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
     * @var Environment
     */
    protected $environment;

    /**
     * @var Spreadsheet
     */
    protected $spreadsheet;

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
    public function render(): string
    {
        $validLocale = Settings::setLocale($this->locale);
        if (!$validLocale) {
            throw new \Exception('No valid locale set');
        }

        $tempFileName = $this->createTempFileFromTemplate();
        $this->spreadsheet = $this->resetTemplate($tempFileName);

        $this->renderValuesIntoTemplate();

        $additionalName = $this->getOption('exportAdditionalFileName') ?? ($this->controllerContext->getRequest()->getControllerName() . '-' . $this->controllerContext->getRequest()->getControllerActionName());
        $downloadFilename = $this->pathSegment . '-' . $additionalName . '-' . $this->getFormatedDateNow() . '.' . $this->getOption('fileExtension');

        $response = $this->controllerContext->getResponse();
        $response->setContentType('application/ms-excel');
        $response->setComponentParameter(
            SetHeaderComponent::class,
            'Content-Disposition',
            'attachment;filename="' . $downloadFilename
        );
        $response->setComponentParameter(
            SetHeaderComponent::class,
            'Cache-Control',
            'max-age=0'
        );

        $objWriter = IOFactory::createWriter($this->spreadsheet, $this->getOption('writer'));
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

    protected function resetTemplate($file): Spreadsheet
    {
        $excelFileObject = IOFactory::load($file);
        $excelFileObject->setActiveSheetIndex(0);
        return $excelFileObject;
    }

    /**
     * @throws Exception
     */
    protected function renderValuesIntoTemplate()
    {
        $values = $this->renderValues($this->spreadsheet, $this->firstRow);
        $rowNumber = $this->firstRow;
        $activeSheet = $this->spreadsheet->getActiveSheet();
        foreach ($values as $row) {
            /** autoheight of cell to avoid cutting content visibility */
            $activeSheet->getRowDimension($rowNumber)->setRowHeight(-1);
            if (\is_array($row)) {
                $columnNumber = 1; // as the next called functions are now converted to A1 etc. we have to start with 1
                foreach ($row as $value) {
                    $cell = $activeSheet->getCell([$columnNumber, $rowNumber]);
                    if (\array_key_exists($columnNumber, $this->columnTypes)) {
                        $cell->setValueExplicit($value, $this->columnTypes[$columnNumber]);
                    } else {
                        $cell->setValueExplicit($value, DataType::TYPE_STRING);
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
     * @param Spreadsheet $sheet
     * @param int $firstRow
     * @return array
     */
    public function renderValues(Spreadsheet $sheet, int $firstRow): array
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

    protected function configureWriter(IWriter $writer): IWriter
    {
        return $writer;
    }
}
