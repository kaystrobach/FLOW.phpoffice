<?php

namespace KayStrobach\PhpOffice\View;

use Neos\Flow\Annotations as Flow;

/**
 * Class StudentExcelView
 *
 * Basiert auf http://www.networkteam.com/blog/post/verwendung-von-custom-views-in-flow.html
 *
 * @package SBS\LaPo\View
 */
class ArrayExcelView extends AbstractExcelView
{
    /**
     * Renders the view
     *
     * @return string The rendered view
     * @api
     */
    public function renderValues(\PHPExcel $excelFileObject, $firstRow)
    {
        $values = array();
        /** @var array $line */
        foreach ($this->variables['values'] as $line) {
            if (is_array($line)) {
                $lineValues = array();
                foreach($line as $cell) {
                    $lineValues[] = $this->valueToString($cell);
                }
                $values[] = $lineValues;
            } else {
                $values[] = [
                    $this->valueToString($line)
                ];
            }
        }
        return $values;
    }

    /**
     * Convert any type of object to a string
     *
     * @param $value
     * @return string
     */
    protected function valueToString($value) {
        if (is_string($value)) {
            return $value;
        } elseif (is_integer($value)) {
            return (string) $value;
        } elseif (is_float($value)) {
            return (string) $value;
        } else {
            return gettype($value) . ' - can´t be casted to string';
        }
    }
}
