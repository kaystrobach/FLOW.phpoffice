<?php

namespace KayStrobach\PhpOffice\View;

/**
 * Class AbstractExcelView
 *
 * Based on http://www.networkteam.com/blog/post/verwendung-von-custom-views-in-flow.html
 */
class ArrayExcelView extends AbstractExcelView
{
    /**
     * Renders the view
     *
     * @param int $firstRow
     * @return array The rendered array values
     * @api
     */
    public function renderValues(int $firstRow): array
    {
        $values = array();
        /** @var array $line */
        foreach ($this->variables['values'] as $line) {
            if (\is_array($line)) {
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
    protected function valueToString($value): string
    {
        if (\is_string($value)) {
            return $value;
        }
        if (\is_int($value)) {
            return (string) $value;
        }
        if (\is_float($value)) {
            return (string) $value;
        }
        return \gettype($value) . ' - can´t be casted to string';
    }
}
