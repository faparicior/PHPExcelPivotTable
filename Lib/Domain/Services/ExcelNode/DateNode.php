<?php

namespace Faparicior\ExcelPivotTable\Domain\Services\ExcelNode;

final class DateNode implements ExcelNode
{
    const DATE_FORMAT = 'd/m/Y';

    protected function __construct()
    {

    }

    /**
     * @param $value
     *
     * @return bool
     */
    public static function checkValue($value)
    {
        $date = \DateTime::createFromFormat(self::DATE_FORMAT, $value);

        if ($date) {
            return $date->format(self::DATE_FORMAT) == $value;
        }

        return false;
    }

    /**
     * @param $value
     * @param \DOMElement $node
     * @param \DOMDocument $document
     *
     * @return \DOMNodeList
     */
    public static function valueToExcelNode($value, \DOMElement $node, \DOMDocument $document)
    {
        $cell = $node->getElementsByTagName('v');

        $greg_start = gregoriantojd(1, 1, 1900);
        $date = \DateTime::createFromFormat(self::DATE_FORMAT, $value);

        $cell[0]->nodeValue = gregoriantojd(
            $date->format('n'),
            $date->format('j'),
            $date->format('Y')
            ) - $greg_start + 2;

        return $cell;
    }
}
