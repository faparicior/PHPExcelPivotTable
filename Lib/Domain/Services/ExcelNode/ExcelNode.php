<?php

namespace Faparicior\ExcelPivotTable\Domain\Services\ExcelNode;

interface ExcelNode
{
    public static function checkValue($value);
    public static function valueToExcelNode($value, \DOMElement $node, \DOMDocument $document);
}
